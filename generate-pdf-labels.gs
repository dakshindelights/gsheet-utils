function generateColumnLabelsStyledPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("No data rows found.");

  const headers = data[0].map(flattenHeader);
  const rows = data.slice(1);

  // Core customer columns (look them up flexibly)
  const nameIndex     = findCol(headers, "Name");
  const phoneIndex    = findCol(headers, "Phone Number");
  const addressIndex  = findCol(headers, "House Number and Street Name");
  const cityIndex     = findCol(headers, "City");
  const postcodeIndex = findCol(headers, "Postcode");

  // Dish columns = any header that has a £ price
  const dishColumns = [];
  headers.forEach((h, i) => {
    const priceMatch = h.match(/£\s*([\d]+(?:\.\d+)?)/);
    if (priceMatch) {
      dishColumns.push({
        name: h.replace(/£.*/,"").trim(),
        price: parseFloat(priceMatch[1]),
        index: i
      });
    }
  });

  // Group rows by city
  const cityGroups = {};
  rows.forEach(r => {
    const city = safeCell(r, cityIndex);
    if (!city) return;
    if (!cityGroups[city]) cityGroups[city] = [];
    cityGroups[city].push(r);
  });

  // Generate doc per city
  Object.keys(cityGroups).forEach(city => {
    const doc = DocumentApp.create(`Labels_MultiColumn_${city}`);
    const body = doc.getBody();
    body.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: "Arial",
      [DocumentApp.Attribute.FONT_SIZE]: 10
    });

    let table = body.appendTable();
    let rowCount = 0;
    const orders = cityGroups[city];

    for (let i = 0; i < orders.length; i += 2) {
      const row = table.appendTableRow();

      for (let j = 0; j < 2; j++) {
        const idx = i + j;
        const cell = row.appendTableCell();
        const para = cell.appendParagraph("");

        if (idx < orders.length) {
          const order = orders[idx];

          const name      = safeCell(order, nameIndex)     || "N/A";
          const address   = safeCell(order, addressIndex)  || "N/A";
          const postcode  = formatUKPostcode(safeCell(order, postcodeIndex));
          const phone     = safeCell(order, phoneIndex)    || "N/A";

          let total = 0;
          addBoldLine(para, "Name: ", name);
          para.appendText("\n");
          para.appendText("Order Details:\n").setBold(true);

          dishColumns.forEach(d => {
            const qty = Number(order[d.index]) || 0;
            if (qty > 0) {
              para.appendText(`${d.name} x${qty}\n`);
              total += qty * d.price;
            }
          });

          addBoldLine(para, "Total Cost: ", `£${total.toFixed(2)}`);
          para.appendText("\n");
          addBoldLine(para, "Address: ", `${address}, ${city}, ${postcode}`);
          addBoldLine(para, "Contact Number: ", phone);
        }
      }

      rowCount++;
      if (rowCount % 3 === 0 && i + 2 < orders.length) {
        body.appendPageBreak();
        table = body.appendTable();
      }
    }

    doc.saveAndClose();
    const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    DriveApp.getRootFolder().createFile(pdf.setName(`Labels_${city}.pdf`));
  });
}

/* === Helpers === */

function flattenHeader(h) {
  return String(h || "")
    .replace(/^"+|"+$/g, "")   // strip quotes
    .replace(/\n+/g, " ")      // replace newlines
    .replace(/\s+/g, " ")      // collapse spaces
    .trim();
}

function findCol(headers, wanted) {
  const target = wanted.toLowerCase();
  return headers.findIndex(h => h.toLowerCase().includes(target));
}

function safeCell(row, idx) {
  return (idx == null || idx < 0) ? "" : row[idx];
}

function addBoldLine(para, label, value) {
  const v = (value == null || String(value).trim() === "") ? "N/A" : String(value);
  para.appendText(label).setBold(true);
  para.appendText(v + "\n");
}

function formatUKPostcode(postcode) {
  if (!postcode) return "N/A";
  const cleaned = String(postcode).replace(/\s+/g, "").toUpperCase();
  return cleaned.length > 3 ? cleaned.slice(0, -3) + " " + cleaned.slice(-3) : cleaned;
}
