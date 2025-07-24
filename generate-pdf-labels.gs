function generateColumnLabelsStyledPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form responses 1');
  const data = formSheet.getDataRange().getValues();

  const headers = data[0];
  const entries = data.slice(1);

  const cityIndex = headers.indexOf('City');
  const nameIndex = headers.indexOf('Full Name');
  const addressIndex = headers.indexOf('House Number and Street Name');
  const postcodeIndex = headers.indexOf('Postcode');
  const mobileIndex = headers.indexOf('Phone Number');

  // Identify dish columns and prices
  const dishColumns = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header.includes("£")) {
      const match = header.match(/^(.*?)\s*[-–]\s*£([\d.]+)/);
      if (match) {
        dishColumns.push({
          name: match[1].trim(),
          price: parseFloat(match[2]),
          index: i
        });
      }
    }
  }

  // Group entries by city
  const cityGroups = {};
  for (let row of entries) {
    const city = row[cityIndex];
    if (!city) continue;
    if (!cityGroups[city]) cityGroups[city] = [];
    cityGroups[city].push(row);
  }

  // Generate a document per city
  Object.keys(cityGroups).forEach(city => {
    const doc = DocumentApp.create(`Labels_MultiColumn_${city}`);
    const body = doc.getBody();
    body.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FONT_SIZE]: 10
    });

    const orders = cityGroups[city];
    let table = body.appendTable();
    let rowCount = 0;

    for (let i = 0; i < orders.length; i += 2) {
      const row = table.appendTableRow();

      for (let j = 0; j < 2; j++) {
        const index = i + j;
        const cell = row.appendTableCell();
        const para = cell.appendParagraph('');

        if (index < orders.length) {
          const order = orders[index];
          const name = order[nameIndex] || 'N/A';
          const address = order[addressIndex] || 'N/A';
          const postcode = formatUKPostcode(order[postcodeIndex]);
          const contactNo = order[mobileIndex] || 'N/A';

          let totalCost = 0;

          // Build label content
          addBoldLine(para, 'Name: ', name);
          para.appendText('\n');
          para.appendText('Order Details:\n').setBold(true);
          
          dishColumns.forEach(dish => {
            const qty = Number(order[dish.index]) || 0;
            if (qty > 0) {
              para.appendText(`${dish.name} x${qty}\n`);
              totalCost += qty * dish.price;
            }
          });

          addBoldLine(para, 'Total Cost: ', `£${totalCost.toFixed(2)}`);
          para.appendText('\n');
          addBoldLine(para, 'Address: ', `${address}, ${city}, ${postcode}`);
          addBoldLine(para, 'Contact Number: ', contactNo);
        }

        cell.setPaddingTop(6);
        cell.setPaddingBottom(6);
        cell.setPaddingLeft(10);
        cell.setPaddingRight(10);
      }

      rowCount++;

      // After every 3 rows (6 labels), start new table and page
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

// Helper function for bold header lines
function addBoldLine(paragraph, label, value) {
  if (!value || value.toString().trim() === '') value = 'N/A';
  paragraph.appendText(label).setBold(true);
  paragraph.appendText(value + '\n');
}

// Format UK postcode with space before last 3 characters
function formatUKPostcode(postcode) {
  if (!postcode) return 'N/A';
  postcode = postcode.replace(/\s+/g, '').toUpperCase();
  if (postcode.length > 3) {
    return postcode.slice(0, -3) + ' ' + postcode.slice(-3);
  }
  return postcode;
}
