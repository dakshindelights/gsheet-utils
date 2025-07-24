function generateOrderSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form responses 1');
  const data = formSheet.getDataRange().getValues();
  const headers = data[0];
  const entries = data.slice(1);

  const dishColumns = [];

  // Identify dish columns based on headers containing "£"
  for (let i = 0; i < headers.length; i++) {
    const rawHeader = headers[i];
    const sanitizedHeader = String(rawHeader).replace(/\n/g, ' ').trim();
    const match = sanitizedHeader.match(/^(.*?)\s*[-–]\s*£\s*([\d.]+)/);
    if (match) {
      const dishName = match[1].trim();
      const unitPrice = parseFloat(match[2]);
      dishColumns.push({
        name: dishName,
        price: unitPrice,
        index: i
      });
    }
  }

  // Aggregate order quantities and total sales
  const summary = {};

  dishColumns.forEach(dish => {
    let qty = 0;
    entries.forEach(row => {
      qty += Number(row[dish.index]) || 0;
    });

    const totalSale = qty * dish.price;

    summary[dish.name] = {
      quantity: qty,
      price: dish.price,
      sale: totalSale
    };
  });

  // Create or clear summary sheet
  const sheetName = 'Order Summary';
  let summarySheet = ss.getSheetByName(sheetName);
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet(sheetName);
  }

  // Write order summary header
  let row = 1;
  summarySheet.getRange(row++, 1, 1, 4).setValues([['Dish Name', 'Quantity', 'Unit Price', 'Total Sale']]);

  // Write dish summary rows
  Object.entries(summary).forEach(([dish, data]) => {
    summarySheet.getRange(row++, 1, 1, 4).setValues([[
      dish,
      data.quantity,
      `£${data.price.toFixed(2)}`,
      `£${data.sale.toFixed(2)}`
    ]]);
  });

  // Calculate and write overall total
  const overallTotal = Object.values(summary).reduce((sum, item) => sum + item.sale, 0);
  row += 2;
  summarySheet.getRange(row++, 1).setValue(`Overall Total Sale: £${overallTotal.toFixed(2)}`);
}
