function generateOrderSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form responses 1');
  const data = formSheet.getDataRange().getValues();
  const headers = data[0];
  const entries = data.slice(1);

  // Dish identification
  const dishColumns = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header.includes("£")) {
      const match = header.match(/^(.*?)\s*[-–]\s*£([\d.]+)/);
      if (match) {
        const dishName = match[1].trim();
        const unitPrice = parseFloat(match[2]);
        let section = 'Other';
        if (dishName.toLowerCase().includes('Chicken Majestic') || dishName.toLowerCase().includes('PERUGU VADA')) section = 'Starters';
        else if (dishName.toLowerCase().includes('Vegetable Keema Pulao') || dishName.toLowerCase().includes('Chicken Pulao')) section = 'Pulaos';
        dishColumns.push({
          name: dishName,
          price: unitPrice,
          index: i,
          section: section
        });
      }
    }
  }

  // Aggregate dish data
  const summary = {};
  const sectionSales = {};

  dishColumns.forEach(dish => {
    let qty = 0;
    entries.forEach(row => {
      qty += Number(row[dish.index]) || 0;
    });

    const totalSale = qty * dish.price;

    summary[dish.name] = {
      quantity: qty,
      price: dish.price,
      sale: totalSale,
      section: dish.section
    };

    if (!sectionSales[dish.section]) sectionSales[dish.section] = 0;
    sectionSales[dish.section] += totalSale;
  });

  // Create summary sheet
  const sheetName = 'Order Summary';
  let summarySheet = ss.getSheetByName(sheetName);
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet(sheetName);
  }

  let row = 1;
  summarySheet.getRange(row++, 1, 1, 4).setValues([['Dish Name', 'Quantity', 'Unit Price', 'Total Sale']]);

  Object.entries(summary).forEach(([dish, data]) => {
    summarySheet.getRange(row++, 1, 1, 4).setValues([[
      dish,
      data.quantity,
      `£${data.price.toFixed(2)}`,
      `£${data.sale.toFixed(2)}`
    ]]);
  });

  row += 2;
  summarySheet.getRange(row++, 1).setValue('Section-wise Totals');
  summarySheet.getRange(row++, 1, 1, 2).setValues([['Section', 'Total Sale']]);

  Object.entries(sectionSales).forEach(([section, sale]) => {
    summarySheet.getRange(row++, 1, 1, 2).setValues([[section, `£${sale.toFixed(2)}`]]);
  });

  row += 2;
  const overallTotal = Object.values(sectionSales).reduce((a, b) => a + b, 0);
  summarySheet.getRange(row++, 1).setValue(`Overall Total Sale: £${overallTotal.toFixed(2)}`);
}

