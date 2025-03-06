/**
 * Product Filter Script
 * This script allows filtering products by price range, color, size, and gender
 */

let sheet;
let headerRow;
let dataRange;
let data;

/**
 * Creates a menu item for the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Product Filters')
    .addItem('Open Filters', 'showFilterDialog')
    .addToUi();
}

/**
 * Shows the filter dialog to the user.
 */
function showFilterDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('FilterDialog')
    .evaluate()
    .setWidth(550)
    .setHeight(550)
    .setTitle('Filter Products');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Filter Products');
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets unique values for a specific column.
 */
function getUniqueValuesForColumn(columnName) {
  initializeSheetData();
  
  const columnIndex = headerRow.indexOf(columnName);
  if (columnIndex === -1) return [];
  
  const values = data.map(row => row[columnIndex]);
  const uniqueValues = [...new Set(values.filter(value => value !== ""))];
  return uniqueValues.sort();
}

/**
 * Initialize sheet data for processing.
 */
function initializeSheetData() {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    headerRow = values[0];
    data = values.slice(1);
  }
}

/**
 * Get data for populating the filter form
 */
function getFilterFormData() {
  try {
    return {
      colors: getUniqueValuesForColumn('COLOR'),
      sizes: getUniqueValuesForColumn('SIZE'),
      genders: getUniqueValuesForColumn('GENDER')
    };
  } catch (error) {
    throw new Error('Failed to load filter data: ' + error.message);
  }
}

/**
 * Filter products based on selected criteria and create new sheets with results
 */
function filterProducts(minPrice, maxPrice, color, size, gender) {
  try {
    initializeSheetData();
    
    let filterDesc = "Filtered_";
    
    if (minPrice !== "" || maxPrice !== "") {
      filterDesc += "Price";
      if (minPrice !== "") filterDesc += minPrice;
      filterDesc += "-";
      if (maxPrice !== "") filterDesc += maxPrice;
      filterDesc += "_";
    }
    
    if (color !== "") filterDesc += "C" + color.substring(0, 3) + "_";
    if (size !== "") filterDesc += "S" + size + "_";
    if (gender !== "") filterDesc += "G" + gender.substring(0, 1) + "_";
    
    if (filterDesc.length > 30) {
      filterDesc = filterDesc.substring(0, 27) + "...";
    }
    
    const filteredSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(filterDesc);
    
    filteredSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    const priceIndex = headerRow.indexOf("PRICE");
    const colorIndex = headerRow.indexOf("COLOR");
    const sizeIndex = headerRow.indexOf("SIZE");
    const genderIndex = headerRow.indexOf("GENDER");
    
    let filteredData = data.filter(row => {
      let priceMatch = true;
      let colorMatch = true;
      let sizeMatch = true;
      let genderMatch = true;
      
      if (minPrice !== "" && maxPrice !== "") {
        const price = parseFloat(row[priceIndex]);
        priceMatch = price >= parseFloat(minPrice) && price <= parseFloat(maxPrice);
      } else if (minPrice !== "") {
        const price = parseFloat(row[priceIndex]);
        priceMatch = price >= parseFloat(minPrice);
      } else if (maxPrice !== "") {
        const price = parseFloat(row[priceIndex]);
        priceMatch = price <= parseFloat(maxPrice);
      }
      
      if (color !== "") {
        colorMatch = row[colorIndex] === color;
      }
      
      if (size !== "") {
        sizeMatch = row[sizeIndex] === size;
      }
      
      if (gender !== "") {
        genderMatch = row[genderIndex] === gender;
      }
      
      return priceMatch && colorMatch && sizeMatch && genderMatch;
    });
    
    if (filteredData.length > 0) {
      filteredSheet.getRange(2, 1, filteredData.length, headerRow.length).setValues(filteredData);
    }
    
    filteredSheet.setFrozenRows(1);
    filteredSheet.autoResizeColumns(1, headerRow.length);
    
    const headerRange = filteredSheet.getRange(1, 1, 1, headerRow.length);
    headerRange.setBackground('#6e2ff2');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    const filterSummary = [];
    filterSummary.push(["Filter Summary"]);
    if (minPrice !== "" || maxPrice !== "") {
      const priceRange = `Price: ${minPrice === "" ? "Any" : '$' + minPrice} to ${maxPrice === "" ? "Any" : '$' + maxPrice}`;
      filterSummary.push([priceRange]);
    }
    if (color !== "") filterSummary.push([`Color: ${color}`]);
    if (size !== "") filterSummary.push([`Size: ${size}`]);
    if (gender !== "") filterSummary.push([`Gender: ${gender}`]);
    filterSummary.push([`Total results: ${filteredData.length}`]);
    
    const summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Summary_" + filterDesc);
    summarySheet.getRange(1, 1, filterSummary.length, 1).setValues(filterSummary);
    
    summarySheet.getRange(1, 1).setFontWeight('bold');
    summarySheet.getRange(1, 1).setBackground('#e8eaed');
    summarySheet.getRange(filterSummary.length, 1).setFontWeight('bold');
    summarySheet.autoResizeColumns(1, 1);
    
    filteredSheet.activate();
    
    return {
      success: true,
      message: `Filter applied: Found ${filteredData.length} products matching your criteria.`,
      sheetName: filterDesc
    };
  } catch (error) {
    throw new Error('Error applying filters: ' + error.message);
  }
}

/**
 * Reset all filters by activating the original sheet
 */
function resetAllFilters() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Filters',
    'This will activate the original data sheet. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    if (sheets.length > 0) {
      sheets[0].activate();
      ui.alert('Filters reset. Original data sheet activated.');
    }
  }
}