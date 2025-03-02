/**
 * Product Filter Script
 * This script allows filtering products by price range, color, size, and gender
 */

// Global variables
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
    .addItem('Show Filter Dialog', 'showFilterDialog')
    .addSeparator()
    .addItem('Reset All Filters', 'resetAllFilters')
    .addToUi();
}

/**
 * Shows the filter dialog to the user.
 */
function showFilterDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('FilterDialog')
    .evaluate()
    .setWidth(400)
    .setHeight(500)
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
  
  // Get all values in the column
  const values = data.map(row => row[columnIndex]);
  
  // Get unique values (excluding empty cells)
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
    data = values.slice(1); // Exclude header row
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
 * Filter products based on selected criteria.
 */
function filterProducts(minPrice, maxPrice, color, size, gender) {
  try {
    initializeSheetData();
    
    // Create a new filtered sheet with timestamp to make each filter result unique
    const timestamp = new Date().toLocaleString().replace(/[\/\s,:]/g, "_");
    const sheetName = `Filtered_${timestamp}`;
    const filteredSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    
    // Set header row
    filteredSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    // Get column indexes
    const priceIndex = headerRow.indexOf("PRICE");
    const colorIndex = headerRow.indexOf("COLOR");
    const sizeIndex = headerRow.indexOf("SIZE");
    const genderIndex = headerRow.indexOf("GENDER");
    
    // Filter data based on criteria
    let filteredData = data.filter(row => {
      let priceMatch = true;
      let colorMatch = true;
      let sizeMatch = true;
      let genderMatch = true;
      
      // Check price range if provided
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
      
      // Check color if provided
      if (color !== "") {
        colorMatch = row[colorIndex] === color;
      }
      
      // Check size if provided
      if (size !== "") {
        sizeMatch = row[sizeIndex] === size;
      }
      
      // Check gender if provided
      if (gender !== "") {
        genderMatch = row[genderIndex] === gender;
      }
      
      // Item matches if it passes all enabled filters
      return priceMatch && colorMatch && sizeMatch && genderMatch;
    });
    
    // Write filtered data
    if (filteredData.length > 0) {
      filteredSheet.getRange(2, 1, filteredData.length, headerRow.length).setValues(filteredData);
    }
    
    // Format and adjust columns
    filteredSheet.setFrozenRows(1);
    filteredSheet.autoResizeColumns(1, headerRow.length);
    
    // Style the header row
    const headerRange = filteredSheet.getRange(1, 1, 1, headerRow.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Add filter summary at the top of the sheet
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
    
    // Insert a new sheet for the summary
    const summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`Summary_${timestamp}`);
    summarySheet.getRange(1, 1, filterSummary.length, 1).setValues(filterSummary);
    
    // Style the summary sheet
    summarySheet.getRange(1, 1).setFontWeight('bold');
    summarySheet.getRange(1, 1).setBackground('#e8eaed');
    summarySheet.getRange(filterSummary.length, 1).setFontWeight('bold');
    summarySheet.autoResizeColumns(1, 1);
    
    // Activate the filtered sheet
    filteredSheet.activate();
    
    return {
      success: true,
      message: `Filter applied: Found ${filteredData.length} products matching your criteria.`,
      sheetName: sheetName
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
    // Get the first sheet (assuming it's the original data sheet)
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    if (sheets.length > 0) {
      sheets[0].activate();
      ui.alert('Filters reset. Original data sheet activated.');
    }
  }
}