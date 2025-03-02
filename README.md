Google Sheets Product Filter Add-on
===================================

A user-friendly Google Sheets add-on that allows users to filter product data based on price range, color, size, and gender. The add-on creates filtered views of your data in new sheets.
Features

Used in this [sheet](https://docs.google.com/spreadsheets/d/1D33Sm7l4jnU_N6myQhxRAmtK1EVag_CtraIuPBji8V4/edit?usp=sharing)
--------

-   **Multi-criteria Filtering**: Filter products by price range, color, size, and gender
-   **Dynamic Dropdown Options**: Filter options are automatically populated from your sheet data
-   **Clean UI**: Modern Google-style interface with responsive design
-   **Summary Reports**: Each filter operation creates a summary sheet with filter criteria and result count
-   **Non-destructive**: Original data remains intact; filtered results appear in new sheets
-   **Data Validation**: Input validation for price ranges
-   **Reset Functionality**: Easy to clear all filters or return to original data

Installation
------------

### Using the Script Files Directly

1.  Open your Google Sheets document containing product data
2.  Go to **Extensions > Apps Script**
3.  Create the following files in the Apps Script editor:
    -   `Code.gs`: Copy the content from the provided script
    -   `FilterDialog.html`: Create this HTML file
    -   `Stylesheet.html`: Create this HTML file with CSS content
    -   `JavaScript.html`: Create this HTML file with JS content
4.  Save the project with a name (e.g., "Product Filter Add-on")
5.  Reload your Google Sheets document
6.  A new menu item "Product Filters" will appear in the menu bar

Usage
-----

### Preparing Your Data

For best results, your data should:

1.  Have a header row with column names in UPPERCASE (e.g., "PRICE", "COLOR", "SIZE", "GENDER")
2.  Organize product data in rows below the header
3.  Ensure the price column contains numeric values

### Using the Filter Dialog

1.  Click on the **Product Filters** menu in the menu bar
2.  Select **Show Filter Dialog**
3.  In the dialog that appears:
    -   Set minimum and/or maximum price (optional)
    -   Select a color from the dropdown (optional)
    -   Select a size from the dropdown (optional)
    -   Select a gender from the dropdown (optional)
4.  Click **Apply Filters** to generate a new sheet with the filtered data
5.  Review the generated summary sheet for filter details

### Resetting Filters

To return to the original data sheet:

1.  Click on the **Product Filters** menu
2.  Select **Reset All Filters**
3.  Confirm when prompted

Project Structure
-----------------

The add-on consists of four main files:

1.  **Code.gs**: Server-side JavaScript that handles the add-on logic, data processing, and spreadsheet operations
2.  **FilterDialog.html**: The main HTML structure for the filter dialog UI
3.  **Stylesheet.html**: CSS styles for the dialog interface
4.  **JavaScript.html**: Client-side JavaScript for handling user interactions and form validation

### Code.gs

Contains the following main functions:

-   `onOpen()`: Creates the add-on menu
-   `showFilterDialog()`: Displays the filter UI
-   `getUniqueValuesForColumn()`: Retrieves unique values for dropdowns
-   `filterProducts()`: Applies filters and creates new sheets
-   `resetAllFilters()`: Returns to the original data sheet

### FilterDialog.html

The main UI template that includes:

-   Header and description
-   Form inputs for all filter criteria
-   Action buttons
-   Loading indicator
-   References to the style and script files

### Stylesheet.html

Contains CSS styles for the filter dialog, including:

-   Google Material Design-inspired styling
-   Responsive design elements
-   Form styling and input validation indicators
-   Loading animation

### JavaScript.html

Contains client-side logic for:

-   Initializing the form and loading filter options
-   Handling form validation
-   Sending filter requests to the server
-   Managing the UI state and feedback

Technical Implementation
------------------------

### Server-Side (Google Apps Script)

The add-on leverages Google Apps Script to interact with the spreadsheet:

1.  **Data Initialization**:

    -   The `initializeSheetData()` function loads the active sheet's data
    -   Header row is identified and separated from data rows
    -   Data is cached to improve performance during filtering operations
2.  **Dynamic Filter Options**:

    -   `getUniqueValuesForColumn()` extracts unique values from specified columns
    -   Values are sorted and returned for populating dropdown menus
3.  **Filtering Logic**:

    -   `filterProducts()` applies selected filters to the dataset
    -   Creates a new sheet with a timestamped name for the filtered results
    -   Applies formatting to the new sheet for better readability
4.  **Summary Generation**:

    -   Creates a separate summary sheet with filter criteria and result counts
    -   Formats the summary sheet for easy reference

### Client-Side (HTML/CSS/JavaScript)

The client-side components handle the user interface and interaction:

1.  **Initialization**:

    -   On load, the dialog requests filter options from the server
    -   Populates dropdown menus with available options
2.  **Validation**:

    -   Price inputs are validated to ensure max price is greater than min price
    -   Real-time validation feedback is provided
3.  **UI/UX Design**:

    -   Google Material Design principles for a familiar interface
    -   Responsive layout that works on different screen sizes
    -   Loading indicators for async operations
    -   Clear status messages for operation results
4.  **Form Submission**:

    -   Collects and validates all filter criteria
    -   Sends data to the server for processing
    -   Handles success and error states
    -   Provides user feedback and auto-closes on success

Customization and Maintenance
-----------------------------

### Adding New Filter Criteria

To add a new filter criterion (e.g., "BRAND"):

1.  **Update Code.gs**:

    -   Add the new criterion to the `getFilterFormData()` function

    ```
    function getFilterFormData() {
      return {
        colors: getUniqueValuesForColumn('COLOR'),
        sizes: getUniqueValuesForColumn('SIZE'),
        genders: getUniqueValuesForColumn('GENDER'),
        brands: getUniqueValuesForColumn('BRAND') // New criterion
      };
    }

    ```

    -   Update the `filterProducts()` function to include the new criterion
2.  **Update FilterDialog.html**:

    -   Add a new form group for the criterion

    ```
    <div class="form-group">
      <label for="brand">Brand</label>
      <div class="select-wrapper">
        <select id="brand">
          <option value="">All Brands</option>
        </select>
      </div>
    </div>

    ```

3.  **Update JavaScript.html**:

    -   Update the `populateFormOptions()` function to handle the new criterion
    -   Update the `applyFilters()` function to include the new criterion

### Modifying the UI

To change the appearance of the dialog:

1.  **Update Stylesheet.html**:

    -   Modify colors, fonts, or spacing as needed
    -   Test on different screen sizes for responsiveness
2.  **Update FilterDialog.html**:

    -   Modify layout or add new UI elements
    -   Update text and labels

### Performance Optimization

For large datasets:

1.  Consider implementing pagination for the filtered results
2.  Add caching mechanisms for frequently used data
3.  Implement a more efficient filtering algorithm for very large datasets

Troubleshooting
---------------

### Common Issues and Solutions

1.  **Menu Not Appearing**:

    -   Reload the spreadsheet
    -   Check if the script has proper authorization
    -   Verify the `onOpen()` function is correctly implemented
2.  **Empty Dropdown Options**:

    -   Verify column names match exactly (case-sensitive)
    -   Check if the data columns contain values
    -   Look for errors in the browser console or Apps Script logs
3.  **Filter Not Working**:

    -   Ensure price values are numeric
    -   Check for script timeout errors with large datasets
    -   Verify filter logic in the `filterProducts()` function
4.  **Dialog Not Loading**:

    -   Check browser console for JavaScript errors
    -   Verify HTML structure and included files
    -   Check for issues with Google's HTML service

### Debug Mode

To enable debug mode for troubleshooting:

1.  Add the following to the top of **Code.gs**:

    ```
    const DEBUG = true;

    function logDebug(message) {
      if (DEBUG) {
        Logger.log(message);
      }
    }

    ```

2.  Add debug logging throughout the code:

    ```
    logDebug('Filter called with: ' + JSON.stringify({minPrice, maxPrice, color, size, gender}));

    ```

3.  View logs in the Apps Script editor under "Execution Log"


