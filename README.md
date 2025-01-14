# google-sheets-clone

Overview


This project is a Google Sheets clone built using HTML, CSS, and JavaScript. It provides a spreadsheet interface with functionalities such as mathematical operations, data quality functions, text formatting, and row/column operations.

Tech Stack
HTML: Used for structuring the web page.
CSS: Used for styling the web page.
JavaScript: Used for adding interactivity and functionality to the web page.
Data Structures
1. Spreadsheet Data
The spreadsheet data is stored in a 2D array where each cell is represented by an object containing its content and style.

2. Cell Dependencies
A dictionary (cellDependencies) is used to keep track of dependencies between cells. This is useful for updating dependent cells when a cell's value changes.

3. Event Listeners
Event listeners are used extensively to handle user interactions such as input, button clicks, and cell selection.

Why These Technologies?
HTML and CSS: These are standard technologies for building web pages. HTML provides the structure, while CSS provides the styling.
JavaScript: JavaScript is essential for adding interactivity to web pages. It allows us to manipulate the DOM, handle events, and perform calculations.
Features
1. Spreadsheet Creation
The spreadsheet is dynamically generated using JavaScript. Rows and columns can be added or deleted, and cells can be edited.

2. Mathematical Functions
Functions such as SUM, AVERAGE, MAX, MIN, COUNT, MEDIAN, and MODE are implemented to perform calculations on selected ranges of cells.

3. Data Quality Functions
Functions such as TRIM, UPPER, LOWER, REMOVE_DUPLICATES, and FIND_AND_REPLACE are implemented to clean and manipulate data.

4. Text Formatting
Cells can be formatted with bold, italic, underline, font size, and font color.

5. Row and Column Operations
Rows and columns can be added, deleted, and resized.

6. Save and Load
Spreadsheet data can be saved to and loaded from localStorage.

Conclusion
This project demonstrates how to build a functional spreadsheet application using HTML, CSS, and JavaScript. The use of event listeners and data structures such as 2D arrays and dictionaries allows for efficient handling of user interactions and data dependencies.
