// Function to fetch and display spreadsheet data
function displaySpreadsheetData(sheetName) {
    const spreadsheetTable = document.getElementById('spreadsheet');

    // Fetch the spreadsheet file
    fetch('teams.xlsx') // Replace with your GitHub Pages URL
        .then((response) => response.arrayBuffer())
        .then((data) => {
            // Parse the spreadsheet data using SheetJS
            const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });

            // Get the selected sheet by name
            const worksheet = workbook.Sheets[sheetName];

            // Convert sheet data to HTML table
            const table = XLSX.utils.sheet_to_html(worksheet, {
                header: '<th></th>',
                tableClass: 'table table-bordered table-hover', // Add styling classes if desired
            });

            // Display the table in the designated element
            spreadsheetTable.innerHTML = table;
        })
        .catch((error) => {
            console.error('Error fetching or displaying spreadsheet:', error);
        });
}

// Function to create buttons for each sheet in the Excel file
function createSheetButtons() {
    const userButtonsContainer = document.getElementById('userButtons');

    // Fetch the spreadsheet file to get sheet names
    fetch('teams.xlsx') // Replace with your GitHub Pages URL
        .then((response) => response.arrayBuffer())
        .then((data) => {
            // Parse the spreadsheet data using SheetJS
            const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });

            // Get sheet names from the workbook
            const sheetNames = workbook.SheetNames;

            // Create buttons for each sheet
            sheetNames.forEach((sheetName) => {
                const button = document.createElement('button');
                button.textContent = sheetName;

                // Add a click event listener to open the corresponding sheet
                button.addEventListener('click', () => {
                    displaySpreadsheetData(sheetName);
                });

                userButtonsContainer.appendChild(button);
            });

            // Display the summary sheet by default when the page loads
            displaySpreadsheetData('Summary');
        })
        .catch((error) => {
            console.error('Error fetching or displaying spreadsheet:', error);
        });
}

// Call the function to create buttons and display the summary sheet when the page loads
window.onload = createSheetButtons;
