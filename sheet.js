let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations
let workbook; // This will hold the entire workbook object

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        // Load and display the first sheet
        loadSheet(workbook.SheetNames[0]);

        // Populate subsheet dropdown if subsheets exist
        const subsheetDropdown = document.getElementById('subsheet-dropdown');
        if (workbook.SheetNames.length > 1) {
            workbook.SheetNames.forEach((sheetName, index) => {
                if (index > 0) { // Skip the first sheet
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    subsheetDropdown.appendChild(option);
                }
            });
            subsheetDropdown.disabled = false; // Enable the dropdown
        }
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to load a specific sheet and display its data
function loadSheet(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    data = XLSX.utils.sheet_to_json(sheet, { defval: null });
    filteredData = [...data];
    displaySheet(filteredData);
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell; // Print 'NULL' for empty cells
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations and update the table
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    // Convert the entered column names (e.g., A, B, C) to column headers
    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());

    filteredData = data.filter(row => {
        // Check if the primary column is null or not
        const isPrimaryNull = row[primaryColumn] === null || row[primaryColumn] === "";

        // Apply the AND/OR logic
        const columnChecks = operationColumns.map(col => {
            if (operation === 'null') {
                return row[col] === null || row[col] === "";
            } else {
                return row[col] !== null && row[col] !== "";
            }
        });

        // Determine if we should display the row based on the selected operation type
        if (operationType === 'and') {
            return !isPrimaryNull && columnChecks.every(check => check);
        } else {
            return !isPrimaryNull && columnChecks.some(check => check);
        }
    });

    // Only display the primary column and the selected operation columns
    filteredData = filteredData.map(row => {
        const filteredRow = {};
        filteredRow[primaryColumn] = row[primaryColumn]; // Always show the primary column
        operationColumns.forEach(col => {
            filteredRow[col] = row[col] === null || row[col] === "" ? 'NULL' : row[col]; // Replace empty cells with 'NULL'
        });
        return filteredRow;
    });

    // Update the displayed table
    displaySheet(filteredData);
}

// Function to open the download modal
function openDownloadModal() {
    document.getElementById('download-modal').style.display = 'flex';
}

// Function to close the download modal
function closeDownloadModal() {
    document.getElementById('download-modal').style.display = 'none';
}

// Function to download filtered data as an Excel file or CSV
function downloadExcel() {
    const filename = document.getElementById('filename').value.trim() || 'download';
    const format = document.getElementById('file-format').value;

    // Ensure all null or empty cells are marked as 'NULL' in the exported data
    const exportData = filteredData.map(row => {
        return Object.keys(row).reduce((acc, key) => {
            acc[key] = row[key] === null || row[key] === "" ? 'NULL' : row[key]; // Ensure 'NULL' for empty cells
            return acc;
        }, {});
    });

    let worksheet = XLSX.utils.json_to_sheet(exportData);
    let workbookToDownload = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbookToDownload, worksheet, 'Filtered Data');

    if (format === 'xlsx') {
        XLSX.writeFile(workbookToDownload, `${filename}.xlsx`);
    } else if (format === 'csv') {
        XLSX.writeFile(workbookToDownload, `${filename}.csv`);
    }

    closeDownloadModal(); // Close the modal after downloading
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', openDownloadModal);
document.getElementById('confirm-download').addEventListener('click', downloadExcel);
document.getElementById('close-modal').addEventListener('click', closeDownloadModal);
document.getElementById('subsheet-dropdown').addEventListener('change', (e) => {
    const selectedSheet = e.target.value;
    if (selectedSheet) {
        loadSheet(selectedSheet);
    } else {
        loadSheet(workbook.SheetNames[0]); // Default back to the first sheet if no subsheet is selected
    }
});

// Initial load (example URL)
loadExcelSheet('your-excel-file-url-here.xlsx'); // Replace with the actual file URL
