
// Function to read the CSV file and display checkboxes
function readCSV(input) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        const data = XLSX.read(text, {type: 'string'});
        const sheetName = data.SheetNames[0];
        const sheet = data.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        displayCheckboxes(json);
    };
    reader.readAsBinaryString(input.files[0]);
}

// Function to display checkboxes for each unique value in the first column
function displayCheckboxes(data) {
    const container = document.getElementById('checkbox-area');
    container.innerHTML = ''; // Clear previous checkboxes
    data.forEach((row, index) => {
        if (index > 0) { // Skip header row
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `checkbox-${index}`;
            // Store the entire row data, but display only the first column value
            checkbox.dataset.rowData = JSON.stringify(row);
            const label = document.createElement('label');
            label.htmlFor = `checkbox-${index}`;
            label.textContent = row[0]; // Display only the first column value
            container.appendChild(checkbox);
            container.appendChild(label);
            container.appendChild(document.createElement('br'));
        }
    });
    document.getElementById('extract').disabled = false; // Enable extract button
}


// Function to extract selected data (first and last columns) and save as Excel
function extractData() {
    const selectedData = [];
    document.querySelectorAll('#checkbox-area input[type="checkbox"]:checked').forEach(checkbox => {
        const fullRowData = JSON.parse(checkbox.dataset.rowData);
        // Push an array containing only the first and last column values
        selectedData.push([fullRowData[0], fullRowData[fullRowData.length - 1]]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(selectedData);
    XLSX.utils.book_append_sheet(wb, ws, 'ExtractedData');
    XLSX.writeFile(wb, 'ExtractedData.xlsx');
}


// Event listeners
document.getElementById('read-file').addEventListener('click', function() {
    const fileInput = document.getElementById('file-input');
    if (fileInput.files.length > 0) {
        readCSV(fileInput);
    }
});

document.getElementById('extract').addEventListener('click', extractData);

document.getElementById('reset').addEventListener('click', function() {
    document.getElementById('checkbox-area').innerHTML = '';
    document.getElementById('extract').disabled = true; // Disable extract button
    document.getElementById('file-input').value = ''; // Reset file input
});

document.getElementById('help').addEventListener('click', function() {
    alert("How to Use This Application:\n\n" 
    + "Step 1: Download your Google Sheets document as a CSV file.\n"
    + "        - Navigate to your Google Sheets document.\n"
    + "        - Go to File > Download > Comma-separated values (.csv).\n\n"
    + "Step 2: Load your CSV file into the application.\n"
    + "        - Click the 'Browse' button next to the 'File(s)' field.\n"
    + "        - Navigate to and select your downloaded .csv file.\n\n"
    + "Step 3: Read the file and select data to extract.\n"
    + "        - Click the 'Read File' button to display the data labels from the file.\n"
    + "        - Use the checkboxes to select the labels (data rows) you wish to extract.\n\n"
    + "Step 4: Extract the selected data.\n"
    + "        - Once you've selected the desired labels, click the 'Extract' button.\n"
    + "        - The application will extract the selected data and save it in a new Excel (.xlsx) file.\n"
    + "        - The new file will be located in the 'Extracted-Files' folder within the application's directory.\n\n"
    + "Step 5: Reset the application (optional).\n"
    + "        - To start a new extraction process or to load a different CSV file, click the 'Reset' button.\n"
    + "        - This will clear all current selections and allow you to begin the process anew.\n\n"
    + "For further assistance, contact support.");
});