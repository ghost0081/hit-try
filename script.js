// Function to handle image preview
function previewImage(input) {
    const preview = input.nextElementSibling;
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            preview.src = e.target.result;
            preview.classList.remove('hidden');
        }
        reader.readAsDataURL(input.files[0]);
    } else {
        preview.src = '';
        preview.classList.add('hidden');
    }
}

// Function to add row to value table
function addValueRow() {
    const tbody = document.querySelector('#valueTable tbody');
    const newRow = tbody.insertRow();
    newRow.innerHTML = `
        <td><input type="text" name="userId[]" required></td>
        <td><input type="text" name="userName[]" required></td>
        <td><input type="text" name="segment[]" required></td>
        <td><input type="text" name="approvedAddress[]" required></td>
        <td><input type="text" name="actualLocatedAt[]" required></td>
        <td><input type="text" name="remarks[]" required></td>
        <td><input type="text" name="authorizedPersonReply[]" required></td>
        <td>
            <input type="file" name="evidence[]" accept="image/*" onchange="previewImage(this)">
            <img src="" class="image-preview hidden" alt="Image Preview">
        </td>
        <td><button type="button" onclick="removeRow(this)">Remove</button></td>
    `;
}

// Function to add row to additional table
function addAdditionalRow() {
    const tbody = document.querySelector('#additionalTable tbody');
    const newRow = tbody.insertRow();
    newRow.innerHTML = `
        <td><input type="text" name="loginId[]" required></td>
        <td><input type="text" name="approvedUserName[]" required></td>
        <td><input type="text" name="additionalSegment[]" required></td>
        <td><input type="text" name="operatedBy[]" required></td>
        <td><select name="ncfmQualified[]"><option value="Y">Y</option><option value="N">N</option></select></td>
        <td><select name="approvedUserAvailable[]"><option value="Y">Y</option><option value="N">N</option></select></td>
        <td><input type="text" name="additionalAuthorizedPersonReply[]" required></td>
        <td><button type="button" onclick="removeRow(this)">Remove</button></td>
    `;
}

// Function to add row to client table
function addClientRow() {
    const tbody = document.querySelector('#clientTable tbody');
    const newRow = tbody.insertRow();
    newRow.innerHTML = `
        <td><input type="text" name="clientCode[]" required></td>
        <td><input type="text" name="clientName[]" required></td>
        <td><input type="text" name="personTrading[]" required></td>
        <td><input type="text" name="relation[]" required></td>
        <td><select name="al[]"><option value="Y">Y</option><option value="N">N</option></select></td>
        <td><input type="text" name="clientAuthorizedPersonReply[]" required></td>
        <td><button type="button" onclick="removeRow(this)">Remove</button></td>
    `;
}

// Function to remove row
function removeRow(button) {
    button.closest('tr').remove();
}

// Function to collect form data
function collectFormData() {
    const formData = {
        basicInfo: {},
        valueTable: [],
        boardsTable: [],
        additionalTable: [],
        clientTable: [],
        fixedAttributes: []
    };

    // Collect basic information
    const basicInfoInputs = document.querySelectorAll('.form-section input:not([type="file"])');
    basicInfoInputs.forEach(input => {
        if (!input.name.includes('[]')) {
            formData.basicInfo[input.name] = input.value;
        }
    });

    // Collect value table data
    document.querySelectorAll('#valueTable tbody tr').forEach(row => {
        const rowData = {};
        row.querySelectorAll('input:not([type="file"])').forEach(input => {
            rowData[input.name.replace('[]', '')] = input.value;
        });
        formData.valueTable.push(rowData);
    });

    // Collect boards table data
    document.querySelectorAll('#boardTable tbody tr').forEach(row => {
        const rowData = {
            srNo: row.cells[0].textContent,
            listOfBoards: row.cells[1].textContent,
            yesNo: row.querySelector('input[name="yesNo[]"]')?.value || '',
            remarks: row.querySelector('input[name="boardRemarks[]"]')?.value || '',
            authorizedPersonReply: row.querySelector('input[name="boardAuthorizedPersonReply[]"]')?.value || ''
        };
        formData.boardsTable.push(rowData);
    });

    // Collect additional table data
    document.querySelectorAll('#additionalTable tbody tr').forEach(row => {
        const rowData = {};
        row.querySelectorAll('input, select').forEach(input => {
            rowData[input.name.replace('[]', '')] = input.value;
        });
        formData.additionalTable.push(rowData);
    });

    // Collect client table data
    document.querySelectorAll('#clientTable tbody tr').forEach(row => {
        const rowData = {};
        row.querySelectorAll('input, select').forEach(input => {
            rowData[input.name.replace('[]', '')] = input.value;
        });
        formData.clientTable.push(rowData);
    });

    // Collect fixed attributes data
    document.querySelectorAll('#fixedAttributesTable tbody tr').forEach(row => {
        formData.fixedAttributes.push({
            particulars: row.cells[0].textContent,
            inspection: row.querySelector('input[name^="inspection"]')?.value || '',
            apReply: row.querySelector('input[name^="apReply"]')?.value || ''
        });
    });

    return formData;
}

// Function to export to Excel
function exportToExcel() {
    const formData = collectFormData();
    const workbook = XLSX.utils.book_new();

    // Create Basic Information worksheet
    const basicInfoSheet = XLSX.utils.json_to_sheet([formData.basicInfo]);
    XLSX.utils.book_append_sheet(workbook, basicInfoSheet, "Basic Information");

    // Create Value Table worksheet
    if (formData.valueTable.length > 0) {
        const valueTableSheet = XLSX.utils.json_to_sheet(formData.valueTable);
        XLSX.utils.book_append_sheet(workbook, valueTableSheet, "Terminal Details");
    }

    // Create Boards Table worksheet
    if (formData.boardsTable.length > 0) {
        const boardsTableSheet = XLSX.utils.json_to_sheet(formData.boardsTable);
        XLSX.utils.book_append_sheet(workbook, boardsTableSheet, "Boards Details");
    }

    // Create Additional Table worksheet
    if (formData.additionalTable.length > 0) {
        const additionalTableSheet = XLSX.utils.json_to_sheet(formData.additionalTable);
        XLSX.utils.book_append_sheet(workbook, additionalTableSheet, "User Operations");
    }

    // Create Client Table worksheet
    if (formData.clientTable.length > 0) {
        const clientTableSheet = XLSX.utils.json_to_sheet(formData.clientTable);
        XLSX.utils.book_append_sheet(workbook, clientTableSheet, "Client Trading Details");
    }

    // Create Fixed Attributes worksheet
    if (formData.fixedAttributes.length > 0) {
        const fixedAttributesSheet = XLSX.utils.json_to_sheet(formData.fixedAttributes);
        XLSX.utils.book_append_sheet(workbook, fixedAttributesSheet, "Particulars");
    }

    // Save the Excel file
    const fileName = `AP_Survey_${formData.basicInfo.apMainCode || 'Report'}_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}

// Add event listeners when the document loads
document.addEventListener('DOMContentLoaded', function() {
    // Add button event listeners
    document.getElementById('addRowBtn')?.addEventListener('click', addValueRow);
    document.getElementById('addAdditionalRowBtn')?.addEventListener('click', addAdditionalRow);
    document.getElementById('addClientRowBtn')?.addEventListener('click', addClientRow);

    // Add export button if it doesn't exist
    if (!document.getElementById('exportToExcelBtn')) {
        const exportBtn = document.createElement('button');
        exportBtn.type = 'button';
        exportBtn.id = 'exportToExcelBtn';
        exportBtn.textContent = 'Export to Excel';
        exportBtn.addEventListener('click', exportToExcel);
        document.querySelector('form').appendChild(exportBtn);
    }

    // Initialize file input change handlers
    document.querySelectorAll('input[type="file"]').forEach(input => {
        input.addEventListener('change', function() {
            previewImage(this);
        });
    });
});