//document.addEventListener('DOMContentLoaded', function () {

//    const excelDropdown = document.getElementById('excelFilesDropdown');
//    const columnsList = document.getElementById('columnsList');
//    const saveBtn = document.getElementById('saveTemplateBtn');
//    const bulkBtn = document.getElementById('bulkGenerateBtn');
//    let selectedFile = null;

//    // =========================
//    // 1️⃣ Load Excel columns
//    // =========================
//    excelDropdown.addEventListener('change', () => {
//        selectedFile = excelDropdown.value;
//        columnsList.innerHTML = '';

//        if (!selectedFile) return;

//        fetch(`/Document/GetColumns?fileName=${selectedFile}`)
//            .then(res => res.json())
//            .then(columns => {
//                columns.forEach(col => {
//                    const li = document.createElement('li');
//                    li.textContent = col;
//                    li.className = 'list-group-item';
//                    li.setAttribute('draggable', 'true');
//                    li.style.cursor = 'grab';

//                    // Drag start event
//                    li.addEventListener('dragstart', ev => {
//                        ev.dataTransfer.setData('text/plain', `{{${col}}}`);
//                    });

//                    columnsList.appendChild(li);
//                });
//            })
//            .catch(err => console.error('Error fetching columns:', err));
//    });

//    // =========================
//    // 2️⃣ Initialize TinyMCE
//    // =========================
//    tinymce.init({
//        selector: '#docEditor',
//        height: 600,
//        menubar: true,
//        plugins: 'lists link image table code',
//        toolbar: 'undo redo | bold italic underline | styles | bullist numlist | table | code',
//        entity_encoding: 'raw',
//        valid_elements: '*[*]',
//        verify_html: false,
//        forced_root_block: 'p',
//        setup: function (editor) {
//            editor.on('dragover', e => e.preventDefault());
//            editor.on('drop', e => {
//                e.preventDefault();
//                const data = e.dataTransfer.getData('text/plain');
//                editor.selection.setContent(data);
//            });
//        }
//    });

//    // =========================
//    // 3️⃣ Save template
//    // =========================
//    saveBtn.addEventListener('click', () => {
//        const templateHtml = tinymce.get('docEditor').getContent({ format: 'raw' });

//        if (!templateHtml || templateHtml.trim() === '') {
//            alert('Editor is empty!');
//            return;
//        }

//        fetch('/Document/SaveTemplate', {
//            method: 'POST',
//            headers: { 'Content-Type': 'application/json' },
//            body: JSON.stringify({
//                fileName: selectedFile || null,
//                templateHtml: templateHtml
//            })
//        })
//            .then(res => {
//                if (res.ok) alert('Template saved successfully!');
//                else alert('Failed to save template.');
//            })
//            .catch(err => console.error('Error saving template:', err));
//    });

//    // =========================
//    // 4️⃣ Bulk Generate DOCX
//    // =========================
//    bulkBtn.addEventListener('click', () => {

//        if (!selectedFile) {
//            alert('Please select an Excel file first!');
//            return;
//        }

//        const templateHtml = tinymce.get('docEditor').getContent({ format: 'raw' });
//        if (!templateHtml || templateHtml.trim() === '') {
//            alert('Template is empty!');
//            return;
//        }

//        // Ask user which column to use for file naming
//        const fileNameColumn = prompt('Enter column name to use for document file names (e.g., EmployeeName):');
//        if (!fileNameColumn) return;

//        fetch('/Document/GenerateDocuments', {
//            method: 'POST',
//            headers: { 'Content-Type': 'application/json' },
//            body: JSON.stringify({
//                excelFileName: selectedFile,
//                templateHtml: templateHtml,
//                fileNameColumn: fileNameColumn
//            })
//        })
//            .then(res => res.blob())
//            .then(blob => {
//                const url = window.URL.createObjectURL(blob);
//                const a = document.createElement('a');
//                a.href = url;
//                a.download = 'GeneratedDocs.zip';
//                document.body.appendChild(a);
//                a.click();
//                a.remove();
//            })
//            .catch(err => console.error('Error generating documents:', err));
//    });

//});




////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
document.addEventListener('DOMContentLoaded', function () {

    const excelDropdown = document.getElementById('excelFilesDropdown');
    const columnsList = document.getElementById('columnsList');
    const generateBtn = document.getElementById('generateDocsBtn');
    let selectedFile = null;

    // Load Excel columns
    excelDropdown.addEventListener('change', () => {
        selectedFile = excelDropdown.value;
        columnsList.innerHTML = '';

        if (!selectedFile) return;

        fetch(`/Document/GetColumns?fileName=${selectedFile}`)
            .then(res => res.json())
            .then(columns => {
                columns.forEach(col => {
                    const li = document.createElement('li');
                    li.textContent = col;
                    li.className = 'list-group-item';
                    li.setAttribute('draggable', 'true');
                    li.style.cursor = 'grab';

                    li.addEventListener('dragstart', ev => {
                        ev.dataTransfer.setData('text/plain', `{{${col}}}`);
                    });

                    columnsList.appendChild(li);
                });
            });
    });

    // Initialize TinyMCE
    tinymce.init({
        selector: '#docEditor',
        height: 600,
        menubar: true,
        plugins: 'lists link image table code',
        toolbar: 'undo redo | bold italic underline | styles | bullist numlist | table | code',
        entity_encoding: 'raw',
        valid_elements: '*[*]',
        verify_html: false,
        forced_root_block: 'p',
        setup: function (editor) {
            editor.on('dragover', e => e.preventDefault());
            editor.on('drop', e => {
                e.preventDefault();
                const data = e.dataTransfer.getData('text/plain');
                editor.selection.setContent(data);
            });
        }
    });

    // Generate documents
    generateBtn.addEventListener('click', () => {
        if (!selectedFile) {
            alert('Please select an Excel file.');
            return;
        }

        const templateHtml = tinymce.get('docEditor').getContent({ format: 'raw' });

        if (!templateHtml || templateHtml.trim() === '') {
            alert('Template is empty!');
            return;
        }

        // Ask for column to use for filename
        const fileNameColumn = prompt('Enter column name to use for generated file names:');
        if (!fileNameColumn) return;

        fetch('/Document/GenerateDocuments', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                ExcelFileName: selectedFile,
                TemplateHtml: templateHtml,
                FileNameColumn: fileNameColumn
            })
        })
            .then(res => res.blob())
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'GeneratedDocs.zip';
                document.body.appendChild(a);
                a.click();
                a.remove();
            })
            .catch(err => console.error(err));
    });
});

saveBtn.addEventListener('click', () => {

    const editor = tinymce.get('docEditor');
    const templateHtml = editor.getContent({ format: 'raw' });

    if (!selectedFile) {
        alert('Please select an Excel file');
        return;
    }

    if (!templateHtml || templateHtml.trim() === '') {
        alert('Editor is empty');
        return;
    }

    // 🔹 Extract placeholders like {{EmployeeName}}
    const matches = templateHtml.match(/{{(.*?)}}/g) || [];

    if (matches.length === 0) {
        alert('No placeholders found in document');
        return;
    }

    // 🔹 Build column mapping
    const columnMapping = {};
    matches.forEach(m => {
        const key = m.replace('{{', '').replace('}}', '').trim();
        columnMapping[key] = key; // Excel column = placeholder name
    });

    console.log('Column Mapping:', columnMapping);

    // 🔹 Call API
    fetch('/Document/GenerateDocuments', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            ExcelFileName: selectedFile,
            TemplateHtml: templateHtml,
            ColumnMapping: columnMapping
        })
    })
        .then(res => {
            if (!res.ok) throw new Error('Failed');
            return res.blob();
        })
        .then(blob => {
            // 🔹 Download ZIP
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'GeneratedDocs.zip';
            document.body.appendChild(a);
            a.click();
            a.remove();
        })
        .catch(err => {
            console.error(err);
            alert('Error generating documents');
        });
});