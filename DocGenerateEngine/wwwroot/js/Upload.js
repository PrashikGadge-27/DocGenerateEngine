document.addEventListener('DOMContentLoaded', function () {
    // Get the form and elements
    const uploadForm = document.getElementById('uploadForm') || document.querySelector('form[asp-action="Upload"]');
    const fileInput = document.getElementById('excelFile');

    // Create a message div if it doesn't exist
    let messageDiv = document.getElementById('uploadMessage');
    if (!messageDiv) {
        messageDiv = document.createElement('div');
        messageDiv.id = 'uploadMessage';
        messageDiv.style.marginTop = '15px';
        uploadForm.parentNode.appendChild(messageDiv);
    }

    uploadForm.addEventListener('submit', function (event) {
        event.preventDefault(); // stop default form submission

        // Check if a file is selected
        if (!fileInput.files.length) {
            messageDiv.innerText = 'Please select a file before uploading.';
            messageDiv.style.color = 'red';
            return;
        }

        const file = fileInput.files[0];
        const formData = new FormData();
        formData.append('excelFile', file);

        // Optional: disable button while uploading
        const submitBtn = uploadForm.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        submitBtn.innerText = 'Uploading...';

        // Make AJAX request using fetch
        fetch('/Excel/Upload', {
            method: 'POST',
            body: formData
        })
            .then(response => {
                if (!response.ok) throw new Error('Upload failed');
                return response.text(); // or JSON if your controller returns JSON
            })
            .then(data => {
                messageDiv.innerText = 'File uploaded successfully!';
                messageDiv.style.color = 'green';
                fileInput.value = ''; // clear file input
            })
            .catch(error => {
                messageDiv.innerText = error.message;
                messageDiv.style.color = 'red';
            })
            .finally(() => {
                submitBtn.disabled = false;
                submitBtn.innerText = 'Upload';
            });
    });
});
