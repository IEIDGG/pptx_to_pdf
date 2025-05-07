document.addEventListener('DOMContentLoaded', () => {
    const pptxFile = document.getElementById('pptxFile');
    const convertButton = document.getElementById('convertButton');
    const resultDiv = document.getElementById('result');
    const errorDiv = document.getElementById('error');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const fileLabel = document.querySelector('.file-label');
    const defaultFileLabelText = 'Choose a PPTX file';

    pptxFile.addEventListener('change', () => {
        if (pptxFile.files.length > 0) {
            fileLabel.textContent = pptxFile.files[0].name;
        } else {
            fileLabel.textContent = defaultFileLabelText;
        }
    });

    convertButton.addEventListener('click', async () => {
        const file = pptxFile.files[0];
        resultDiv.innerHTML = '';
        errorDiv.innerHTML = '';
        loadingSpinner.style.display = 'none';

        if (!file) {
            errorDiv.textContent = 'Please select a PPTX file.';
            return;
        }

        if (!file.name.toLowerCase().endsWith('.pptx')) {
            errorDiv.textContent = 'Invalid file type. Only PPTX files are accepted.';
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        convertButton.disabled = true;
        convertButton.textContent = 'Converting...';
        loadingSpinner.style.display = 'block';

        try {
            const response = await fetch('/convert/', {
                method: 'POST',
                body: formData,
            });

            if (response.ok) {
                const blob = await response.blob();
                const downloadUrl = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = downloadUrl;
                a.download = file.name.replace(/\.pptx$/i, '.pdf');
                a.textContent = 'Download PDF';
                resultDiv.appendChild(a);
            } else {
                const errorData = await response.json();
                errorDiv.textContent = `Error: ${errorData.detail || response.statusText}`;
            }
        } catch (error) {
            console.error('Conversion error:', error);
            errorDiv.textContent = 'An unexpected error occurred. Please try again.';
        } finally {
            convertButton.disabled = false;
            convertButton.textContent = 'Convert to PDF';
            loadingSpinner.style.display = 'none';
            pptxFile.value = ''; // Clear the file input
            fileLabel.textContent = defaultFileLabelText; // Reset file label
        }
    });
}); 