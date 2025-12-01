document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file');
    const fileInfo = document.getElementById('file-info');
    const filenameDisplay = document.getElementById('filename');
    const removeFileBtn = document.getElementById('remove-file');
    const form = document.getElementById('upload-form');
    const generateBtn = document.getElementById('generate-btn');
    const statusArea = document.getElementById('status-area');
    const resultArea = document.getElementById('result-area');
    const downloadLink = document.getElementById('download-link');
    const resetBtn = document.getElementById('reset-btn');

    // Drag and Drop functionality
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    function highlight(e) {
        dropZone.classList.add('dragover');
    }

    function unhighlight(e) {
        dropZone.classList.remove('dragover');
    }

    dropZone.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', function () {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (file.name.endsWith('.xlsx')) {
                filenameDisplay.textContent = file.name;
                fileInfo.style.display = 'flex';
                document.querySelector('.upload-content').style.display = 'none';

                // Manually set the file input files if coming from drop
                if (fileInput.files.length === 0 || fileInput.files[0] !== file) {
                    const dataTransfer = new DataTransfer();
                    dataTransfer.items.add(file);
                    fileInput.files = dataTransfer.files;
                }
            } else {
                alert('Please upload a valid Excel (.xlsx) file.');
            }
        }
    }

    removeFileBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.value = '';
        fileInfo.style.display = 'none';
        document.querySelector('.upload-content').style.display = 'block';
    });

    // Form Submission
    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        if (!fileInput.files.length) {
            alert('Please select a file first.');
            return;
        }



        // UI State: Loading
        form.style.display = 'none';
        statusArea.style.display = 'block';
        const statusText = document.getElementById('status-text');
        const individualDownloads = document.getElementById('individual-downloads');
        individualDownloads.innerHTML = ''; // Clear previous results

        const formData = new FormData(form);

        try {
            const response = await fetch(form.action, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'An unknown error occurred.');
            }

            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '';

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;

                buffer += decoder.decode(value, { stream: true });
                const lines = buffer.split('\n\n');
                buffer = lines.pop(); // Keep the last incomplete chunk

                for (const line of lines) {
                    if (line.startsWith('data: ')) {
                        const jsonStr = line.slice(6);
                        try {
                            const data = JSON.parse(jsonStr);

                            if (data.type === 'progress') {
                                statusText.textContent = data.message;
                            } else if (data.type === 'result') {
                                const btn = document.createElement('a');
                                btn.href = data.download_url;
                                btn.className = 'download-btn individual-btn';
                                btn.innerHTML = `<i class="fa-solid fa-download"></i> ${data.topic}`;
                                btn.style.display = 'block';
                                btn.style.marginBottom = '10px';
                                individualDownloads.appendChild(btn);
                                statusText.textContent = `Completed: ${data.topic} (${data.progress})`;
                            } else if (data.type === 'complete') {
                                statusArea.style.display = 'none';
                                resultArea.style.display = 'block';
                                downloadLink.href = data.download_url;
                            } else if (data.type === 'error') {
                                throw new Error(data.message);
                            }
                        } catch (e) {
                            console.error('Error parsing SSE data:', e);
                        }
                    }
                }
            }

        } catch (error) {
            statusArea.style.display = 'none';
            form.style.display = 'block';
            alert(`Error: ${error.message}`);
        }
    });

    // Reset
    resetBtn.addEventListener('click', () => {
        resultArea.style.display = 'none';
        form.style.display = 'block';
        form.reset();
        fileInput.value = '';
        fileInfo.style.display = 'none';
        document.querySelector('.upload-content').style.display = 'block';
        document.getElementById('individual-downloads').innerHTML = '';
    });
});
