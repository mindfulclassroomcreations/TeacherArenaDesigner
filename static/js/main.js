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

    // Form Submission - Async with polling
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
        
        // Determine which endpoint to use based on form action
        const isAcademy = form.action.includes('academy');
        const asyncEndpoint = isAcademy ? '/generate-academy-async' : '/generate-caterpillar-async';

        try {
            // Start the background task
            statusText.textContent = 'Starting generation...';
            const response = await fetch(asyncEndpoint, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'Failed to start generation');
            }

            const { task_id, status_url } = await response.json();
            statusText.textContent = 'Generation in progress...';

            // Poll for progress
            const pollInterval = setInterval(async () => {
                try {
                    const statusResponse = await fetch(status_url);
                    const status = await statusResponse.json();

                    if (status.state === 'PENDING') {
                        statusText.textContent = 'Waiting to start...';
                    } else if (status.state === 'PROGRESS') {
                        const percent = status.total > 0 ? Math.round((status.current / status.total) * 100) : 0;
                        statusText.textContent = `${status.status} (${percent}%)`;

                        // Update individual downloads
                        if (status.individual_files && status.individual_files.length > 0) {
                            individualDownloads.innerHTML = ''; // Clear and rebuild
                            status.individual_files.forEach(file => {
                                const btn = document.createElement('a');
                                btn.href = file.download_url;
                                btn.className = 'download-btn individual-btn';
                                btn.innerHTML = `<i class="fa-solid fa-download"></i> ${file.topic}`;
                                btn.style.display = 'block';
                                btn.style.marginBottom = '10px';
                                individualDownloads.appendChild(btn);
                            });
                        }
                    } else if (status.state === 'SUCCESS') {
                        clearInterval(pollInterval);
                        statusText.textContent = 'Generation complete!';

                        // Show final results
                        statusArea.style.display = 'none';
                        resultArea.style.display = 'block';
                        
                        if (status.result && status.result.download_url) {
                            downloadLink.href = status.result.download_url;
                        }

                        // Show all individual files
                        if (status.result && status.result.individual_files) {
                            individualDownloads.innerHTML = '';
                            status.result.individual_files.forEach(file => {
                                const btn = document.createElement('a');
                                btn.href = file.download_url;
                                btn.className = 'download-btn individual-btn';
                                btn.innerHTML = `<i class="fa-solid fa-download"></i> ${file.topic}`;
                                btn.style.display = 'block';
                                btn.style.marginBottom = '10px';
                                individualDownloads.appendChild(btn);
                            });
                        }
                    } else if (status.state === 'FAILURE') {
                        clearInterval(pollInterval);
                        throw new Error(status.error || 'Task failed');
                    }
                } catch (pollError) {
                    clearInterval(pollInterval);
                    statusArea.style.display = 'none';
                    form.style.display = 'block';
                    alert(`Error: ${pollError.message}`);
                }
            }, 2000); // Poll every 2 seconds

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
