async function submitData() {
    const blNo = document.getElementById('blNo').value;
    const file = document.getElementById('file').files[0];
    
    try {
        const blNoResult = await window.electronAPI.submitBlNo(blNo);
        const fileResult = await window.electronAPI.uploadXlsx(file.path);
        
        if (blNoResult.success && fileResult.success) {
            alert('Submitted successfully!');
        } else {
            alert('Submission failed!');
        }
    } catch (error) {
        alert('Error: ' + error.message);
    }
}