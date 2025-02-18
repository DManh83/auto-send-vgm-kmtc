async function submitData() {
    const blNo = document.getElementById('blNo').value;
    const file = document.getElementById('file').files[0];

    if(!blNo) {
      alert('BL No. is required');
      return;
    }
    
    try {
        const blNoResult = await window.electronAPI.submitButton(blNo, file?.path);
        // const fileResult = await window.electronAPI.uploadFile(file.path);
        
        if (blNoResult.success) {
            alert('Submitted successfully!');
        } else {
            alert('Submission failed!');
        }
    } catch (error) {
        alert('Error: ' + error.message);
    }
}

function resetData() {
    const file = document.getElementById('file');
    file.value = '';
    const blNo = document.getElementById('blNo');
    blNo.value = '';
    file.dispatchEvent(new Event('change'));
    blNo.dispatchEvent(new Event('change'));
}
