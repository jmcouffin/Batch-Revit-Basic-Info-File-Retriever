document.getElementById('fileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (!file) {
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        findBasicFileInfo(data);
    };
    reader.readAsArrayBuffer(file);
});

function findBasicFileInfo(data) {
    const decoder = new TextDecoder('utf-16le');
    const text = decoder.decode(data);
    const searchString = 'BasicFileInfo';
    const index = text.indexOf(searchString);

    if (index > -1) {
        const infoStartIndex = index + searchString.length;
        const infoString = text.substring(infoStartIndex, infoStartIndex + 13);
        displayResult(infoString, index);
    } else {
        displayResult('BasicFileInfo stream not found.');
    }
}

function displayResult(message, index = '') {
    document.getElementById('output').textContent = `Result: ${message} ${index !== '' ? `at offset ${index}` : ''}`;
}
