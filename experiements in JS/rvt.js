document.getElementById('fileInput').addEventListener('change', function(event) {
    const files = event.target.files;
    for (let i = 0; i < files.length; i++) {
        processRvtFile(files[i]);
    }
});

function processRvtFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const fileInfo = extractRvtFileInfo(data);
        displayResult(file.name, fileInfo);
    };
    reader.readAsArrayBuffer(file);
}

function extractRvtFileInfo(data) {
    const decoder = new TextDecoder('ascii');
    const text = decoder.decode(data);

    // Extract information using regex, similar to the Python code
    const cmp = /Central Model Path: (.*)/.exec(text);
    const locale = /Locale when saved: (.*)/.exec(text);
    const build = /Build: (.*)/.exec(text);
    const rvtFormat = /Format: (.*)/.exec(text);

    return {
        decode: text,
        cmp: cmp ? cmp[1] : 'None',
        locale: locale ? locale[1] : 'None',
        build: build ? build[1] : 'None',
        format: rvtFormat ? rvtFormat[1] : 'None'
    };
}

function displayResult(fileName, fileInfo) {
    const outputDiv = document.getElementById('output');
    const result = `
        <h3>${fileName}</h3>
        <p>Central Model Path: ${fileInfo.cmp}</p>
        <p>Locale when saved: ${fileInfo.locale}</p>
        <p>Build Number: ${fileInfo.build}</p>
        <p>Format: ${fileInfo.format}</p>
    `;
    outputDiv.innerHTML += result;
}
