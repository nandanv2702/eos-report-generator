// Requirements:

// when multiple files are submitted, call a function that returns an array of Promises with the whole file 

// then, once all thoes Promises are resolved, for each workbook get the sheet data and push it into a new array

// send the new array to the viewer

console.log("script two enabled");

// Polling for the sake of my intern tests
var interval = setInterval(function() {
    if(document.readyState === 'complete') {

        document.getElementById("files").addEventListener("change", e => {
            handleFiles(e)
            .then(res => {
                Promise.all(res).then(data => {
                    
                    data.forEach(workbook => {
                        let sheets = workbook.SheetNames
                    });

                });
            });

        }, false);

        clearInterval(interval);
    }
}, 100);

async function handleFiles(e){
    const file_promises = [];

    let fileList = e.target.files, f=fileList[0];

    for(var i = 0; i < fileList.length; i++){

        let file = fileList[i]
        const promise = new Promise((resolve, reject) => {
            const fileReader = new FileReader();
            fileReader.readAsArrayBuffer(file);
            fileReader.onload = (e) => {
                console.log("Here", e.target.result);

                let data = new Uint8Array(fileReader.result);
                let workbook = XLSX.read(data, {type: 'array'});
    
                resolve(workbook);
            };
        });

        file_promises.push(promise);

    };



    return file_promises;
    // returns an array that's a list of Promises ==> each workbook is a Promise object

};

function readWorkbook(workbook){

};