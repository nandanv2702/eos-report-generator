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
                    const sheet_results = [];

                    data.forEach(workbook => {
                        let sheets = workbook.SheetNames;

                        sheets.map(sheet => {
                            console.log(`sheetname is ${sheet}`)
                            sheet_results.push(readSheet(workbook.Sheets[sheet]))
                        });
                    });

                    Promise.all(sheet_results).then(res => console.log(res))

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

async function readSheet(worksheet){
    const promise = new Promise((resolve, reject) => {

        try {
            const sheet_rows = [];

        console.log('readaing sheet');
        
        let formatted_data = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw:false});

        let start_idx;

        formatted_data.some((elem, index) => {
            if(elem == "AGC ISSUES"){
                start_idx = index + 2
            };
        });

        for(i=start_idx; i < formatted_data.length; i++){
            if(formatted_data[i][0] == null){
                break;
            };

            let raw_row = formatted_data[i]

            let row = [raw_row[2],raw_row[3],raw_row[5],raw_row[6],raw_row[9], raw_row[12], raw_row[14], raw_row[16], raw_row[17], raw_row[18]]
            row.map(elem => (elem == null || undefined) ? elem = "N/A" : elem = elem.toUpperCase())
            
            sheet_rows.push(row)
            
        };

        resolve(sheet_rows);
        } catch(error){
            reject(error)
        }
    });  
    return promise;
};