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

                        console.log(`there are ${sheets.length} sheets`)

                        sheets.map(sheet => {
                            if(sheet.toLowerCase() !== "template"){
                                sheet_results.push(readSheet(workbook.Sheets[sheet]))
                            }

                            console.log(`sheetname is ${sheet}`)
                        });
                    });

                    Promise.all(sheet_results)
                    .then(res => {
                        let compiled_results = [["Location", "Issue", "Downtime (s)", "Root Cause", "Reason", "Corrective Action", "Owner", "Shift", "Date", "Time"]];

                        res.map(sheet_entry => compiled_results.push(...sheet_entry));

                        return compiled_results
                        
                    })
                    .then(data => {
                        // saveFile(data)
                        const sort_by_leg = {}


                        console.log(data);
                        data.map((row, idx) => {
                            if(idx != 0){
                                let keys = Object.keys(sort_by_leg);

                                if(row[2] === undefined){
                                    console.log(row)  
                                }
                        
                                if(!keys.includes(row[0])){
                                    console.log('this')
                                    sort_by_leg[row[0]] = parseInt(row[2])
                                } else {
                                    sort_by_leg[row[0]] = sort_by_leg[row[0]] + parseInt(row[2])
                                };
                            };


                        });

                        $('#data').DataTable( { 
                            "aaData": data,
                            "aoColumns": [
                                {"title": "column1"},
                                {"title": "column2"},
                                {"title": "column3"},
                                {"title": "column1"},
                                {"title": "column2"},
                                {"title": "column3"},
                                {"title": "column1"},
                                {"title": "column2"},
                                {"title": "column3"},
                                {"title": "soemotiej"}
                                ]
                        }); 
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

            console.log(`start index is ${start_idx}`);

            for(i=start_idx; i < formatted_data.length; i++){
                if(formatted_data[i][2] === null || formatted_data[i][2] === undefined){
                    break;
                };

                let raw_row = formatted_data[i]

                console.log(formatted_data[i][2])

                let row = [raw_row[2],raw_row[3],cleanNumber(raw_row[5]),raw_row[6],raw_row[9], raw_row[12], raw_row[14], cleanNumber(raw_row[16]), raw_row[17], cleanNumber(raw_row[18])]
                let cleaned_row = row.map(elem => {
                    
                    if(elem == null || elem == undefined || elem == `empty`){
                        elem = "N/A"
                        
                    } else {
                        try {
                            elem = elem.toUpperCase()
                        } catch(error){
                            console.log(`element may be a number: ${error}\n${elem}`)
                        }
                        
                    }

                    return elem;
                });
                // let new_row = cleaned_row.filter(elem => {
                //     return elem !== undefined
                // });

               let new_row = cleaned_row.map(elem => {
                   try {
                    elem.trim()
                   } catch(err){
                       console.log(`may be a number: ${err}\n${elem}\n`)
                   }
                   return elem;
               })
                
                if(new_row.length !== 0){
                    sheet_rows.push(new_row)
                };
     
        };

        resolve(sheet_rows);
        } catch(error){
            reject(error)
        }
    });  
    return promise;
};

function saveFile(data){
    const book = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(book, sheet, 'EOS_compiled_data');
    XLSX.writeFile(book, `eosdata.xlsx`);
};

function cleanNumber(number){
    let invalid = [null, undefined, "-", "N/A"]

    if(invalid.includes(number)){
        number = 0
    } else if(number instanceof String){
        try {
            number = parseInt(number)
        } catch {
            number = 0
        }
    }

    return number;
}