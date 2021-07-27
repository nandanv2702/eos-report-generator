console.log("enabled");

// Polling for the sake of my intern tests
var interval = setInterval(function() {
    if(document.readyState === 'complete') {

        document.getElementById("files").addEventListener("change", e => {
            console.log(handleFile(e))
            
            

        }, false);

        clearInterval(interval);
    }
}, 100);

function handleFile(e) {
    const rows = [];
    console.log("file in");
    let fileList = e.target.files, f = fileList[0];
    if(fileList.length > 0){
        console.log(fileList);

        let headers = new Promise((resolve, reject) => resolve(["Location", "Issue", "Downtime (s)", "Root Cause", "Reason", "Corrective Action", "Owner", "Shift", "Date", "Time"]));

        rows.push(headers)

        for(var i = 0; i < fileList.length; i++){
            const file = fileList[i]
            console.log('reading wb')
            
            rows.push(readWorkbook(file, f));
        };
    };
    return rows;
};

function readWorkbook(file, f){
    let reader = new FileReader();

    const workbook_rows = [];

    reader.onload = function() {
        let data = new Uint8Array(reader.result);
        workbook = XLSX.read(data, {type: 'array'});

        let sheets = workbook.SheetNames;

        console.log(workbook)

        // sheets.forEach(sheet => {
        //     if(sheet.toLowerCase() !== "template"){
        //         ws = workbook.Sheets[sheet];

        //         workbook_rows.push(readSheet(ws));

        //     }

        // });
    };

    return reader.readAsArrayBuffer(f);


};

function readSheet(ws){

    const sheet_rows = [];

    console.log('readaing sheet')
    let formatted_data = XLSX.utils.sheet_to_json(ws, {header: 1, raw:false});

    let start_idx;

    formatted_data.some((elem, index) => {
        if(elem == "AGC ISSUES"){
            start_idx = index + 2
        }
    });

    for(i=start_idx; i < formatted_data.length; i++){
        if(formatted_data[i][0] == null){
            break;
        }

        let raw_row = formatted_data[i]

        let row = [raw_row[2],raw_row[3],raw_row[5],raw_row[6],raw_row[9], raw_row[12], raw_row[14], raw_row[16], raw_row[17], raw_row[18]]
        row.map(elem => (elem == null || undefined) ? elem = "N/A" : elem = elem.toUpperCase())
        
        sheet_rows.push(row)
        
    };

    return sheet_rows;
};

function saveFile(data){
    console.log(data)
};