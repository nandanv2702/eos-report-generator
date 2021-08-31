// Requirements:

// when multiple files are submitted, call a function that returns an array of Promises with the whole file 

// then, once all thoes Promises are resolved, for each workbook get the sheet data and push it into a new array

// send the new array to the viewer

let dataTableSet = false;
let global_res = [];

var interval = setInterval(function () {
    if (document.readyState === 'complete') {

        document.getElementById("download_xlsx").addEventListener("click", () => {
            saveFile(global_res)
        })

        document.getElementById("files").addEventListener("change", e => {

            document.getElementById("download_xlsx").style.display = ""

            // destroy datatable
            document.getElementById("data").innerHTML = ""

            handleFiles(e)
                .then(res => {
                    document.getElementById("loading").style.display = 'flex'
                    Promise.all(res).then(data => {
                        const sheet_results = [];

                        data.forEach(workbook => {
                            let sheets = workbook.SheetNames;

                            console.log(`there are ${sheets.length} sheets`)

                            sheets.map(sheet => {
                                if (sheet.toLowerCase() !== "template") {
                                    sheet_results.push(readSheet(workbook.Sheets[sheet]))
                                }

                                console.log(`sheetname is ${sheet}`)
                            });
                        });

                        Promise.all(sheet_results)
                            .then(res => {
                                let compiled_results = [];

                                res.map(sheet_entry => compiled_results.push(...sheet_entry));

                                return compiled_results

                            })
                            .then(res => {
                                global_res = res;
                                const sort_by_leg = {}


                                console.log(res);
                                res.map((row, idx) => {
                                    if (idx != 0) {
                                        let keys = Object.keys(sort_by_leg);

                                        if (row[2] === undefined) {
                                            console.log(row)
                                        }

                                        if (!keys.includes(row[0])) {
                                            console.log('this')
                                            sort_by_leg[row[0]] = parseInt(row[2])
                                        } else {
                                            sort_by_leg[row[0]] = sort_by_leg[row[0]] + parseInt(row[2])
                                        };
                                    };


                                });

                                // sort station by highest downtime and render top 8 in pie chart
                                // Create items array
                                var items = Object.keys(sort_by_leg).map(function (key) {
                                    return [key, sort_by_leg[key]];
                                });

                                // Sort the array based on the second element
                                items.sort(function (first, second) {
                                    return second[1] - first[1];
                                });


                                // Create a new array with only the first 8 items
                                console.log(items.slice(0, 8));

                                let final_render = items.slice(0, 8);

                                // Calculate total downtime
                                let total_downtime = items.reduce((prevVal, currVal) => {
                                    return prevVal + parseInt(currVal[1])
                                }, 0)

                                // get percentage downtime for each of the top 8 legs
                                let percentages = final_render
                                    .map(val => {
                                        return ((parseInt(val[1]) / parseInt(total_downtime) * 100).toString().slice(0, 4) + "%")
                                    })

                                // deletes old chart (if exists)
                                deleteChart()

                                // creates new chart from data
                                createChart(final_render, percentages)

                                console.log(`FINAL DATASET IS \n ${JSON.stringify(res)}`)

                                if (dataTableSet) {
                                    $("#data").dataTable().fnDestroy()
                                };

                                $('#data').DataTable({
                                    "aaData": res,
                                    "aoColumns": [
                                        { "title": "Location" },
                                        { "title": "Issue" },
                                        { "title": "Downtime (s)" },
                                        { "title": "Root Cause" },
                                        { "title": "Reason" },
                                        { "title": "Corrective Action" },
                                        { "title": "Owner" },
                                        { "title": "Shift" },
                                        { "title": "Date" },
                                        { "title": "Time" }
                                    ],
                                    order: [[2, 'desc']]
                                });

                                dataTableSet = true

                                const uniqueLocs = new Set(res.map(row => row[0]))

                                console.log(uniqueLocs)

                                const pivotTableRes = res.map((
                                    [Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time]
                                ) => ({
                                    Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time
                                }))

                                pivotTableRes.map(row => {
                                    areaMapper(row)
                                    row['Issue'].trim()
                                    row['Reason'].trim()
                                    row['RootCause'].trim()

                                    Object.defineProperty(row, 'Downtime (s)',
                                        Object.getOwnPropertyDescriptor(row, 'Downtime'));
                                    delete row['Downtime'];
                                })

                                console.log();

                                var derivers = $.pivotUtilities.derivers;
                                var renderers = $.extend($.pivotUtilities.renderers,
                                    $.pivotUtilities.c3_renderers);

                                $("#output").pivotUI(pivotTableRes, {
                                    renderers: renderers,
                                    cols: ["Leg Name"],
                                    rows: ["Issue"],
                                    rowOrder: "value_a_to_z",
                                    colOrder: "value_z_to_a",
                                    aggregatorName: "Sum",
                                    rendererName: "Stacked Bar Chart",
                                    rendererOptions: {
                                        c3: {
                                            tooltip: {
                                                show: true
                                            }
                                        }
                                    }

                                });

                                document.getElementById("loading").style.display = 'none'

                            });
                    });
                });

        }, false);

        clearInterval(interval);
    }
}, 100);

async function handleFiles(e) {
    const file_promises = [];

    let fileList = e.target.files, f = fileList[0];

    for (var i = 0; i < fileList.length; i++) {

        let file = fileList[i]
        const promise = new Promise((resolve, reject) => {
            const fileReader = new FileReader();
            fileReader.readAsArrayBuffer(file);
            fileReader.onload = (e) => {
                console.log("Here", e.target.result);

                let data = new Uint8Array(fileReader.result);
                let workbook = XLSX.read(data, { type: 'array' });

                resolve(workbook);
            };
        });

        file_promises.push(promise);

    };

    return file_promises;

};

async function readSheet(worksheet) {
    const promise = new Promise((resolve, reject) => {

        try {
            const sheet_rows = [];

            console.log('readaing sheet');

            let formatted_data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            let start_idx;

            formatted_data.some((elem, index) => {
                if (elem[0] == "AGC ISSUES") {
                    start_idx = index + 2
                };
            });

            console.log(`start index is ${start_idx}`);

            for (i = start_idx; i < formatted_data.length; i++) {
                if (formatted_data[i][2] === null || formatted_data[i][2] === undefined) {
                    break;
                };

                let raw_row = formatted_data[i]

                console.log(formatted_data[i][2])

                let row = [raw_row[2], raw_row[3], cleanNumber(raw_row[5]), raw_row[6], raw_row[9], raw_row[12], raw_row[14], cleanNumber(raw_row[16]), raw_row[17], cleanNumber(raw_row[18])]
                let cleaned_row = row.map(elem => {

                    if (elem == null || elem == undefined || elem == `empty`) {
                        elem = "N/A"

                    } else {
                        try {
                            elem = elem.toUpperCase()
                        } catch (error) {
                            console.log(`element may be a number: ${error}\n${elem}`)
                        }

                    }

                    return elem;
                });

                let new_row = cleaned_row.map(elem => {
                    try {
                        elem.trim()
                    } catch (err) {
                        console.log(`may be a number: ${err}\n${elem}\n`)
                    }
                    return elem;
                })

                if (new_row.length !== 0) {
                    sheet_rows.push(new_row)
                };

            };

            resolve(sheet_rows);
        } catch (error) {
            reject(error)
        }
    });
    return promise;
};

function saveFile(data) {
    const book = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(book, sheet, 'EOS_Compiled_Data');
    XLSX.writeFile(book, `EOS_Report_RawData.xlsx`);
};

function cleanNumber(number) {
    let cleaned_num = parseInt(number, 10);

    if (Number.isNaN(cleaned_num)) {
        return 0;
    };

    return cleaned_num;
}

function deleteChart() {
    document.getElementById("myChart").remove();
}

function createChart(final_render, percentages) {
    let canvas = document.createElement("canvas");
    canvas.setAttribute("id", "myChart");
    document.querySelector("#chart_holder").appendChild(canvas);

    let ctx = document.getElementById("myChart");

    chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: final_render.map(loc => loc[0]),
            datasets: [{
                data: final_render.map(loc => loc[1]),
                backgroundColor: [
                    '#DCD106',
                    '#D1DA9F',
                    '#E35184',
                    '#FAD7FF',
                    '#6ECC49',
                    '#ED2B34',
                    '#D1DA9F',
                    '#839C99'
                ],
                hoverOffset: 4,
            }]
        },
        options: {
            plugins: {
                legend: false,
                title: {
                    display: true,
                    text: "Station vs. Downtime (s)",
                    fullsize: true
                },
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: "Downtime (s)",
                        fullsize: true
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: "Station",
                        fullsize: true
                    }
                }
            },
            hover: {
                animationDuration: 1
            },
            animation: {
                onComplete: () => {
                    let chartInstance = this.chart,
                        ctx = chartInstance.ctx;
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'bottom';

                    chartInstance.data.datasets.forEach(function (dataset, i) {
                        let meta = chartInstance.getDatasetMeta(i)
                        meta.data.forEach(function (bar, index) {
                            ctx.fillText(percentages[index], bar.x, bar.y - 5);
                        });
                    });
                }
            }
        },
    });

}

function areaMapper(row) {
    const mapper = {
        "B01": "B-Leg",
        "FA2": "Final Assembly",
        "A06": "A-Leg",
        "B04": "B-Leg",
        "BE2": "E-Leg",
        "A07": "A-Leg",
        "E07": "E-Leg",
        "B06": "B-Leg",
        "A08": "A-Leg",
        "B12": "B-Leg",
        "H10": "H-Leg",
        "BD2": "D-Leg",
        "E12": "E-Leg",
        "A01": "A-Leg",
        "BB1": "B-Leg",
        "L1FL": "FL",
        "BA3": "A-Leg",
        "PC1": "PC",
        "B03": "B-Leg",
        "D08": "D-Leg",
        "B11": "B-Leg",
        "C10": "C-Leg",
        "D09": "D-Leg",
        "D01": "D-Leg",
        "B05": "B-Leg",
        "B10": "B-Leg",
        "B08": "B-Leg",
        "B09": "B-Leg",
        "BC3": "B-Leg",
        "B02": "B-Leg",
        "H09": "H-Leg",
        "D06": "D-Leg",
        "C12": "C-Leg",
        "A09": "A-Leg",
        "BDE1": "D-Leg",
        "A11": "A-Leg",
        "C05": "C-Leg",
        "D11": "D-Leg",
        "C08": "C-Leg",
        "C06": "C-Leg",
        "A05": "A-Leg",
        "B07": "B-Leg",
        "A03": "A-Leg",
        "BDE3": "D-Leg",
        "H05": "H-Leg",
        "H03": "H-Leg",
        "A12": "A-Leg",
        "D02": "D-Leg",
        "BF6": "F-Leg",
        "C04": "C-Leg",
        "C11": "C-Leg",
        "RT": "RT",
        "G05": "G-Leg",
        "E01": "E-Leg",
        "D04": "D-Leg",
        "G06": "G-Leg",
        "G03": "G-Leg",
        "H06": "H-Leg",
        "H12": "H-Leg",
        "A02": "A-Leg",
        "BE3": "E-Leg",
        "LD01": "LD",
        "C01": "C-Leg",
        "PN01": "PN",
        "G04": "G-Leg",
        "H04": "H-Leg",
        "H08": "H-Leg",
        "H02": "H-Leg",
        "MINOR B": "Minor B",
        "FA": "Final Assembly",
        "D07": "D-Leg",
        "G02": "G-Leg",
        "E08": "E-Leg",
        "E11": "E-Leg",
        "BUFFERL2": "Buffer ",
        "C07": "C-Leg",
        "H07": "H-Leg",
        "BB3": "B-Leg",
        "A04": "A-Leg",
        "D10": "D-Leg",
        "C09": "C-Leg",
        "C02": "C-Leg",
        "TL02": "TL",
        "BD3": "D-Leg",
        "BE1": "E-Leg",
        "BA1": "A-Leg",
        "G01": "G-Leg",
        "H01": "H-Leg",
        "PC": "PC",
        "D03": "D-Leg",
        "A10": "A-Leg",
        "TL01": "TL",
        "BL02": "BL",
        "D05": "D-Leg",
        "C03": "C-Leg",
        "BFA2": "Final Assembly",
        "BC1": "C-Leg",
        "E02": "E-Leg",
        "E03": "E-Leg",
        "BFA": "Final Assembly",
        "D10": "D-Leg",
        "BF2": "F-Leg",
        "H11": "H-Leg",
        "FA1": "Final Assembly",
        "TL BUFFER": "TL",
        "PC2": "PC",
        "RA7": "RA",
        "D12": "D-Leg",
        "BB2": "B-Leg",
        "BF1": "F-Leg",
        "TL2": "TL",
        "FRAME LOAD": "Frame Load",
        "BE02": "E-Leg",
        "RT2": "RT",
        "VIN": "VIN",
        "BE01": "E-Leg",
        "TL": "TL",
        "BD1": "D-Leg",
        "H-LEG": "H-Leg",
        "RT1": "RT",
        "BC01": "C-Leg",
        "FLEG": "F-Leg",
        "BB01": "B-Leg",
        "BRT": "RT",
        "BF3": "F-Leg",
        "A08": "A-Leg",
        "LD02": "LD",
        "E10": "E-Leg",
        "BF4": "F-Leg",
        "BCR": "C-Leg",
        "BTL": "TL",
        "BH01": "H-Leg",
        "BPC": "PC",
        "RT3": "RT",
        "700": "700",
        "RA": "RA",
        "BL01": "BL01",
        "704": "704",
        "LOAD LANE": "LOAD LANE",
        "F LEG": "f-leg",
        "VDS #1": "VDS",
        "BD02": "D-Leg",
        "BC03": "C-Leg",
        "AB05": "B-Leg",
        "VN01": "VN",
        "OFL": "OFL",
        "BA01": "A-Leg",
        "MINOR B1": "Minor B",
        "RTN": "RTN",
        "B10": "B-Leg",
        "#719": "719",
        "VIN LOAD": "VIN",
        "B07L": "B-Leg",
        "BC2": "C-Leg",
        "VDS": "VDS",
        "RT4": "RT",
        "D4": "D-Leg",
        "A7": "A-Leg",
        "BCR1": "C-Leg",
        "BE03": "E-Leg",
        "RA5": "RA",
        "BB!": "B-Leg",
        "UNLOAD": "Unload",
        "BHI": "H-Leg",
        "BY": "BY",
        "BH1": "H-Leg",
        "VMO1": "VM",
        "E2": "E-Leg",
        "RT6": "RT",
        "LT2": "LT",
        "C1BUFF": "Buffer",
        "B11L": "B-Leg",
        "LT01": "LT",
        "BC": "BC",
        "LT1": "LT",
        "VINLOAD": "VIN",
        "BFA1": "Final Assembly",
        "VIN OAD": "VIN",
        "H04": "H-leg",
        "L2 LOAD": "LOAD LANE",
        "RA2": "RA",
        "XF04": "F-Leg",
        "BH03": "H-Leg",
        "CBC3": "C-Leg",
        "BA2": "A-Leg",
        "BUFFER B LEG": "Buffer",
        "TL1": "TL",
        "E09": "E-Leg",
        "FA (#468)": "Final Assembly",
        "L2FL": "FL",
        "D LEG": "D-Leg",
        "BA1 (CHARGERS)": "A-Leg",
        "BRB": "RB",
        "789": "789",
        "BL1A": "A-Leg",
        "RA1": "RA",
        "RA4": "RA",
        "LINE 2": "Line 2",

    }

    row['Leg Name'] = mapper[row['Location']]

    return row
}