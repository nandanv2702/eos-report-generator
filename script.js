'use strict';

let dataTableSet = false;
let global_downtime_result = [];

var interval = setInterval(function () {
    if (document.readyState === 'complete') {

        clearInterval(this)

        document.getElementById("download_xlsx").addEventListener("click", () => {
            console.log(global_downtime_result)
            saveFile(global_downtime_result)
        })

        document.getElementById("files").addEventListener("change", async (e) => {

            document.getElementById("download_xlsx").style.display = ""

            // destroy datatable
            document.getElementById("data").innerHTML = ""

            const res = await handleFiles(e)
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
                        const downtimeResults = []
                        const cartResults = []

                        res.map(sheet_entry => {
                            downtimeResults.push(...sheet_entry.downtimeEntries)
                            cartResults.push(...sheet_entry.cartIssues)
                        });

                        global_downtime_result = downtimeResults;
                        const sort_by_leg = {}

                        console.log(downtimeResults);
                        downtimeResults.map((row, idx) => {
                            if (idx != 0) {
                                let keys = Object.keys(sort_by_leg);

                                if (row[3] === undefined) {
                                    console.log(row)
                                }

                                if (!keys.includes(row[1])) {
                                    console.log('this')
                                    sort_by_leg[row[1]] = parseInt(row[3])
                                } else {
                                    sort_by_leg[row[1]] = sort_by_leg[row[1]] + parseInt(row[3])
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

                        downtimeResults.forEach(row => {
                            row[9] = new Date(row[9])
                        })

                        console.log(`FINAL DATASET IS \n ${JSON.stringify(downtimeResults)}`)

                        if (dataTableSet) {
                            $("#data").dataTable().fnDestroy()
                        };

                        $('#data').DataTable({
                            "aaData": downtimeResults,
                            "aoColumns": [
                                { "title": "Cart" },
                                { "title": "Location" },
                                { "title": "Issue" },
                                { "title": "Downtime (s)" },
                                { "title": "Root Cause" },
                                { "title": "Reason" },
                                { "title": "Corrective Action" },
                                { "title": "Owner" },
                                { "title": "Shift" },
                                {
                                    "title": "Date",
                                    "type": "date",
                                    "render": function (value) {

                                        var dt = new Date(value);

                                        return (dt.getMonth() + 1) + "/" + dt.getDate() + "/" + dt.getFullYear();
                                    }
                                },
                                { "title": "Time" }
                            ],
                            order: [[3, 'desc']]
                        });

                        dataTableSet = true

                        const uniqueLocs = new Set(downtimeResults.map(row => row[1]))

                        console.log(uniqueLocs)

                        const pivotTableRes = downtimeResults.map((
                            [Cart, Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time]
                        ) => ({
                            Cart, Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time
                        }))

                        const dates = []

                        pivotTableRes.map(row => {
                            areaMapper(row)
                            row['Cart'].trim()
                            row['Issue'].trim()
                            row['Reason'].trim()
                            row['RootCause'].trim()

                            Object.defineProperty(row, 'Downtime (s)',
                                Object.getOwnPropertyDescriptor(row, 'Downtime'));
                            delete row['Downtime'];

                            row['Week'] = getWeek(new Date(row['Date']))
                        })

                        console.log(`dates are ${JSON.stringify(dates)}`);

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
                            hiddenAttributes: ["Time"],
                            rendererName: "Stacked Bar Chart",
                            rendererOptions: {
                                c3: {
                                    tooltip: {
                                        show: true
                                    },
                                }
                            }

                        });

                        document.getElementById("loading").style.display = 'none'

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

function findStarterElements(formatted_data, searchText) {
    let startIndexArray = []
    formatted_data.filter((elem, index) => {
        if (elem[0] === searchText.toString()) {
            startIndexArray.push([index, ...elem])
        }
    })
    return startIndexArray
}

async function readDowntimeEntries(formatted_data, startIndexArray) {
    const sheet_rows = []

    let start_idx = startIndexArray[1][0] + 2

    console.log(`start index is ${start_idx}`);

    for (let i = start_idx; i < formatted_data.length; i++) {
        if (formatted_data[i][2] === null || formatted_data[i][2] === undefined) {
            break;
        };

        let raw_row = formatted_data[i]

        console.log(formatted_data[i][2])

        let row = [raw_row[0], raw_row[2], raw_row[3], cleanNumber(raw_row[5]), raw_row[6], raw_row[9], raw_row[12], raw_row[14], cleanNumber(raw_row[16]), raw_row[17], cleanNumber(raw_row[18])]
        let cleaned_row = row.map(elem => {

            if (elem === null || elem === undefined || elem == `empty`) {
                elem = "N/A"

            } else {
                try {
                    elem = elem.toUpperCase().trim()
                } catch (error) {

                    // console.log(`element may be a number: ${error}\n${elem}`)
                }

            }

            return elem;
        });

        if (cleaned_row.length !== 0) {
            sheet_rows.push(cleaned_row)
        };

    };

    return sheet_rows
}

async function readCartIssues(formatted_data, startIndexArray) {
    const sheet_rows = []

    let start_idx = startIndexArray[0][0] + 2

    console.log(`start index is ${start_idx}`);

    for (let i = start_idx; i < formatted_data.length; i++) {
        if (formatted_data[i][0] === null || formatted_data[i][0] === undefined) {
            break;
        };

        let raw_row = formatted_data[i]

        let row = [raw_row[0], raw_row[3], raw_row[17]]
        let cleaned_row = row.map(elem => {

            if (elem === null || elem === undefined || elem == `empty`) {
                elem = "N/A"

            } else {
                try {
                    elem = elem.toUpperCase().trim()
                } catch (error) {

                    // console.log(`element may be a number: ${error}\n${elem}`)
                }

            }

            return elem;
        });

        if (cleaned_row.length !== 0) {
            sheet_rows.push(cleaned_row)
            console.log(cleaned_row);
        };
    }

    return sheet_rows
}

async function readSheet(worksheet) {
    return new Promise(async (resolve, reject) => {

        try {
            const sheet_rows = {};

            console.log('readaing sheet');

            let formatted_data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            const startIndexArray = findStarterElements(formatted_data, "AGC ISSUES")

            sheet_rows.downtimeEntries = await readDowntimeEntries(formatted_data, startIndexArray)
            sheet_rows.cartIssues = await readCartIssues(formatted_data, startIndexArray)


            resolve(sheet_rows);
        } catch (error) {
            reject(error)
        }
    });
};

function saveFile(data) {
    data.unshift(["Cart", "Location", "Issue", "Downtime", "RootCause", "Reason", "CorrectiveAction", "Owner", "Shift", "Date", "Time"])
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

    let chart = new Chart(ctx, {
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
                    let chartInstance = chart,
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

function getWeek(date, dowOffset = 0) {
    /*getWeek() was developed by Nick Baicoianu at MeanFreePath: http://www.meanfreepath.com */
    const newYear = new Date(date.getFullYear(), 0, 1);
    let day = newYear.getDay() - dowOffset; //the day of week the year begins on
    day = (day >= 0 ? day : day + 7);
    const daynum = Math.floor((date.getTime() - newYear.getTime() -
        (date.getTimezoneOffset() - newYear.getTimezoneOffset()) * 60000) / 86400000) + 1;
    //if the year starts before the middle of a week
    if (day < 4) {
        const weeknum = Math.floor((daynum + day - 1) / 7) + 1;
        if (weeknum > 52) {
            const nYear = new Date(date.getFullYear() + 1, 0, 1);
            let nday = nYear.getDay() - dowOffset;
            nday = nday >= 0 ? nday : nday + 7;
            /*if the next year starts before the middle of
              the week, it is week #1 of that year*/
            return nday < 4 ? 1 : 53;
        }
        return weeknum;
    }
    else {
        return Math.floor((daynum + day - 1) / 7);
    }
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