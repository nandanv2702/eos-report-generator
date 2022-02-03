'use strict';

let dataTableSet = false;
let globalDowntimeResults = [];
let globalCartResults = [];
const invalidSheetNames = ["template", "dropdown list", "drop down list", "drop down"]

var interval = setInterval(function () {
    if (document.readyState === 'complete') {

        clearInterval(interval)

        document.getElementById("download_xlsx").addEventListener("click", () => {
            console.log(globalDowntimeResults)
            saveFile(globalDowntimeResults, globalCartResults)
        })

        document.getElementById("files").addEventListener("change", async (e) => {

            await main(e);

        }, false);

        clearInterval(interval);
    }
}, 100);


/**
 * Reads Excel workbooks, gets and sorts data, draws charts, draws pivottables
 * @param {Event} e - file/(s) upload
 */
async function main(e) {
    document.getElementById("download_xlsx").style.display = "";

    // destroy datatable
    document.getElementById("data").innerHTML = "";
    document.getElementById("loading").style.display = 'flex';

    const fileContents = readFiles(e);


    Promise.all(fileContents).then(data => {
        const sheet_results = [];

        data.forEach(workbook => {
            readWorkbook(workbook, sheet_results);
        });

        Promise.all(sheet_results)
            .then(sheetResults => {
                const { downtimeResults, cartResults } = setGlobalResults(sheetResults);

                const downtimeByLeg = sortResultsByStation(downtimeResults);

                // sort station by highest downtime and render top 8 in pie chart
                // Create items array
                let { topEightStations, percentages } = getDowntimeStatistics(downtimeByLeg);

                // deletes old chart (if exists)
                deleteChart();

                // creates new chart from data
                createChart(topEightStations, percentages);

                refactorDates(downtimeResults);

                console.log(`FINAL DATASET IS \n ${JSON.stringify(downtimeResults)}`);

                if (dataTableSet) {
                    $("#data").dataTable().fnDestroy();
                };

                createDowntimeDatatable(downtimeResults);

                // var derivers = $.pivotUtilities.derivers;
                var renderers = $.extend($.pivotUtilities.renderers,
                    $.pivotUtilities.c3_renderers);

                createDowntimePivotTable(downtimeResults, renderers);

                createCartIssuesPivotTable(cartResults, renderers);

                document.getElementById("loading").style.display = 'none';
                document.getElementById("outputarea").hidden = false;

            });
    });
}

/**
 * Reads donwtimeEntries from 
 * @param {Array<Array<String>>} sheetResults 
 * @returns {Array<Array<String>>} downtimeResults, cartResults
 */
function setGlobalResults(sheetResults) {
    const downtimeResults = [];
    const cartResults = [];

    sheetResults.map(sheetEntry => {
        downtimeResults.push(...sheetEntry.downtimeEntries);
        cartResults.push(...sheetEntry.cartIssues);
    });

    globalDowntimeResults = [...downtimeResults];
    globalCartResults = [...cartResults];
    return { downtimeResults, cartResults };
}

function readWorkbook(workbook, sheet_results) {
    let sheets = workbook.SheetNames;

    console.log(`there are ${sheets.length} sheets`);

    sheets.map(sheetName => {
        if (!invalidSheetNames.includes(sheetName.trim().toLowerCase())) {
            sheet_results.push(readSheet(workbook.Sheets[sheetName]));
        }

        console.log(`sheetname is ${sheetName}`);
    });
}

function createCartIssuesPivotTable(cartResults, renderers) {
    const cartIssuesPivotResults = cartResults.map((
        [Cart, Description, WorkOrderNumber]
    ) => (
        {
            Cart,
            Description,
            WorkOrderNumber
        }
    ));

    $("#cartIssuesPivotTable").pivotUI(cartIssuesPivotResults, {
        renderers: renderers,
        rows: ["Cart"],
        rowOrder: "value_z_to_a",
        colOrder: "value_a_to_z",
        aggregatorName: "Count",
        rendererName: "Col Heatmap",
        rendererOptions: {
            c3: {
                tooltip: {
                    show: true
                },
            }
        }
    });
}

function createDowntimePivotTable(downtimeResults, renderers) {
    const downtimePivotResults = downtimeResults.map((
        [Cart, Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time]
    ) => ({
        Cart, Location, Issue, Downtime, RootCause, Reason, CorrectiveAction, Owner, Shift, Date, Time
    }));

    const dates = [];

    downtimePivotResults.map(row => {
        areaMapper(row);
        row['Cart'].trim();
        row['Issue'].trim();
        row['Reason'].trim();
        row['RootCause'].trim();

        Object.defineProperty(row, 'Downtime (s)',
            Object.getOwnPropertyDescriptor(row, 'Downtime'));
        delete row['Downtime'];

        row['Week'] = getWeek(new Date(row['Date']));
    });

    console.log(`dates are ${JSON.stringify(dates)}`);

    $("#downtimePivotTable").pivotUI(downtimePivotResults, {
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
}

function refactorDates(downtimeResults) {
    downtimeResults.forEach(row => {
        row[9] = new Date(row[9]);
    });
}

function getDowntimeStatistics(downtimeByLeg) {
    var items = Object.keys(downtimeByLeg).map(function (key) {
        return [key, downtimeByLeg[key]];
    });

    // Sort the array based on the second element
    items.sort(function (first, second) {
        return second[1] - first[1];
    });


    // Create a new array with only the first 8 items
    console.log(items.slice(0, 8));

    let topEightStations = items.slice(0, 8);

    // Calculate total downtime
    let total_downtime = items.reduce((prevVal, currVal) => {
        return prevVal + parseInt(currVal[1]);
    }, 0);

    // get percentage downtime for each of the top 8 legs
    let percentages = topEightStations
        .map(val => {
            return ((parseInt(val[1]) / parseInt(total_downtime) * 100).toString().slice(0, 4) + "%");
        });

    return { topEightStations, percentages };
}

function sortResultsByStation(downtimeResults) {
    const downtimeByLeg = {};

    console.log(downtimeResults);
    downtimeResults.map((row, idx) => {
        if (idx != 0) {
            let keys = Object.keys(downtimeByLeg);

            if (row[3] === undefined) {
                console.log(row);
            }

            if (!keys.includes(row[1])) {
                console.log('this');
                downtimeByLeg[row[1]] = parseInt(row[3]);
            } else {
                downtimeByLeg[row[1]] = downtimeByLeg[row[1]] + parseInt(row[3]);
            };
        };
    });
    console.log(`SORTED BY LEG IS ${JSON.stringify(downtimeByLeg)}`);
    return downtimeByLeg;
}

function createDowntimeDatatable(downtimeResults) {
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
}

/**
 * Reads uploaded files as ArrayBuffer, converts them to XLSX workbooks using SheetJS, returns an array of promises
 * @param {Event} e 
 * @returns {Promise<Promise<Array>>} Array of Promises, where each promise is an Excel workbook
 */
function readFiles(e) {
    const file_promises = [];

    let fileList = e.target.files, f = fileList[0];

    for (var i = 0; i < fileList.length; i++) {

        let file = fileList[i]
        const promise = new Promise((resolve, reject) => {
            const fileReader = new FileReader();
            fileReader.readAsArrayBuffer(file);
            fileReader.onload = () => {
                try {
                    let data = new Uint8Array(fileReader.result);
                    let workbook = XLSX.read(data, { type: 'array' });

                    resolve(workbook);
                } catch (err) {
                    reject([])
                }
            };
        });

        file_promises.push(promise);

    };

    return file_promises;

};

/**
 * Finds all matches for the left-most cell in one worksheet with the specified searchText
 * @param {Array<Array<String>>} formattedData - cells in one worksheet
 * @param {String} searchText - the elements you want to find in the worksheet
 * @returns {Array<Array<[Number, String]>>}
 */
function findStarterElements(formattedData, searchText) {
    console.log(`FILTEREDDATA IS ${JSON.stringify(formattedData)}`);
    let startIndexArray = []
    formattedData.filter((elem, index) => {
        if (elem[0] === searchText.toString()) {
            startIndexArray.push([index, ...elem])
        }
    })
    return startIndexArray
}

/**
 * Converts null, undefined, 'empty' (from SheetJS) to "N/A", converts all text to uppercase and trims whitespace
 * @param {Array<String>} row 
 * @returns {Array<String>}
 */
function getCleanedRow(row) {
    return row.map(elem => {

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
}

/**
 * Reads downtime entries from XLSX sheet from row in startIndexArray, cleans strings and numbers using {@link getCleanedRow} and {@link cleanNumber}.
 * Returns sheet with downtime entries.
 * @param {Array<Array<String>>} formattedData - cells in one worksheet
 * @param {Array<Array<[Number, String]>>} startIndexArray 
 * @returns {Array<Array<String>>}
 */
async function readDowntimeEntries(formattedData, startIndexArray) {
    let sheetRows = []

    console.log(startIndexArray)

    try {

        let startIdx = startIndexArray[1][0] + 2

        console.log(`start index is ${startIdx}`);

        for (let i = startIdx; i < formattedData.length; i++) {
            if (formattedData[i][2] === null || formattedData[i][2] === undefined) {
                break;
            };

            let raw_row = formattedData[i]

            console.log(formattedData[i][2])

            let row = [raw_row[0], raw_row[2], raw_row[3], cleanNumber(raw_row[5]), raw_row[6], raw_row[9], raw_row[12], raw_row[14], cleanNumber(raw_row[16]), raw_row[17], cleanNumber(raw_row[18])]
            let cleaned_row = getCleanedRow(row)

            if (cleaned_row.length !== 0) {
                sheetRows.push(cleaned_row)
            };

        };
    } catch (err) {
        console.error(err.message)
        sheetRows = []
    }


    return sheetRows
}

/**
 * Reads downtime entries from XLSX sheet from row in startIndexArray, cleans strings and numbers using {@link getCleanedRow} and {@link cleanNumber}.
 * Returns sheet with downtime entries.
 * @param {Array<Array<String>>} formattedData - cells in one worksheet
 * @param {Array<Array<[Number, String]>>} startIndexArray 
 * @returns {Array<Array<String>>}
 */
async function readCartIssues(formattedData, startIndexArray) {
    const sheet_rows = []

    let start_idx = startIndexArray[0][0] + 2

    console.log(`start index is ${start_idx}`);

    for (let i = start_idx; i < formattedData.length; i++) {
        const firstCell = formattedData[i][0];

        if (firstCell === null || firstCell === undefined || firstCell.toUpperCase() === "DRIVE WHEEL SLIPS BY LEG") {
            break;
        };

        let raw_row = formattedData[i]

        let row = [raw_row[0], raw_row[3], raw_row[17]]
        let cleaned_row = getCleanedRow(row)

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

            let formattedData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            const startIndexArray = findStarterElements(formattedData, "AGC ISSUES")

            if (startIndexArray.length === 0) {
                reject([])
            }

            console.log(`START INDEX ARRAY IS ${startIndexArray}`)

            sheet_rows.downtimeEntries = await readDowntimeEntries(formattedData, startIndexArray)
            sheet_rows.cartIssues = await readCartIssues(formattedData, startIndexArray)


            resolve(sheet_rows);
        } catch (error) {
            reject(error)
        }
    });
};

function saveFile(downtime, cart) {
    console.log(cart);
    const downtimeResults = [...downtime]
    const cartResults = [...cart]

    downtimeResults.unshift(["Cart", "Location", "Issue", "Downtime", "RootCause", "Reason", "CorrectiveAction", "Owner", "Shift", "Date", "Time"])
    cartResults.unshift(["Cart", "Description", "Work Order #"])

    const book = XLSX.utils.book_new();
    const downtimeSheet = XLSX.utils.aoa_to_sheet(downtimeResults);
    const cartSheet = XLSX.utils.aoa_to_sheet(cartResults)

    XLSX.utils.book_append_sheet(book, downtimeSheet, 'EOS_Compiled_Data');
    XLSX.utils.book_append_sheet(book, cartSheet, 'Cart_Issues_Compiled');

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
    document.getElementById("stationVsDowntimeSummary").remove();
}

function createChart(topEightStations, percentages) {
    let canvas = document.createElement("canvas");
    canvas.setAttribute("id", "stationVsDowntimeSummary");
    document.querySelector("#chart_holder").appendChild(canvas);

    let ctx = document.getElementById("stationVsDowntimeSummary");

    let chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topEightStations.map(loc => loc[0]),
            datasets: [{
                data: topEightStations.map(loc => loc[1]),
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

                    chartInstance.data.datasets.forEach(function (_, i) {
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