let inputElement  = document.getElementById("commands");
let executeButton = document.getElementById("execute");
let outputElement = document.getElementById("output");
let errorElement  = document.getElementById("error");
let dbFileElement = document.getElementById("dbupload");
let isError = false;

function error(e) {
    console.log(e);
    errorElement.style.height = '2em';
    errorElement.hidden = false;
    errorElement.innerText = e.message;
}

function noerror() {
    errorElement.hidden = true;
    isError = false;
}

// Run a command in the database
function execute(commands) {
    results = db.exec(commands);
    if (results.length == 0) {
        error(Error('Command returned no results.'));
        isError = true;
        return;
    };
    outputElement.innerHTML = (tableCreate(results[0].columns, results[0].values)).outerHTML;
    if (results.length != 1) {
        alert("Please only run one SQL command at a time.");
    };
    // Resize the div containers
    let innerDiv = document.getElementById('innerContainerDiv')
    innerDiv.style.height = outputElement.style.height;
}

// Create an HTML table
let tableCreate = function () {
    function valconcat(vals, tagName) {
        if (vals.length === 0) return '';
        var open = '<'+tagName+'>', close='</'+tagName+'>';
        return open + vals.join(close + open) + close;
    }
    function buildDataCell(val) {
        let cell = document.createElement('td');
        cell.innerHTML = val;
        return cell.outerHTML;
    }
    
    function buildDataRow(vals) {
        let row = document.createElement('tr');
        let children = []
        let anchor = document.createElement('a');

        // The URI must be located in the last column and the text of
        // the link is the first column
        anchor.href = vals.slice(-1)[0]
        anchor.innerText = vals[0]
        anchor.target = "_blank"

        row.innerHTML = buildDataCell(anchor.outerHTML) + vals.slice(1).map(buildDataCell).join('')
        
        return row.outerHTML
    }
    return function (columns, values){
        let innerContainer = document.createElement('div');
        innerContainer.setAttribute('id', 'innerContainerDiv');
        innerContainer.setAttribute('class', 'inner-container');
        let headerDiv = document.createElement('div');
        headerDiv.setAttribute('id', 'headerdiv');
        headerDiv.setAttribute('class', 'table-header');
        let headerTable = document.createElement('table');
        headerTable.setAttribute('id', 'headertable');

        headerTable.innerHTML = '<thead>' + valconcat(columns, 'th') + '</thead>';
        headerDiv.innerHTML = headerTable.outerHTML;

        let bodyDiv = document.createElement('div');
        bodyDiv.setAttribute('id', 'bodydiv');
        bodyDiv.setAttribute('class', 'table-body');
        bodyDiv.setAttribute('onscroll',
                              "document.getElementById('headerdiv').scrollLeft = this.scrollLeft;");
        let bodyTable = document.createElement('table');
        bodyTable.setAttribute('id', 'bodytable');

        // let rows = values.map(function(v){ return valconcat(v, 'td'); });
        // bodyTable.innerHTML = '<tbody>' + valconcat(rows, 'tr') + '</tbody>';

        bodyTable.innerHTML = values.map(buildDataRow).join('')
        bodyDiv.innerHTML = bodyTable.outerHTML;

        innerContainer.innerHTML = headerDiv.outerHTML + bodyDiv.outerHTML;

        return innerContainer;
    }
}();

function getCommands () {
    let preamble = document.getElementById('preamble').innerText
    
    return preamble + inputElement.value + ';'
}

// Execute the commands when the button is clicked
function executeEditorContents () {
    noerror()
    execute (getCommands());
    if (isError){
        return;
    }
    // Resize the div container
    let innerDiv = document.getElementById('innerContainerDiv')
    let headerDiv = document.getElementById('headerdiv');
    let bodyDiv = document.getElementById('bodydiv');
    let height = headerDiv.clientHeight + bodyDiv.clientHeight;
    innerDiv.setAttribute('style', `height:${height}px;`);

    // Update header row cell width
    let headerTable = document.getElementById('headertable');
    let headerRow = headerTable.children[0].children[0];
    let bodyTable = document.getElementById('bodytable');
    let bodyRow = bodyTable.children[0].children[0];

    // Update headerTable width based on bodyTable with an additional 16px for the vertical
    // scrollbar
    headerTable.width = bodyTable.offsetWidth + 16;

    // Update each cell's width and subtract 4px for padding on the left and right
    for (let ii = 0; ii < headerRow.children.length; ii++) {
        headerRow.children[ii].width = bodyRow.children[ii].offsetWidth - 4;
    };
    headerRow.children[0].width -= 4;
}
executeButton.addEventListener("click", executeEditorContents, true);



// Load a db from a file
dbFileElement.onchange = function() {
    let f = dbFileElement.files[0];
    let r = new FileReader();
    r.onload = function() {
	let Uints = new Uint8Array(r.result);
        db = new SQL.Database(Uints);
    }
    r.readAsArrayBuffer(f);
}


// Export results table to CSV
function exportTableToCSV(filename) {
    let csv = [];

    // Extract header columns and store in array
    extractHeader(csv);
    
    // Extract body table
    extractBody(csv);
    
    // Download CSV file
    downloadCSV(csv.join("\n"), filename);
}

function extractBody(csv) {
    // Extract data rows from body table and store in csv array
    let rows = document.getElementById('bodytable').querySelectorAll('tr');

    for (let i = 0; i < rows.length; i++) {
        let row = [];
        let cells = rows[i].querySelectorAll('td');

        for (let j = 0; j < cells.length; j++) {
            row.push(cells[j].innerText);
        }
        csv.push('"'+row.join('","')+'"');
    }
}

function extractHeader(csv) {
    // Extract header columns from result table
    let headerCols = document.getElementById('headertable').querySelectorAll('th');
    let header = [];
    for (let j = 0; j < headerCols.length; j++) {
        header.push(headerCols[j].innerText);        
    }
    csv.push('"'+header.join('","')+'"');
}

function downloadCSV(csv, filename) {
    let csvFile;
    let downloadLink;

    // CSV file
    csvFile = new Blob([csv], {type: "text/csv"});

    // Download link
    downloadLink = document.createElement("a");

    // File name
    downloadLink.download = filename;

    // Create a link to the file
    downloadLink.href = window.URL.createObjectURL(csvFile);

    // Hide download link
    downloadLink.style.display = "none";

    // Add the link to DOM
    document.body.appendChild(downloadLink);

    // Click download link
    downloadLink.click();
}
