let inputElement  = document.getElementById("commands");
let executeButton = document.getElementById("execute");
let outputElement = document.getElementById("tabulator-table");
let errorElement  = document.getElementById("error");
let dbFileElement = document.getElementById("dbupload");
let isError = false;

function error(e) {
    console.log(e);
    errorElement.style.height = '2em';
    errorElement.hidden = false;
    errorElement.innerText = e.message;
    outputElement.hidden = true;
    isError = true;
}

function noerror() {
    errorElement.hidden = true;
    isError = false;
    outputElement.hidden = false;
}

// Run a command in the database
function execute(commands) {

    if (typeof(db) == "undefined") {
        error(Error('No database file defined. Please load database using link above.'));
        return;
    };
    
    results = db.exec(commands);
    if (results.length == 0) {
        error(Error('Command returned no results.'));
        return;
    };

    if (results.length != 1) {
        alert("Please only run one SQL command at a time.");
    };

    table = buildTabulatorTable(results[0].columns, results[0].values);
}


function getCommands () {
    let basicDiv = document.getElementById('basic');
    let advancedDiv = document.getElementById('advanced');
    if (basicDiv.hidden & !advancedDiv.hidden) {
        return document.getElementById('preamble').innerText + buildAdvancedCommand() + ';';
    } else if (!basicDiv.hidden & advancedDiv.hidden) {
        return document.getElementById('preamble').innerText + buildBasicCommand() + ';';
    } else {
        error(Error('Conflicting state in Basic/Advanced views'));
    }
}

// Execute the commands when the button is clicked
function executeEditorContents () {
    noerror()

    // clear results table
    outputElement.innerHTML = '';
    // outputElement.style.height = 0;
    
    execute (getCommands());
    if (isError){
        return;
    }
}
executeButton.addEventListener("click", executeEditorContents, true);

function handleSearchKeyPress(e) {
    if (e.keyCode == 13) {
        e.preventDefault();
        executeEditorContents();
    }
}

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

// Export query to text file
function exportQuery(filename) {
    downloadFile(buildBasicCommand(), filename, 'text/plain');
}


// Export results table to CSV
function exportTableToCSV(filename) {
    table.download('csv', filename);
}

function downloadFile(csv, filename, fileType) {
    let outFile;
    let downloadLink;

    // CSV file
    outFile = new Blob([csv], {type: "`${fileType}`"});

    // Download link
    downloadLink = document.createElement("a");

    // File name
    downloadLink.download = filename;

    // Create a link to the file
    downloadLink.href = window.URL.createObjectURL(outFile);

    // Hide download link
    downloadLink.style.display = "none";

    // Add the link to DOM
    document.body.appendChild(downloadLink);

    // Click download link
    downloadLink.click();
}


// Switch to Advanced view
function switchAdvancedView() {
    document.getElementById('basic').hidden=true;
    document.getElementById('commands').innerHTML = buildBasicCommand()
    document.getElementById('advanced').hidden=false;
}

// Switch to Basic view
function switchBasicView() {
    document.getElementById('advanced').hidden=true;
    document.getElementById('basic').hidden=false;
}


// Function to control the state of the radio checkboxes
function clickedRadioAll() {
    let allRadio = document.getElementById('all-radio');
    if (allRadio.checked) {
        setRadioFiletypes(true);
    } else {
        setRadioFiletypes(false);
    }
}


// Function set the disabled parameter for all the radio filetype options except all files.
function setRadioFiletypes(state) {
    document.getElementById('pdf-docs-radio').disabled = state;
    document.getElementById('word-docs-radio').disabled = state;
    document.getElementById('xls-docs-radio').disabled = state;
}


function buildBasicCommand() {
    let sqlCommand = '';
    let sortCommand = '';

    let searchString = document.getElementById('searchbar').value;
    // Pre- and post-pend the search string with SQL wildcard
    searchString = '"%' + searchString + '%"';

    // Replace * and whitespace with SQL wildcard
    searchString = searchString.replace(/[\*|\s]/g, '%');

    let searchField = document.getElementById('search-field');
    sqlCommand += 'WHERE';
    if (searchField.value == 'select-filename') {
        sqlCommand += ' "File Name" ';
        sortCommand = '\nORDER BY "File Name"';
    } else if (searchField.value == 'select-relpath') {
        sqlCommand += ' "Relative Path" ';
        sortCommand = '\nORDER BY "Relative Path"';
    } else if (searchField.value == 'select-filekey') {
        sqlCommand += ' "Unique Id" ';
        sortCommand = '\nORDER BY "File Name"';
    } else {
        error(Error('Undefined search field'));
    }
    
    sqlCommand += 'LIKE ';
    sqlCommand += searchString;

    // Set the extension clauses
    if (!document.getElementById('all-radio').checked) {
        let extensions = [];
        if (document.getElementById('pdf-docs-radio').checked) {
            extensions.push(' "File Type" = ".pdf" ');
        }
        if (document.getElementById('word-docs-radio').checked) {
            extensions.push(' "File Type" LIKE ".doc%" ');
        }
        if (document.getElementById('xls-docs-radio').checked) {
            extensions.push(' "File Type" LIKE ".xl%" ');
        }
        if (extensions.length == 1) {
            sqlCommand += ' \n AND ' + extensions[0];
        } else if (extensions.length > 1) {
            sqlCommand += ' \n AND ( ' + extensions[0];
            for (let i = 1; i < extensions.length; i++) {
                sqlCommand += ' \n OR ' + extensions[i];
            }
            sqlCommand += ' ) ';
        }
    }

    // Determine if duplicates should be included or not
    if (!document.getElementById('duplicates').checked) {
        sqlCommand += '\nGROUP BY "Checksum ID"';
    };
    
    return sqlCommand + sortCommand;
}

function buildAdvancedCommand() {
    return inputElement.value;
}


// Tabulator Commands
function buildExampleTabulatorTable() {

    let tabledata = [
        {id:1, name:"Oli Bob", age:"12", col:"red", dob:""},
        {id:2, name:"Mary May", age:"1", col:"blue", dob:"14/05/1982"},
        {id:3, name:"Christine Lobowski", age:"42", col:"green", dob:"22/05/1982"},
        {id:4, name:"Brendon Philips", age:"125", col:"orange", dob:"01/08/1980"},
        {id:5, name:"Margret Marmajuke", age:"16", col:"yellow", dob:"31/01/1999"},
    ];

    let table = new Tabulator("#tabulator-table", {
        height:100,
        data:tabledata,
        layout:"fitColumns",
        columns:[
	    {title:"Name", field:"name", width:150},
	    {title:"Age", field:"age", align:"left", formatter:"progress"},
	    {title:"Favourite Color", field:"col"},
	    {title:"Date Of Birth", field:"dob", sorter:"date", align:"center"},
        ],
        rowClick:function(e, row){
            alert("Row " + row.getData().id + " Clicked!");
        },
    });

}
        

function buildTabulatorTable(columns, values) {
    function processInputColumns(columns) {
        let tableColumns = []
        for (let i=0; i < columns.length; i++) {
            if (columns[i] == "File Name") {
                tableColumns.push({id:i, title:"File Name", field:"fname", minWidth:150});
            } else if (columns[i] == "File Size") {
                tableColumns.push({id:i, title:"File Size", field:"hread", minWidth:50});
            } else if (columns[i] == "File Type") {
                tableColumns.push({id:i, title:"File Type", field:"ext", minWidth:50});
            } else if (columns[i] == "Relative Path") {
                tableColumns.push({id:i, title:"Relative Path", field:"rpath", minWidth:150});
            } else if (columns[i] == "Unique ID") {
                tableColumns.push({id:i, title:"Unique ID", field:"filekey"});
            } else {
                tableColumns.push({id:i, title:columns[i], field:columns[i].replace(/\s/, '')});
            }
        }

        return tableColumns;
    }

    function processInputValues(values, tableColumns) {
        let tableData = []

        for (let i=0; i < values.length; i++) {
            let row = {}
            // let anchor = document.createElement('a');

            // anchor.href = values[i].slice(-1)[0];
            // anchor.innerText = values[i][0];
            // anchor.target = '_blank';
           
            // row['id'] = i;
            // row['fname'] = anchor;

            // for (let j=1; j < values[i].length; j++) {
            //     row[tableColumns[j]['field']] = values[i][j];
            // }

            for (let j=0; j < values[i].length; j++) {
                row[tableColumns[j]['field']] = values[i][j];
            }
            
            tableData.push(row);
        }

        return tableData;
    }

    function openFile(rowData) {
        if (rowData.hasOwnProperty('URI')) {
            let fileLink = document.createElement('a');
            
            fileLink.href = rowData['URI'].replace(/#/,'%23');
            fileLink.target = '_blank';
            fileLink.style.display = 'none';
            
            document.body.appendChild(fileLink);
            fileLink.click();
            document.body.removeChild(fileLink);
            
        } else {
            alert('URI not defined for this row');
        }
    }

    let tableColumns = processInputColumns(columns);
    let tableData = processInputValues(values, tableColumns);
    
    let table = new Tabulator("#tabulator-table", {
        height:400,
        data:tableData,
        layout:"fitData",
        columns:tableColumns,
        rowClick:function(e, row){
            openFile(row.getData());
        },
    });

    return table;
    
}
