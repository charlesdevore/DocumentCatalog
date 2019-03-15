var execBtn = document.getElementById("execute");
var outputElm = document.getElementById('output');
var errorElm = document.getElementById('error');
var commandsElm = document.getElementById('commands');
var dbFileElm = document.getElementById('dbfile');

// Connect to the HTML element we 'print' to
function print(text) {
    outputElm.innerHTML = text.replace(/\n/g, '<br>');
}
function error(e) {
    console.log(e);
    errorElm.style.height = '2em';
    errorElm.textContent = e.message;
}

function noerror() {
    errorElm.style.height = '0';
}


// Run a command in the database
function execute(commands) {
    results = db.exec(commands)
    for (var ii=0; ii<results.length; ii++) {
        outputElm.appendChild(tableCreate(results[ii].columns, results[ii].values));
    }
}

// Create an HTML table
var tableCreate = function () {
    function valconcat(vals, tagName) {
        if (vals.length === 0) return '';
        var open = '<'+tagName+'>', close='</'+tagName+'>';
        return open + vals.join(close + open) + close;
    }
    return function (columns, values){
        var tbl  = document.createElement('table');
        tbl.setAttribute('class', 'fixed_header');
        var html = '<thead>' + valconcat(columns, 'th') + '</thead>';
        var rows = values.map(function(v){ return valconcat(v, 'td'); });
        html += '<tbody>' + valconcat(rows, 'tr') + '</tbody>';
	tbl.innerHTML = html;
        return tbl;
    }
}();

// Execute the commands when the button is clicked
function execEditorContents () {
    noerror()
    execute (editor.getValue() + ';');
}
execBtn.addEventListener("click", execEditorContents, true);

// Performance measurement functions
var tictime;
if (!window.performance || !performance.now) {window.performance = {now:Date.now}}
function tic () {tictime = performance.now()}
function toc(msg) {
    var dt = performance.now()-tictime;
    console.log((msg||'toc') + ": " + dt + "ms");
}

// Add syntax highlihjting to the textarea
var editor = CodeMirror.fromTextArea(commandsElm, {
    mode: 'text/sql',
    viewportMargin: Infinity,
    indentWithTabs: true,
    smartIndent: true,
    lineNumbers: true,
    matchBrackets : true,
    autofocus: true,
    extraKeys: {
	"Ctrl-Enter": execEditorContents,
    }
});

// Load a db from a file
dbFileElm.onchange = function() {
    var f = dbFileElm.files[0];
    var r = new FileReader();
    r.onload = function() {
	var Uints = new Uint8Array(r.result);
        db = new SQL.Database(Uints);
    }
    r.readAsArrayBuffer(f);
}

