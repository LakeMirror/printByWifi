var DocumentUtil = function() {

    function addNewText(text) {
        var body = document.getElementById('body');
        body.append(text);
    }

    function addTable(id, excelData) {
        var body = document.getElementById('body');
        body.append(`<div id = ${id}></div>`);


    }

}