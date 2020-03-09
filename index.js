const fs = require('fs');
const Excel = require('exceljs');
const config = require('./config/config.js');
const destpath = '../../../OneDrive - Ximply/';
const format = '.xlsx';
const filepath = '../../../OneDrive - Ximply/Template'+format;

function getMondayAndFriday(date){
    
    let days = {}
    let monday = date.getDate() - date.getDay() + (date.getDay() === 0 ? -6 : 1);
    days.monday = new Date(date.setDate(monday)).toISOString().slice(0,10);
    days.friday = new Date(date.setDate(monday+4)).toISOString().slice(0,10);
    return days;

}

function formatDate(s){

    let arr = s.split('-');
    return arr[2]+'.'+arr[1]+'.'+arr[0];
    
}

function execute(name) {
    
    let days = getMondayAndFriday(new Date());
    let filename = days.monday + ' to ' + days.friday + format;
    var workbook = new Excel.Workbook();

    workbook.xlsx
    .readFile(filepath)
    .then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);

        row.getCell(2).value = name;

        row = worksheet.getRow(4);
        row.getCell(1).value = formatDate(days.monday);
        row.commit();

        if (!fs.existsSync(destpath+name)){
            fs.mkdirSync(destpath+name);
        }
        return workbook.xlsx.writeFile(destpath+name+'/'+filename);

    });
}

config.employees.forEach(
    (employee) => { execute(employee.name); }
);