var fs = require('fs');
var XLSX = require('xlsx');
var fields = ['INDV_CASE_ID', 'AHBX_CASE_ID', 'SPECL_ENRL_EVENT_ID'];

// load the excel file
var workbook = XLSX.readFile('sample_pcr_file.xlsx');
//console.log(workbook);

// Get the sheet "tracker"
var sheet_name = workbook.SheetNames[1];
var worksheet = workbook.Sheets[sheet_name];

// sheet -> json -> sheet -> csv

var desired_value = XLSX.utils.sheet_to_json(worksheet, { header: ["date_added", "srq", "ind_case_id", "calheers_case_id", "enrol_evt_id", "added_by", "comments", "pcr_number", "pcr_status", "pcr_rundate"] });

var output = desired_value;
var filter_output = []; // array of objects

for (let i = 0; i < output.length; i++) {
    if (output[i]['pcr_status'] != 'Applied') {
        let temp = output[i];
        delete temp['srq']
        delete temp['date_added']
        delete temp['added_by']
        "comments", "pcr_number", "pcr_status", "pcr_rundate"
        delete temp['comments']
        delete temp['pcr_number']
        delete temp['pcr_status']
        delete temp['pcr_rundate']

        filter_output.push(temp);
    }
}

filter_output.splice(0, 1);

console.log(filter_output.length);
var iterCount = Math.ceil(filter_output.length / 100);
console.log(iterCount);
var iterArray = [];
for (var i = 0; i < iterCount; i++) {

    iterArray.push(filter_output.slice(i * 100, filter_output.length < (i + 1) * 100 ? filter_output.length : (i + 1) * 100))

    console.log(filter_output.slice(i * 100, (i + 1) * 100))

}

console.log(iterArray)
var d = new Date();
var filename = d.getDate().toString() +
    d.getMonth().toString() +
    d.getFullYear().toString();

for (var i = 0; i < iterArray.length; i++) {
    var temp = XLSX.utils.json_to_sheet(iterArray[i]);
    console.log(temp)
    temp = XLSX.utils.sheet_to_csv(temp);
    console.log(temp);

    // write stream and export __filename__.csv
    fs.writeFile(filename + '-' + (i + 1) + '.csv', temp, function(err) {
        if (err) throw err;
        console.log('file saved');
    });
}