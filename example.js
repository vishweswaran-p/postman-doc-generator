const generator = require('postman-doc-generator');

let inputs = {
    json_path: 'JSON_PATH_HERE',
    xlsx_path: 'XLSX_PATH_HER'
};

generator.buildExcelFile(inputs).then(result => {
    console.log(result);
})
.catch(err => {
    console.log(err);
});