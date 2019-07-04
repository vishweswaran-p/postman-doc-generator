const fs = require('fs');
const Excel = require('exceljs');
const prompt = require('prompt');

const errors = {
    REQUIRED : new Error('JSON path and XLSX path are required.'),
    EMPTY: new Error('JSON path and XLSX path cannot be empty.'),
    INVALID_PATH: new Error('JSON path or XLSX path is invalid.'),
    INVALID_JSON_FILE: new Error('JSON file is not valid.'),
    UNKNOWN: (err) => {
        return new Error(err);
    }
};

/**
 * @method buildExcelFile
 * @param inputs
 * @return {Promise<any>}
 */
function buildExcelFile(inputs) {
    return new Promise((resolve, reject) => {
        try {
            if(inputs.hasOwnProperty('json_path') && inputs.hasOwnProperty('xlsx_path')) {
                if(inputs.json_path.length !== 0 && inputs.xlsx_path.length !== 0) {
                    if(inputs.json_path.includes('.json') && inputs.xlsx_path.includes('.xlsx')) {

                        // Reading the json data from the postman collection file
                        let postman = JSON.parse(fs.readFileSync(inputs.json_path));

                        if(postman.hasOwnProperty('item')) {
                            let workbook = new Excel.Workbook();
                            workbook.creator = 'Vishnu';
                            workbook.lastModifiedBy = 'Vishnu';
                            workbook.created = new Date();
                            workbook.modified = new Date();
                            workbook.properties.date1904 = true;
                            let worksheet = workbook.addWorksheet('API Document', {
                                views: [{
                                    xSplit: 1,
                                    ySplit: 1
                                }]
                            }, {properties: {tabColor: {argb: 'ff5050'}}});

                            worksheet.state = 'visible';

                            worksheet.columns =  [
                                {header: 'S.No', key: 'S.No', width: 6},
                                {header: 'Category', key: 'Category', width: 20},
                                {header: 'API Name', key: 'API Name', width:30 },
                                {header: 'API Headers', key: 'API Headers', width:40 },
                                {header: 'API Method', key: 'API Method', width:20 },
                                {header: 'API Handler', key: 'API Handler', width:40 },
                            ];

                            let rows = [];

                            let k = 0;

                            for(let i=0;i<postman.item.length;i++) {
                                for(let j=0;j<postman.item[i].item.length;j++) {
                                    k++;
                                    rows.push({
                                        'S.No':k,
                                        'Category':postman.item[i].name,
                                        'API Name':postman.item[i].item[j].name,
                                        'API Headers':(postman.item[i].item[j].request.header.map(e => {return e.key}).join(',')).split(',').sort().join(', '),
                                        'API Method':postman.item[i].item[j].request.method,
                                        'API Handler':postman.item[i].item[j].request.url.raw.slice(postman.item[i].item[j].request.url.raw.indexOf('/'), postman.item[i].item[j].request.url.raw.lastIndexOf('/'))
                                    });
                                }
                                rows.push({});
                            }

                            worksheet.views = [
                                {state: 'frozen', xSplit: 0, ySplit: 1 }
                            ];

                            worksheet.addRows(rows);

                            worksheet.eachRow((row, rowNumber) => {
                                row.eachCell((cell, colNumber) => {
                                    row.getCell(colNumber).alignment = { wrapText: true };
                                });
                                if(rowNumber == 1) {
                                    row.eachCell((cell, colNumber) => {
                                        row.getCell(colNumber).style = { 'font': {'bold': true,'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'}};
                                        row.getCell(colNumber).alignment = { wrapText: true };
                                        row.getCell(colNumber).fill = {
                                            type: 'pattern',
                                            pattern:'solid',
                                            fgColor:{argb:'b3b3b3'},
                                            bgColor:{argb:'000000'}
                                        };
                                    });
                                } else {
                                    row.getCell(5).alignment = { horizontal: 'right'};
                                    row.getCell(6).alignment = { horizontal: 'right'};
                                }
                            });

                            workbook.xlsx.writeFile(inputs.xlsx_path).then(result => {
                                resolve({status:'success',xlsx_path:inputs.xlsx_path});
                            })
                            .catch(err => {
                                throw err;
                            });
                        } else {
                            reject(errors.INVALID_JSON_FILE);
                        }
                    } else {
                        reject(errors.INVALID_PATH);
                    }
                } else {
                    reject(errors.EMPTY);
                }
            } else {
                reject(errors.REQUIRED);
            }
        } catch(err) {
            reject(errors.UNKNOWN(err));
        }
    })
}

module.exports = {
    buildExcelFile: buildExcelFile
};