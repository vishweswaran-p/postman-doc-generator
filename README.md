# postman-doc-generator

This module is used to create a xlsx file about the list of API's from the postman collection JSON file.

## Getting Started

Install the module using the following command.

```
npm i postman-doc-generator
```

## Example

Below given is the example to use this package.

```
const docGenerator = require('postman-doc-generator');

let inputs = {
    json_path: '/home/user/postman/postman_collection.json',
    xlsx_path: '/home/user/Desktop/doc.xlsx'
};

docGenerator.buildExcelFile(inputs).then(result => {
    console.log(result); 
    // Returns this result 
    // { 
    //   status: 'success',
    //   xlsx_path: '/home/user/Desktop/doc.xlsx' 
    // }
})
.catch(err => {
    console.log(err);
});
```


## Authors

* **Vishweswaran P** - *Initial work* - [Github](https://github.com/vishweswaran-p)

## License

This project is licensed under the ISC License.
