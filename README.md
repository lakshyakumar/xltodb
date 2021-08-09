# XlToDb

XlToDb is a node package that collects data from an excel file and adds the collected rows to the MongoDB database. The package uses the column header or the 1st row as the key elements and forms an object with each row. Quick and easy to implement.

## Installation

Use the node npm to install XlToDb.

```bash
npm install xltodb
```

## Example Usage

```javascript
const xltodb = require('xltodb')

// initiate mongo client
const { MongoClient } = require("mongodb");
const client = new MongoClient('mongodb_connection_string')

// Instantiate Transify
let XlToDb = new xltodb('./test_sheet.xlsx')

// Convert the excel data to json object
let dataInJson = XlToDb.xlToJson()

// Save the data to MongoDB
XlToDb.saveToDb(client, "DataBase_Name", "collection_name").then((result)=>{
    console.log(result)
})
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.
