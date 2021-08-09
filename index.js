// Requiring the module
const reader = require('xlsx')

// Creating class Transify
const XlToDb = class {
    //defining a constructor for the class
    constructor(filePath, sheetName = '*') {
        // reading the .xlsx file
        this.file = reader.readFile(filePath)
        //sheetnames if provided
        this.sheetName = sheetName
        //json data from the sheet
        this.data = []
    }

    // Function to create json objects from .xlsx sheet rows
    xlToJson = () => {
        // try to extract data from the .xlxs file
        try {
            const sheets = this.sheetName === '*' ? this.file.SheetNames : [...this.sheetName]
            for (let i = 0; i < sheets.length; i++) {
                // Read data sheet-wise
                const temp = reader.utils.sheet_to_json(
                    this.file.Sheets[sheets[i]])
                temp.forEach((res) => {
                    //add the sheet data to the data array
                    this.data.push(res)
                })
            }
            // return data
            return this.data
        } catch (e) {
            return `cannot xlToJson() your excel, some error occured ${e}`
        }
    }

    // Function to add the data to the mongoDB database
    saveToDb = (client, dataBase_name, collectionName) => {
        // Returning a promise as DB interaction is async in nature
        return new Promise(async (resolve, reject) => {
            // Checking if the data object is not empty
            if (this.data.length) {
                try {
                    // connecting DB client
                    await client.connect();
                    // setting the database
                    const database = client.db(dataBase_name);
                    // setting the collection
                    const collection = database.collection(collectionName);
                    // this option prevents additional documents from being inserted if one fails
                    const options = { ordered: true };
                    // saving the data to the MongoDB
                    const result = await collection.insertMany(this.data, options);
                    resolve(result)

                } catch (e) {
                    // Reject the Promise in case of an issue
                    reject(e)
                } finally {
                    // finally closing the database connection
                    await client.close();
                }
            }
            // reject the promise in case of empty sheet
            reject('Empty data object')
        })


    }

}
// Exporting the Transify class
module.exports = Transify



