const csv2sql = require('csv2sql-stream');
const XLSX = require('xlsx');
const dotenv = require('dotenv');
const massive = require('massive');
const workbook = XLSX.readFile('schema.ods', { encoding: 'UTF-8'});

async function init(){

    try{
        dotenv.config();
    } catch(error) {
        throw error;
    }
    
    try{
        var db = await massive({ host: process.env.DB_HOST, port: process.env.DB_PORT, database: process.env.DB_NAME, user: process.env.DB_USER, password: process.env.DB_PASSWORD });
    } catch (err){
        console.log(`Could not establish connection to database:\n${err}`);
        throw err;
    }


    try{
        await db.createTables();
    } catch (err){
        console.log(`Error while attempting to create tables:\n${err}`);
        throw err;
    }

    console.log('Created tables ...');

    const promises = []

    for(const worksheetName in workbook.Sheets){
        promises.push(
            new Promise((resolve, reject) => {

                let out = '';

                csv2sql.transform(worksheetName, XLSX.stream.to_csv(workbook.Sheets[worksheetName]))
                .on('data', function(sql){
                    out+=sql;
                })
                .on('end', function(rows){
                    console.log(`${rows} rows inserted into table ${worksheetName}`);
                    resolve(out);
                })
                .on('error', function(error){
                    reject(error);
                });
            })
        );
    }
    
    try{
        var query = (await Promise.all(promises)).join('');
    } catch (err){
        throw err;
    }

    console.log('Converted spreadsheet into SQL inserts ...');

    try{
        var result = await db.query(query);
    } catch {
        throw err;
    }

    console.log('Ran those inserts against the database ...');
}


init()
.then(() => process.exit(0))
.catch((err) => process.exit(1));