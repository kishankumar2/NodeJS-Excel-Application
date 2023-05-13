const express = require("express")
const app = express()
var nforce = require('nforce');
const Excel = require('exceljs');
const fs = require('fs');
const { log } = require("console");

app.set('port', process.env.PORT || 5000);
//app.use(express.static("public"))
app.use(express.urlencoded({ extended: true }))
app.use(express.json())


app.set("view engine", "ejs")
app.enable('trust proxy');

function isSetup() {
    return (process.env.CONSUMER_KEY != null) && (process.env.CONSUMER_SECRET != null);
  }

app.get('/', function(req, res) {
    console.log('Inside the get method');
    
    if (isSetup()) {
        var org = nforce.createConnection({
            clientId: process.env.CONSUMER_KEY,
            clientSecret: process.env.CONSUMER_SECRET,
            mode: 'single' // optional, 'single' or 'multi' user mode, multi default
          });
       
          // authenticated
          org.authenticate({ username: process.env.User_Name, password: process.env.Password}, function(err,resp) {
            if (!err) {
              console.log('Access Token: ' + resp.access_token);
              
              org.query({ query: 'SELECT id, name, industry FROM Account' }, function(err, results) {
                if (!err) {
                    console.log('results.records: ' + JSON.stringify(results.records));
                    const wb = new Excel.Workbook();
                    const AccountSheet = wb.addWorksheet('Accounts', {properties:{tabColor:{argb:'264653'}}});
                    AccountSheet.addRows([['ID','Account Name','Industry']]);
                    console.log('Parsing Starts');
                    console.log(results.records[0]);
                    results.records.forEach(function(record){
                        AccountSheet.addRow([record._fields.id,record._fields.name,record._fields.industry]);
                     });

                    
                    console.log('Parsing Ends');
                    AccountSheet.columns.forEach(column => {
                        column.border = {
                          top: { style: "thick" },
                          left: { style: "thick" },
                          bottom: { style: "thick" },
                          right: { style: "thick" }
                        };
                      });
                      const fileName = './assets/simple.xlsx';
                      wb.xlsx.writeFile(fileName)
                        .then(() => {
                               console.log('file created');
                               var dataa =fs.readFileSync(fileName,'base64');
                    console.log('dataa'+dataa);  
                    res.send(dataa);
                              })
                        .catch(err => {
                            console.log(err.message);
                        });


                    
                 //res.render('index1');
                }
                else {
                    console.log('Error: ' + err.message);
                  res.send(err.message);
                }
              });
            }
            else {
              if (err.message.indexOf('invalid_grant') >= 0) {
                res.send(err.message);
                
              }
              else {
                res.send(err.message);
              }
            }
          });
       
        }
    

       
    

});


app.listen(app.get('port'), function () {
    console.log('Express server listening on port ' + app.get('port'));
});
