//npm module
const rfcClient = require('node-rfc').Client;
var mysql = require('mysql');
var fs = require('fs');
var ejs = require('ejs');
var _ = require('lodash');
var json2xls = require('json2xls');
var moment = require("moment");
var cron = require('node-cron');
const excel = require('node-excel-export');
var async = require('async');
const sgMail = require('@sendgrid/mail');
//sendgrid password
sgMail.setApiKey(process.env.SENDGRID_API_KEY);
console.log(moment("2019-06-20"));

function connectSAP(callback) {
    // ABAP system RFC connection parameters
    //dev env
    // const abapSystem = {
    //     user: 'dattaraj',
    //     passwd: 'best123',
    //     ashost: '10.91.40.20',
    //     sysnr: '02',
    //     client: '200',
    //     lang: 'EN',
    // };

    //prod env
    const abapSystem = {
        user: 'dattaraj',
        passwd: 'best1234',
        ashost: '10.91.40.21',
        sysnr: '00',
        client: '400',
        lang: 'EN',
    };

    // create new client
    const client = new rfcClient(abapSystem);

    // echo SAP NW RFC SDK and nodejs/RFC binding version
    console.log('Client version: ', client);

    // open connection
    client.connect(function (err) {
        if (err) {
            // check for login/connection errors
            return console.error('could not connect to server', err);
        } else {
            console.log("connected");
        }

        // invoke ABAP function module, passing structure and table parameters

        // // ABAP structure
        // const structure = {
        //   RFCINT4: 345,
        //   RFCFLOAT: 1.23456789,
        //   // or RFCFLOAT: require('decimal.js')('1.23456789'), // as Decimal object
        //   RFCCHAR4: 'ABCD',
        //   RFCDATE: '20180625', // in ABAP date format
        //   // or RFCDATE: new Date('2018-06-25'), // as JavaScript Date object
        // };
        // // ABAP table
        // let table = [structure];
        var functionName = 'ZHR_INVESTMENT_ERROR_UPLOAD';
        var reqData = {
            DATE_FROM: moment("2019-06-20").format('YYYYMMDD'),
            DATE_TO: moment("2019-06-20").format('YYYYMMDD'),
            // EMPLOYEE_CODE: 15105387
        }
        client.invoke(functionName, reqData, function (err, res) {
            // console.log(res);
            if (err) {
                callback(err, null);
                // return console.error('Error invoking ' + functionName + ':', err);
            }
            // console.log(functionName + ' call result:', res);
            var jsonfile = JSON.stringify(res);
            fs.writeFile('sap.json', jsonfile, 'utf8', function (err, data) {

            });
            callback(null, {
                res: res,
                excelData: res && res.ET_RESPONSE ? res.ET_RESPONSE : []
            });
        });
    });
}

function connectSQL(callback) {
    // var con = mysql.createConnection({
    //     host: 'test.kelltontech.net',
    //     port: '3306',
    //     user: 'bestsellers',
    //     password: 'bestsellers128',
    //     database: "bestsellers"

    // });
    var con = mysql.createConnection({
        host: '10.91.4.52',
        port: '3306',
        user: 'bestseller',
        password: 'bestLLerlssd',
        database: "bestsellers_test"
    });
    con.connect(function (err) {
        if (err) throw err;

        console.log("Connected! sql");
        con.query("select * from v_investment_details", function (err, result, fields) {
            if (err)
                callback(err, null);

            // console.log("data-------", result);
            // result = _.filter(result, function (data) {
            //     return moment(new Date(data.inv_date).setHours(0, 0, 0, 0)).isBetween(moment().subtract(8, 'days'), moment().subtract(8, 'days'));
            //     // return moment(new Date(data.inv_date).setHours(0, 0, 0, 0)).isBetween(moment().subtract(1, 'days'), moment());
            //     // return new Date(data.inv_date).setHours(0, 0, 0, 0) === new Date("2018/07/31").setHours(0, 0, 0, 0);
            // })

            var resultNew = [];
            if (result.length > 0) {
                _.forEach(result, function (data) {
                    var index = _.findIndex(resultNew, function (d2) {
                        return d2.empId === data.pernr;
                    })
                    if (index == -1) {
                        pushData = {
                            empId: data.pernr,
                        }
                        if (data.infty_code.startsWith("P")) {
                            pushData['IT' + data.infty_code.substring(2, 5)] = 'Success';
                        }
                        resultNew.push(pushData);
                    } else {
                        if (data.infty_code.startsWith("P")) {
                            resultNew[index]['IT' + data.infty_code.substring(2, 5)] = 'Success';
                        }
                        // if (_.isEmpty(resultNew[index][data.infty_code])) {
                        //     resultNew[index][data.infty_code] = [];
                        // } else {}
                        // resultNew[index][data.infty_code].push(data);
                    }
                });
            }
            var jsonfile = JSON.stringify(resultNew);
            fs.writeFile('sql.json', jsonfile, 'utf8', function (err, data) {

            });
            callback(null, {
                res: result,
                excelData: resultNew
            });
        });
    });

    con.on('error', function (err) {
        console.log('db error', err.code ? err.code : '');
        if (err.code === 'PROTOCOL_CONNECTION_LOST' || err.code === 'PROTOCOL_CONNECTION_LOST') { // Connection to the MySQL server is usually
            // connectSQL();                          // lost due to either server restart, or a
        } else { // connnection idle timeout (the wait_timeout
            throw err; // server variable configures this)
        }
    });


}

function excelGenerate(resData, sqlData) {
    //in attachment
    // You can define styles as json object
    const styles = {
        headerDark: {
            // fill: {
            //     fgColor: {
            //         rgb: 'FF000000'
            //     }
            // },
            font: {
                color: {
                    rgb: 'FF000000'
                },
                // sz: 14,
                bold: true,
                // underline: true
            }
        },
        cellPink: {
            fill: {
                fgColor: {
                    rgb: 'FFFFCCFF'
                }
            }
        },
        cellGreen: {
            fill: {
                fgColor: {
                    rgb: 'FF00FF00'
                }
            }
        }
    };

    const specification = {
        empId: {
            displayName: 'EMP. CODE',
            headerStyle: styles.headerDark,
            width: 80
        },
        empName: {
            displayName: 'EMP. NAME',
            headerStyle: styles.headerDark,
            width: 80
        },
        IT580: {
            displayName: 'IT580',
            headerStyle: styles.headerDark,
            width: 80
        },
        IT581: {
            displayName: 'IT581',
            headerStyle: styles.headerDark,
            width: 80
        },
        IT584: {
            displayName: 'IT584',
            headerStyle: styles.headerDark,
            width: 80
        },
        IT585: {
            displayName: 'IT585',
            headerStyle: styles.headerDark,
            width: 80
        },
        IT586: {
            displayName: 'IT586',
            headerStyle: styles.headerDark,
            width: 80
        }
    };

    const specs = {
        pernr: {
            displayName: 'EMP. CODE',
            headerStyle: styles.headerDark,
            width: 80
        },
        name: {
            displayName: 'NAME',
            headerStyle: styles.headerDark,
            width: 80
        },
        infty_code: {
            displayName: 'INV. CODE',
            headerStyle: styles.headerDark,
            width: 80
        },
        description: {
            displayName: 'DESCRIPTION',
            headerStyle: styles.headerDark,
            width: 80
        },
        inv_date: {
            displayName: 'INV. DATE',
            headerStyle: styles.headerDark,
            width: 80,
            cellFormat: function (value, row) {
                return moment(value).format("L");
            }
        },
        inv_amt: {
            displayName: 'INV. AMT',
            headerStyle: styles.headerDark,
            width: 80
        },
        is_provision: {
            displayName: 'PROVISION',
            headerStyle: styles.headerDark,
            width: 80
        }
    };

    var folder = "./";
    var path = "Reco_Report.xlsx";
    var finalPath = folder + path;
    var path1 = "investment_report.xlsx";
    var finalPath1 = folder + path1;
    async.parallel([
            function (callback) {
                const report = excel.buildExport(
                    [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
                        {
                            name: 'Reco_Report', // <- Specify sheet name (optional)
                            heading: [], // <- Raw heading array (optional)
                            merges: [], // <- Merge cell ranges
                            specification: specification, // <- Report specification
                            data: resData // <-- Report data
                        }
                    ]
                );

                fs.writeFile(finalPath, report, "binary", callback)
            },
            function (callback) {
                const report1 = excel.buildExport(
                    [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
                        {
                            name: 'Investment_Data', // <- Specify sheet name (optional)
                            heading: [], // <- Raw heading array (optional)
                            merges: [], // <- Merge cell ranges
                            specification: specs, // <- Report specification
                            data: sqlData // <-- Report data
                        }
                    ]
                );
                fs.writeFile(finalPath1, report1, "binary", callback)
            }
        ],
        function (err, results) {
            // fs.readFile(finalPath, function (err, excel) {
            //     if (err) {
            //         console.log(err);
            //     } else {}
            // });
            // fs.readFile(finalPath1, function (err, excel) {
            //     if (err) {
            //         console.log(err);
            //     } else {}
            // });
            if (!err) {
                sendMail(finalPath, finalPath1);
            }
        })


}

function sendMail(finalPath, finalPath1) {
    const msg = {
        from: 'bwsupport@bestseller.com',
        to: [
            // 'Ganesh.kothavale@bestseller.com'
            'abhishek.ghosh@bestseller.com',
            'ankit.shah@bestseller.com',
            'shreya.ambetkar@bestseller.com', 'ronak.pandya@bestseller.com'
        ],
        cc: 'Ganesh.kothavale@bestseller.com',
        subject: 'Investment Declaration Report:',
        text: 'Please find attachment for Investment Declaration Report'
    };
    if (_.isEmpty(finalPath)) {
        msg.text = 'No Investment Declaration Report Found';
    } else {
        msg.text = 'Please find attachment for Investment Declaration Report';
        var file = fs.readFileSync(finalPath);
        var file1 = fs.readFileSync(finalPath1);
        var base64File = new Buffer(file).toString("base64");
        var base64File1 = new Buffer(file1).toString("base64");
        msg.attachments = [{
                content: base64File,
                filename: 'reco_report_' + moment().subtract(1, 'days').format("ll") + '.xlsx',
                type: 'plain/text',
                disposition: 'attachment',
                contentId: 'mytext'
            },
            {
                content: base64File1,
                filename: 'investment_data_' + moment().subtract(1, 'days').format("ll") + '.xlsx',
                type: 'plain/text',
                disposition: 'attachment',
                contentId: 'mytext'
            }
        ];
    }

    sgMail.send(msg);
}

// cron.schedule('0 */12 * * *', () => {
// cron.schedule('0 9 * * *', () => {
async.parallel([
        function (callback) {
            connectSAP(callback);
        },
        function (callback) {
            connectSQL(callback);
        }
    ],
    function (err, results) {
        if (err)
            return true;
        console.log("excelData", results[1].excelData.length + "-" + results[0].excelData.length);
        var opArray = [];
        if (_.isEmpty(results[1].excelData)) {
            // sendMail("");
            console.log("no data found", new Date());
        } else {
            _.forEach(results[1].excelData, function (sqldata) {
                _.forEach(results[0].excelData, function (sapData) {
                    if (sqldata.empId === sapData.EMPLOYEE_ID) {
                        var opData = {
                            empId: sqldata.empId,
                            empName: sapData.EMPLOYEE_NAME
                        }

                        function checkData(str) {
                            if (_.isUndefined(sqldata[str]) || _.isEmpty(sqldata[str]) || _.isNull(sqldata[str])) {
                                opData[str] = "N/A";
                            } else if (sqldata[str] === "Success") {
                                if (sapData[str] == "Error" || sapData[str] == "N/A" || _.isEmpty(sapData[str])) {
                                    opData[str] = "Error";
                                } else {
                                    opData[str] = "Success";
                                }
                            }
                        }
                        checkData("IT580");
                        checkData("IT581");
                        checkData("IT584");
                        checkData("IT585");
                        checkData("IT586");
                        opArray.push(opData);

                    }
                })
            })
            excelGenerate(opArray, results[1].res);
        }
    });
// });

cron.schedule('0 8 * * *', () => {
    var con = mysql.createConnection({
        host: '10.91.4.52',
        port: '3306',
        user: 'bestseller',
        password: 'bestLLerlssd',
        database: "bestsellers_test"

    });
    con.connect(function (err) {
        if (err) throw err;

        console.log("Connected! sql");
        con.query("select count(distinct(unique_key)) from v_payment_operations;", function (err, result, fields) {
            if (err)
                callback(err, null);

            // console.log(result);

            var day = moment().subtract(1, 'days');
            const msg = {
                from: 'bwsupport@bestseller.com',
                to: ['praveen.ganji@bestseller.com', 'ankit.shah@bestseller.com'],
                cc: 'Ganesh.kothavale@bestseller.com',
                subject: 'MBC Reimbursement ' + day.format("DD MMM YYYY"),
                // text: 'Total IDoc count : ' + result['count(distinct(unique_key))']
            };

            if (result && result[0] && result[0]['count(distinct(unique_key))']) {
                msg.text = 'Total IDoc count ' + result[0]['count(distinct(unique_key))']
            } else {
                msg.text = 'Total IDoc count 0'
            }
            sgMail.send(msg);
        });
    });

    con.on('error', function (err) {
        if (err.code === 'PROTOCOL_CONNECTION_LOST' || err.code === 'PROTOCOL_CONNECTION_LOST') { // Connection to the MySQL server is usually
            // connectSQL();                          // lost due to either server restart, or a
        } else { // connnection idle timeout (the wait_timeout
            throw err; // server variable configures this)
        }
    })
});