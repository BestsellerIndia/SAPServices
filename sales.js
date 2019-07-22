const hanaClient = require("@sap/hana-client");
const connection = hanaClient.createConnection();
const connection1 = hanaClient.createConnection();
var mssql = require('mssql');
var moment = require("moment");
var fs = require('fs');
var async = require('async');
var _ = require('lodash');
const excel = require('node-excel-export');
const sgMail = require('@sendgrid/mail');
//sendgrid password
sgMail.setApiKey(process.env.SENDGRID_API_KEY);

var reqData = {
    DATE_FROM: moment().subtract(1, 'days').format('YYYYMMDD'),
    DATE_TO: moment().format('YYYYMMDD'),
}

function carData(callback) {
    const connectionParams = {
        host: "10.91.4.50",
        port: 30015,
        uid: "SUPPORT",
        pwd: "Best1234 ", // space is there 
        databaseName: "BP1"
    }

    connection.connect(connectionParams, (err) => {
        if (err) {
            callback(err, null);
            return console.error("Connection error", err);
        }

        const sql = `SELECT * FROM "_SYS_BIC"."ZBS_Retail/CAR_SALES" ('PLACEHOLDER' = ('$$date_to$$', '${reqData.DATE_TO}'), 'PLACEHOLDER' = ('$$date_from$$', '${reqData.DATE_FROM}'));`;

        connection.exec(sql, (err, rows) => {
            connection.disconnect();

            if (err) {
                callback(err, null);
                return console.error('SQL execute error:', err);
            }

            var jsonfile = JSON.stringify(rows);
            fs.writeFile('car.json', jsonfile, 'utf8', function (err, data) {

            });
            callback(null, rows);
        });
    });
}

function posData(config, jsonName, callback) {
    var connectionStr = {
        server: config.IP_ADDRESS,
        user: config.USERNAME,
        password: config.PASSWORD,
        // database: "bestsellers_test"
    }
    var conn = new mssql.ConnectionPool(connectionStr);
    conn.connect().then(function () {
            var req = new mssql.Request(conn);
            req.query(`use tpcentraldb SELECT lRetailStoreID as site , szDate as pos_date , sum(dTurnover) as pos_sales, sum(dTaQty) as pos_qty  FROM TxSaleLineItem where szDate between ${reqData.DATE_FROM} and ${reqData.DATE_TO} ` +
                `group by lRetailStoreID, szDate order by szDate`).then(function (records) {
                console.log(records.recordset.length);
                var jsonfile = JSON.stringify(records.recordset);
                callback(null, records.recordset);
                fs.writeFile('./pos_json/' + jsonName + '.json', jsonfile, 'utf8', function (err, data) {

                });
            }).catch(function (err) {
                callback(err, null);
            })
        })
        .catch(function (err) {
            callback(err, null);
        })
}

function excelGenerate(resData) {
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
        site: {
            displayName: 'STORE CODE',
            headerStyle: styles.headerDark,
            width: 80
        },
        sales_date: {
            displayName: 'SALES DATE',
            headerStyle: styles.headerDark,
            width: 80
        },
        pos_sales: {
            displayName: 'POS SALES',
            headerStyle: styles.headerDark,
            width: 80
        },
        pos_qty: {
            displayName: 'POS QTY',
            headerStyle: styles.headerDark,
            width: 80
        },
        car_sales: {
            displayName: 'CAR SALES',
            headerStyle: styles.headerDark,
            width: 80
        },
        car_qty: {
            displayName: 'CAR QTY',
            headerStyle: styles.headerDark,
            width: 80
        },
        sales_diff: {
            displayName: 'SALES DIFF.',
            headerStyle: styles.headerDark,
            width: 80
        },
        qty_diff: {
            displayName: 'QTY DIFF.',
            headerStyle: styles.headerDark,
            width: 80
        },
        connectionStatus: {
            displayName: 'Connection Status',
            headerStyle: styles.headerDark,
            width: 180
        },

    };


    var folder = "./";
    var path = "car_pos.xlsx";
    var finalPath = folder + path;
    const report = excel.buildExport(
        [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
            {
                name: 'car_pos', // <- Specify sheet name (optional)
                heading: [], // <- Raw heading array (optional)
                merges: [], // <- Merge cell ranges
                specification: specification, // <- Report specification
                data: resData // <-- Report data
            }
        ]
    )
    fs.writeFile(finalPath, report, "binary", function (err, data) {
        if (err) throw err;
        sendMail(finalPath);
    })
}

function sendMail(finalPath) {
    const msg = {
        from: 'bwsupport@bestseller.com',
        to: [
            'Ganesh.kothavale@bestseller.com'
            // 'abhishek.ghosh@bestseller.com',
            // 'ankit.shah@bestseller.com',
            // 'shreya.ambetkar@bestseller.com', 'ronak.pandya@bestseller.com'
        ],
        // cc: 'Ganesh.kothavale@bestseller.com',
        subject: 'pos and car details',
        text: 'Please find attachment for pos and car Report'
    };
    if (_.isEmpty(finalPath)) {
        msg.text = 'No pos and car Report Found';
    } else {
        msg.text = 'Please find attachment for pos and car Report';
        var file = fs.readFileSync(finalPath);
        var base64File = new Buffer(file).toString("base64");
        msg.attachments = [{
            content: base64File,
            filename: 'car_pos_' + moment().subtract(1, 'days').format("ll") + '.xlsx',
            type: 'plain/text',
            disposition: 'attachment',
            contentId: 'mytext'
        }];
    }

    // sgMail.send(msg);
    console.log("mail sent");
}
async.parallel([
        function (callback) {
            carData(callback);
        },
        function (callback) {
            async.waterfall([
                function (callback) {
                    // var allPos = [{
                    //         storeCode: 1508,
                    //         server: '10.91.171.134', //1508
                    //         user: 'sa',
                    //         password: 'retail!123',
                    //         // database: "bestsellers_test"
                    //     },
                    //     {
                    //         // server:'10.91.164.104',
                    //         storeCode: 1234,
                    //         server: '10.91', //1234
                    //         user: 'sa',
                    //         password: 'retail!123',
                    //         // database: "bestsellers_test"
                    //     }
                    // ]
                    const connectionParams1 = {
                        host: "10.91.4.50",
                        port: 30015,
                        uid: "SUPPORT",
                        pwd: "Best1234 ", // space is there 
                        databaseName: "BP1"
                    }

                    connection1.connect(connectionParams1, (err) => {
                        if (err) {
                            callback(err, null);
                            return console.error("Connection error", err);
                        }

                        const sql1 = `select * from "SAPCP1"."ZSTORE_INFO"`;

                        connection1.exec(sql1, (err, allPos) => {
                            connection.disconnect();

                            if (err) {
                                callback(err, null);
                                return console.error('SQL execute error:', err);
                            }

                            var jsonfile = JSON.stringify(allPos);
                            fs.writeFile('allpos.json', jsonfile, 'utf8', function (err, data) {

                            });
                            callback(null, allPos);
                        });
                    });

                },
                function (allPos, callback) {
                    async.concatLimit(allPos, 1, function (pos, callback) {
                        pos.count = 0;
                        var posCallback = function (err, data) {
                            pos.count++;
                            if (err) {
                                if (pos.count < 3) {
                                    console.log("count :", pos.count);
                                    posData(pos, pos.IP_ADDRESS.replace(/\./g, "_"), posCallback);
                                } else {
                                    console.log("in err");
                                    callback(null, {
                                        site: pos.SITE,
                                        pos_date: reqData.DATE_FROM,
                                        pos_sales: 0,
                                        pos_qty: 0,
                                        connectionStatus: 'Failed'
                                    }, {
                                        site: pos.SITE,
                                        pos_date: reqData.DATE_TO,
                                        pos_sales: 0,
                                        pos_qty: 0,
                                        connectionStatus: 'Failed'
                                    });
                                }
                            } else {
                                console.log("in success")
                                callback(null, data);
                            }
                        };
                        posData(pos, pos.IP_ADDRESS.replace(/\./g, "_"), posCallback);

                    }, callback)
                }
            ], callback)
        }
    ],
    function (err, results) {
        // if (err)
        //     return true;

        // console.log("count of car : ", results[0]);
        // console.log("count of pos : ", results[1]);

        _.forEach(results[0], function (car) {
            _.forEach(results[1], function (pos) {
                var pos_date = pos.pos_date.slice(0, 4) + "-" + pos.pos_date.slice(4, 6) + "-" + pos.pos_date.slice(6, 8);
                if (pos.SITE == 1234 && car.site == "1234" && car.sales_date == 20190711)
                    console.log('i found you', pos_date, car.sales_date);
                if (car.site.toString() === pos.site.toString() &&
                    new Date(car.sales_date).setHours(0, 0, 0, 0) === new Date(pos_date).setHours(0, 0, 0, 0)
                ) {
                    car.pos_sales = pos.pos_sales;
                    car.pos_qty = pos.pos_qty;
                    car.sales_diff = car.pos_sales - car.car_sales;
                    car.qty_diff = car.pos_qty - car.car_qty;
                    car.connectionStatus = pos.connectionStatus === 'Failed' ? 'Failed' : 'Success';
                }
            })
        })

        excelGenerate(results[0]);
    });