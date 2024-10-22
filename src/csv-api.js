import { MongoClient, ObjectId, ServerApiVersion } from "mongodb";
import "./loadenv.js";
import consts from "./consts.js";
import { addFileToApp, addFileToAppForVersionedApp, removeFileFromApp, removeFileFromVersionedApp, } from "./fileMgmt.js";
import fs from "fs";
import { deleteFileFromAzureBlobStorage, getBlobAsStream, } from "./blobUtils.js";
import { BlobServiceClient } from "@azure/storage-blob";
import { toArrayBuffer } from "./workRequests.js";
import db from "./mdbUtils.js";
import ExcelJS from 'exceljs';
import path from 'path';
import { name } from "ejs";

const uri = process.env.MONGODB_SRV_URI;

const client = new MongoClient(uri, {
    serverApi: {
        version: ServerApiVersion.v1,
        strict: true,
        deprecationErrors: true,
    },
});

await client.connect();

// export const processCsvUpload = async (body, files) => {
//     const collection = db.collection(consts.MDB_CSV);
//     const fileInfo = {
//         ...body,
//         uploadedAt: new Date(),
//         status: 'Processed',
//     };

//     if (files && files.file) {
//         const file = files.file;
//         console.log("Uploaded file:", file.name);

//         // Add file information to fileInfo
//         fileInfo.fileName = file.name;
//         fileInfo.fileSize = file.size;
//         fileInfo.fileType = file.type;

//         // Insert document into database
//         const result = await collection.insertOne(fileInfo);

//         // Process the single file
//         const fileName = encodeURIComponent(result.insertedId.toString()) + '/' + encodeURIComponent(file.name);
//         const clientFileBuffer = fs.readFileSync(file.path);

//         await addFileToApp(
//             consts.MDB_CSV,
//             consts.BLOB_CSV,
//             result.insertedId.toString(),
//             fileName,
//             file.name,
//             clientFileBuffer
//         );

//         // Optionally, remove the temporary file
//         fs.unlinkSync(file.path);

//         return {
//             message: 'CSV file processed successfully',
//             fileName: file.name
//         };
//     } else {
//         console.log("No file received");
//         return {
//             message: 'No file received',
//         };
//     }
// };

// //getVessels
// export const getVessels = async () => {
//     const collection = db.collection(consts.MDB_MARINE_VESSELS);

//     const vessels = await collection.find({
//         $and: [
//             { imoNumber: { $ne: null, $exists: true } },
//             { vesselName: { $ne: null, $exists: true } },
//             { imoNumber: { $ne: "" } },
//             { vesselName: { $ne: "" } }
//         ]
//     }).toArray();

//     return vessels;
// };

export const checkOFAC = async (name, imoNumber) => {
    // console.log("imoNumber:::", imoNumber);
    let myHeaders = new Headers()
    myHeaders.append('Content-Type', 'application/json')
    myHeaders.append('apikey', process.env.OFAC_API_KEY)
    let raw = {
        apiKey: process.env.OFAC_API_KEY,
        sources: ['sdn', 'nonsdn', 'eu'],
        cases: [
            {
                "name": name,
                "type": "vessel",
                "ids": [{
                    "id": imoNumber
                }]
            }
        ]
    }

    let requestOptions = {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            apikey: process.env.OFAC_API_KEY
        },
        body: JSON.stringify(raw),
        redirect: 'follow'
    }
    const apiUrl = process.env.OFAC_BASE_URL + '/v4/search'
    let data = await fetch(apiUrl, requestOptions)
    const ofacResponse = await data.json()
    console.debug('ofac response:', ofacResponse)
    return ofacResponse
}

//screenVessels
export const screenVessels = async (body) => {
    const collection = db.collection(consts.MDB_MARINE_VESSELS);
    const bodyObject = typeof body === 'string' ? JSON.parse(body) : body;
    // console.log("1. Direct access:", bodyObject.imoNumber);


    const ofacResp = await checkOFAC(bodyObject.vesselName, bodyObject.imoNumber);
    // //find vessel by body._id
    // const vessel = await collection.findOne({ _id: new ObjectId(bodyObject._id) });
    // console.log("vessel:", vessel);
    //update vessel with ofac response
    await collection.updateOne(
        { _id: new ObjectId(bodyObject._id) },
        { $set: { OFAC: ofacResp } }
    );
}


export const getVessels = async () => {
    const vesselsCollection = db.collection(consts.MDB_MARINE_VESSELS);

    const vesselsWithClaims = await vesselsCollection.aggregate([
        // Match vessels with valid imoNumber and vesselName
        {
            $match: {
                imoNumber: { $type: "number", $ne: null },
                vesselName: { $type: "string", $ne: "" }
            }
        },
        // Lookup claims for each vessel
        {
            $lookup: {
                from: consts.MDB_MARINE_CLAIMS,
                let: { vesselImo: "$imoNumber" },
                pipeline: [
                    {
                        $match: {
                            $expr: { $eq: ["$imoNumber", "$$vesselImo"] }
                        }
                    },
                    {
                        $project: {
                            _id: 1,
                            claimId: 1,
                            lossDescription: 1,
                            dateOfLoss: 1,
                            underwritingYear: 1,
                            claimAmount: 1,
                            currency: 1,
                            imoNumber: 1
                        }
                    }
                ],
                as: "claims"
            }
        },
        // Project the final output
        {
            $project: {
                _id: 1,
                OFAC: 1,
                imoNumber: 1,
                vesselName: 1,
                claims: 1,
                claimCount: { $size: "$claims" }
            }
        }
    ]).toArray();

    return vesselsWithClaims;
};
export const processCsvUpload = async (body, files) => {
    const collection = db.collection(consts.MDB_CSV);
    const fileInfo = {
        ...body,
        uploadedAt: new Date(),
        status: 'Processing',
    };

    if (files && files.file) {
        const file = files.file;
        console.log("Uploaded file:", file.name);

        // Add file information to fileInfo
        fileInfo.fileName = file.name;
        fileInfo.fileSize = file.size;
        fileInfo.fileType = file.type;

        // Insert document into database
        const result = await collection.insertOne(fileInfo);

        // Process the Excel file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(file.path);

        // Create a new worksheet for the summary
        // const summarySheet = workbook.addWorksheet('Summary USD');

        // Implement the macro logic here
        const summarySheet = await processExcelFile(workbook);
        // console.log("summarySheet:", summarySheet);
        // Save the updated workbook
        // const updatedExcelPath = path.join(path.dirname(file.path), 'updated_' + file.name);
        // await workbook.xlsx.writeFile(updatedExcelPath);

        // Generate CSV from the summary sheet
        const csvContent = await generateCsvBufferFromString(summarySheet);
        const csvFileName = 'summary.csv';
        // const csvPath = path.join(path.dirname(file.path), csvFileName);
        // fs.writeFileSync(csvPath, csvContent);

        // Upload both files using addFileToApp
        const excelFileName = encodeURIComponent(result.insertedId.toString()) + '/' + encodeURIComponent(file.name);
        const csvFileNameEncoded = encodeURIComponent(result.insertedId.toString()) + '/' + encodeURIComponent(csvFileName);

        // Write CSV buffer to a temporary file
        const tempDir = path.join(process.cwd(), 'temp');
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir);
        }
        const csvTempPath = path.join(tempDir, csvFileName);
        fs.writeFileSync(csvTempPath, csvContent);

        const excelBuffer = fs.readFileSync(file.path);
        const csvBuffer = fs.readFileSync(csvTempPath);

        await addFileToApp(
            consts.MDB_CSV,
            consts.BLOB_CSV,
            result.insertedId.toString(),
            excelFileName,
            file.name,
            excelBuffer
        );

        await addFileToApp(
            consts.MDB_CSV,
            consts.BLOB_CSV,
            result.insertedId.toString(),
            csvFileNameEncoded,
            csvFileName,
            csvBuffer
        );

        // Update status in the database
        await collection.updateOne(
            { _id: result.insertedId },
            { $set: { status: 'Processed' } }
        );

        // Clean up temporary files
        fs.unlinkSync(file.path);
        // fs.unlinkSync(updatedExcelPath);
        fs.unlinkSync(csvTempPath);

        return {
            message: 'Excel file processed and CSV summary generated successfully',
            excelFileName: file.name,
            csvFileName: csvFileName
        };
    } else {
        console.log("No file received");
        return {
            message: 'No file received',
        };
    }
};

export const processSummary = async (body, files) => {
    const collection = db.collection(consts.MDB_CSV);
    let fileInfo = {
        ...body,
        uploadedAt: new Date(),
        status: 'Processing',
    };
    // console.log("fileInfo:", body)
    let type = '';
    // if body.reportType === 'claims_paid' then type equals Claims and if premiums premiums_paid tjhen type equals Premiums
    if (body.reportType === 'claims_paid') {
        type = 'Claims';
    } else if (body.reportType === 'premiums_paid') {
        type = 'Premiums';
    }


    if (files && files.file) {
        const file = files.file;

        delete fileInfo.reportType;

        // Check if a document with the same reporting period exists
        const existingDocument = await collection.findOne({
            reportingPeriod: body.reportingPeriod,
            businessUnit: body.businessUnit
        });
        if (existingDocument) {
            // Check if the document already has the current report type
            if (existingDocument[body.reportType]) {
                console.log(":::HEREEE")
                return {
                    message: `This report has already been uploaded for the period ${body.reportingPeriod}`,
                };
            } else {
                console.log(":::Not an existing")
                // console.log("Claims:", existingDocument.claims_paid);
                // console.log("Premiums:", existingDocument.premium_paid);

                // Insert document into database
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(file.path);
                let summarySheet;
                let csvFileName;
                let uploadedExcelName;
                let mgObj;
                if (type === 'Claims') {
                    // summarySheet = (await claimsSummary(workbook)).csvString;
                    // mgObj = (await claimsSummary(workbook)).structuredObject;
                    const { csvString, structuredObject } = await claimsSummary(workbook);
                    summarySheet = csvString;
                    mgObj = structuredObject;

                    fileInfo = { ...fileInfo, claims_paid: mgObj };
                    csvFileName = 'claimsummary.csv';
                    uploadedExcelName = 'uploadedClaimExcel.xlsx';
                    await collection.updateOne(
                        { _id: existingDocument._id },
                        { $set: { claims_paid: mgObj, status: 'processed' } }
                    );
                } else if (type === 'Premiums') {
                    // summarySheet = (await premiumSummary(workbook)).csvString;
                    // mgObj = (await premiumSummary(workbook)).mongoDbData;

                    const { csvString, mongoDbData } = await premiumSummary(workbook);
                    summarySheet = csvString;
                    mgObj = mongoDbData;
                    fileInfo = { ...fileInfo, premium_paid: mgObj };
                    console.log("mgObj:", mgObj);
                    csvFileName = 'premiumsummary.csv';
                    uploadedExcelName = 'uploadedPremiumExcel.xlsx';
                    await collection.updateOne(
                        { _id: existingDocument._id },
                        { $set: { premium_paid: mgObj, status: 'processed' } });
                }
                // console.log("fileInfo:", fileInfo);
                // console.log("ReportType:", body.reportType);
                // Update the existing document


                // const summarySheet = await claimsSummary(workbook);

                const csvContent = await generateCsvBufferFromString(summarySheet);



                // Upload both files using addFileToApp
                const excelFileName = encodeURIComponent(existingDocument._id.toString()) + '/' + encodeURIComponent(file.name);
                const csvFileNameEncoded = encodeURIComponent(existingDocument._id.toString()) + '/' + encodeURIComponent(csvFileName);

                // Write CSV buffer to a temporary file
                const tempDir = path.join(process.cwd(), 'temp');
                if (!fs.existsSync(tempDir)) {
                    fs.mkdirSync(tempDir);
                }
                const csvTempPath = path.join(tempDir, csvFileName);
                fs.writeFileSync(csvTempPath, csvContent);

                const excelBuffer = fs.readFileSync(file.path);
                const csvBuffer = fs.readFileSync(csvTempPath);

                await addFileToApp(
                    consts.MDB_CSV,
                    consts.BLOB_CSV,
                    existingDocument._id.toString(),
                    excelFileName,
                    uploadedExcelName,
                    excelBuffer
                );

                await addFileToApp(
                    consts.MDB_CSV,
                    consts.BLOB_CSV,
                    existingDocument._id.toString(),
                    csvFileNameEncoded,
                    csvFileName,
                    csvBuffer
                );


                //get doc2 from collection
                let doc3 = await collection.findOne({ _id: existingDocument._id });

                let resp = await generateStatementOfAccounts(doc3);
                let resp2 = await generateEURStatementOfAccounts(doc3);
                console.log("resp:", resp, "resp2:", resp2);
                return {
                    message: `This report has processed`,
                };
            }
        } else {
            // Insert document into database
            let result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(file.path);
            let summarySheet;
            let csvFileName;
            let uploadedExcelName;
            let mgObj;
            if (type === 'Claims') {

                summarySheet = (await claimsSummary(workbook)).csvString;
                mgObj = (await claimsSummary(workbook)).structuredObject;
                fileInfo = { ...fileInfo, claims_paid: mgObj, status: 'Pending Premiums Paid' };
                result = await collection.insertOne(fileInfo)
                csvFileName = 'claimsummary.csv';
                uploadedExcelName = 'uploadedClaimExcel.xlsx';
            } else if (type === 'Premiums') {
                summarySheet = (await premiumSummary(workbook)).csvString;
                mgObj = (await premiumSummary(workbook)).mongoDbData;
                fileInfo = { ...fileInfo, premium_paid: mgObj, status: 'Pending Claims Paid' };
                result = await collection.insertOne(fileInfo)
                console.log("mgObj:", mgObj);
                csvFileName = 'premiumsummary.csv';
                uploadedExcelName = 'uploadedPremiumExcel.xlsx';
            }
            // const summarySheet = await claimsSummary(workbook);

            const csvContent = await generateCsvBufferFromString(summarySheet);



            // Upload both files using addFileToApp
            const excelFileName = encodeURIComponent(result.insertedId.toString()) + '/' + encodeURIComponent(file.name);
            const csvFileNameEncoded = encodeURIComponent(result.insertedId.toString()) + '/' + encodeURIComponent(csvFileName);

            // Write CSV buffer to a temporary file
            const tempDir = path.join(process.cwd(), 'temp');
            if (!fs.existsSync(tempDir)) {
                fs.mkdirSync(tempDir);
            }
            const csvTempPath = path.join(tempDir, csvFileName);
            fs.writeFileSync(csvTempPath, csvContent);

            const excelBuffer = fs.readFileSync(file.path);
            const csvBuffer = fs.readFileSync(csvTempPath);

            await addFileToApp(
                consts.MDB_CSV,
                consts.BLOB_CSV,
                result.insertedId.toString(),
                excelFileName,
                uploadedExcelName,
                excelBuffer
            );

            await addFileToApp(
                consts.MDB_CSV,
                consts.BLOB_CSV,
                result.insertedId.toString(),
                csvFileNameEncoded,
                csvFileName,
                csvBuffer
            );

            // Clean up temporary files
            fs.unlinkSync(file.path);
            // fs.unlinkSync(updatedExcelPath);
            fs.unlinkSync(csvTempPath);

            return {
                message: 'Excel file processed and CSV summary generated successfully',
                excelFileName: file.name,
                csvFileName: csvFileName
            };
        }
    } else {
        console.log("No file received");
        return {
            message: 'No file received',
        };
    }
};



export const premiumSummary = async (input) => {
    let workbook;

    if (input instanceof ExcelJS.Workbook) {
        workbook = input;
    } else if (typeof input === 'string' || input instanceof Buffer) {
        workbook = new ExcelJS.Workbook();
        if (typeof input === 'string') {
            await workbook.xlsx.readFile(input);
        } else {
            await workbook.xlsx.load(input);
        }
    }

    const entityData = new Map();
    const mongoDbData = {
        totals: {
            USD: { orderGrossPremium: 0, transactionAmount: 0 },
            EUR: { orderGrossPremium: 0, transactionAmount: 0 }
        },
        entities: []
    };

    const getCellValue = (cell) => {
        if (cell && cell.formula) {
            console.log(`Formula found: ${cell.formula}`);
            return cell.result;
        }
        return cell && cell.value;
    };

    workbook.eachSheet((worksheet, sheetId) => {
        const sheetName = worksheet.name;
        console.log(`\nProcessing sheet: ${sheetName}`);

        entityData.set(sheetName, new Map());

        let lastRowWithFormulas = 0;
        let headerRowIndex = 0;
        let currencyType = '';

        // Find the header row and the last row with formulas
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (row.values.some(cell => cell && cell.formula)) {
                lastRowWithFormulas = rowNumber;
            }
            if (headerRowIndex === 0 && getCellValue(row.getCell('Q'))) {
                headerRowIndex = rowNumber;
            }
        });

        console.log(`Last row with formulas: ${lastRowWithFormulas}`);
        console.log(`Header row index: ${headerRowIndex}`);

        if (lastRowWithFormulas > 0) {
            const lastFormulaRow = worksheet.getRow(lastRowWithFormulas);

            // Search for 'EUR' or 'USD' in the last row with formulas
            lastFormulaRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const cellValue = getCellValue(cell);
                if (typeof cellValue === 'string') {
                    if (cellValue.includes('EUR')) {
                        currencyType = 'EUR';
                    } else if (cellValue.includes('USD')) {
                        currencyType = 'USD';
                    }
                }
            });

            // If currency not found, search for 'Total' and then check the column
            if (!currencyType) {
                lastFormulaRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    const cellValue = getCellValue(cell);
                    if (typeof cellValue === 'string' && cellValue.toLowerCase().includes('total')) {
                        // Found 'Total', now check this column for currency
                        for (let rowIndex = lastRowWithFormulas; rowIndex >= headerRowIndex; rowIndex--) {
                            const checkCell = worksheet.getCell(rowIndex, colNumber);
                            const checkCellValue = getCellValue(checkCell);
                            if (typeof checkCellValue === 'string') {
                                if (checkCellValue.includes('EUR')) {
                                    currencyType = 'EUR';
                                    break;
                                } else if (checkCellValue.includes('USD')) {
                                    currencyType = 'USD';
                                    break;
                                }
                            }
                        }
                        if (currencyType) return false; // Stop searching if currency found
                    }
                });
            }
        }
        if (lastRowWithFormulas > 0 && headerRowIndex > 0) {
            const lastRow = worksheet.getRow(lastRowWithFormulas);
            const headerRow = worksheet.getRow(headerRowIndex);
            const uwYearRow = worksheet.getRow(headerRowIndex + 1);
            // Get total values from columns O and P
            const totalOrderGrossPremium = parseFloat(getCellValue(lastRow.getCell('O'))) || 0;
            const totalTransactionAmount = parseFloat(getCellValue(lastRow.getCell('P'))) || 0;

            console.log(`${sheetName} (${currencyType}) - Total Order Gross Premium: ${totalOrderGrossPremium.toFixed(2)}, Total Transaction Amount: ${totalTransactionAmount.toFixed(2)}`);

            // Update the entityData with the total values
            entityData.get(sheetName).set('Total', {
                orderGrossPremium: totalOrderGrossPremium,
                transactionAmount: totalTransactionAmount,
                currency: currencyType
            });

            // Process data for individual UW years
            let currentUWYear = '';

            headerRow.eachCell({ includeEmpty: false }, (headerCell, colNumber) => {
                const headerValue = getCellValue(headerCell);
                const column = headerCell.address.replace(/\d+/, '');

                const uwYearCell = uwYearRow.getCell(column);
                const uwYearValue = getCellValue(uwYearCell);

                if (uwYearValue && uwYearValue.toString().startsWith('U/W')) {
                    currentUWYear = uwYearValue.toString().replace('U/W', '').trim();
                    if (!entityData.get(sheetName).has(currentUWYear)) {
                        entityData.get(sheetName).set(currentUWYear, { orderGrossPremium: 0, transactionAmount: 0, currency: currencyType });
                    }
                }

                const value = parseFloat(getCellValue(lastRow.getCell(column))) || 0;

                if (currentUWYear) {
                    const data = entityData.get(sheetName).get(currentUWYear);
                    if (headerValue.toLowerCase().includes('order gross premium')) {
                        data.orderGrossPremium = value;
                        console.log(`${sheetName} (${currencyType}) - Order Gross Premium ${currentUWYear}: ${value.toFixed(2)}`);
                    } else if (headerValue.toLowerCase().includes('transaction amount')) {
                        data.transactionAmount = value;
                        console.log(`${sheetName} (${currencyType}) - Transaction Amount ${currentUWYear}: ${value.toFixed(2)}`);
                    }
                }
            });
        }
        console.log('---');
    });

    // Calculate totals for USD and EUR
    let totalUSDOrderGrossPremium = 0;
    let totalUSDTransactionAmount = 0;
    let totalEUROrderGrossPremium = 0;
    let totalEURTransactionAmount = 0;

    entityData.forEach((entityYearData, entityName) => {
        const totalData = entityYearData.get('Total');
        if (totalData.currency === 'USD') {
            totalUSDOrderGrossPremium += totalData.orderGrossPremium;
            totalUSDTransactionAmount += totalData.transactionAmount;
        } else if (totalData.currency === 'EUR') {
            totalEUROrderGrossPremium += totalData.orderGrossPremium;
            totalEURTransactionAmount += totalData.transactionAmount;
        }
    });

    // Update mongoDbData totals
    mongoDbData.totals.USD.orderGrossPremium = totalUSDOrderGrossPremium.toFixed(2);
    mongoDbData.totals.USD.transactionAmount = totalUSDTransactionAmount.toFixed(2);
    mongoDbData.totals.EUR.orderGrossPremium = totalEUROrderGrossPremium.toFixed(2);
    mongoDbData.totals.EUR.transactionAmount = totalEURTransactionAmount.toFixed(2);

    // Populate the entities array
    entityData.forEach((entityYearData, entityName) => {
        const entity = {
            entity: entityName,
            currency: '',
            uwYears: [],
            totalOrderGrossPremium: '',
            totalTransactionAmount: ''
        };

        entityYearData.forEach((data, uwYear) => {
            if (uwYear === 'Total') {
                entity.totalOrderGrossPremium = data.orderGrossPremium.toFixed(2);
                entity.totalTransactionAmount = data.transactionAmount.toFixed(2);
                entity.currency = data.currency;
            } else {
                entity.currency = data.currency;
                entity.uwYears.push({
                    year: uwYear,
                    orderGrossPremium: data.orderGrossPremium !== 0 ? data.orderGrossPremium.toFixed(2) : '-',
                    transactionAmount: data.transactionAmount !== 0 ? data.transactionAmount.toFixed(2) : '-'
                });
            }
        });

        mongoDbData.entities.push(entity);
    });

    // Generate CSV content
    let csvContent = [];

    // Add totals for USD and EUR
    csvContent.push(['TOTAL (Order Gross Premium) USD', `USD ${totalUSDOrderGrossPremium.toFixed(2)}`]);
    csvContent.push(['TOTAL (Transaction Amount) USD', `USD ${totalUSDTransactionAmount.toFixed(2)}`]);
    csvContent.push(['TOTAL (Order Gross Premium) EUR', `EUR ${totalEUROrderGrossPremium.toFixed(2)}`]);
    csvContent.push(['TOTAL (Transaction Amount) EUR', `EUR ${totalEURTransactionAmount.toFixed(2)}`]);
    csvContent.push([]);

    // Add headers
    csvContent.push(['Entity', 'UW Year', 'Order Gross Premium', 'Transaction Amount']);

    // Fill data
    entityData.forEach((entityYearData, entityName) => {
        const sortedYears = Array.from(entityYearData.keys()).filter(year => year !== 'Total').sort();
        sortedYears.forEach(uwYear => {
            const data = entityYearData.get(uwYear);
            csvContent.push([
                entityName,
                uwYear,
                data.orderGrossPremium !== 0 ? `${data.currency} ${data.orderGrossPremium.toFixed(2)}` : '-',
                data.transactionAmount !== 0 ? `${data.currency} ${data.transactionAmount.toFixed(2)}` : '-'
            ]);
        });
    });

    // Add Grand Total
    csvContent.push(['Grand Total', '', `USD ${totalUSDOrderGrossPremium.toFixed(2)} / EUR ${totalEUROrderGrossPremium.toFixed(2)}`, `USD ${totalUSDTransactionAmount.toFixed(2)} / EUR ${totalEURTransactionAmount.toFixed(2)}`]);

    // Convert CSV content to string
    const csvString = csvContent.map(row => row.join(';')).join('\n');

    // Return both the CSV string and the MongoDB-ready data
    return { csvString, mongoDbData };
};

function calculateUSDClaimsByYear(claimsData) {
    const yearTotals = {
        '19/20': 0,
        '20/21': 0,
        '21/22': 0,
        '22/23': 0,
        '23/24': 0
    };

    claimsData.entities.forEach(entity => {
        entity.uwYears.forEach(yearData => {
            if (yearData.USD) {
                const year = yearData.year;
                const usdValue = parseFloat(yearData.USD) || 0;
                if (yearTotals.hasOwnProperty(year)) {
                    yearTotals[year] += usdValue;
                }
            }
        });
    });

    return yearTotals;
}
function calculateEURClaimsByYear(claimsData) {
    const yearTotals = {
        '19/20': 0,
        '20/21': 0,
        '21/22': 0,
        '22/23': 0,
        '23/24': 0
    };

    claimsData.entities.forEach(entity => {
        entity.uwYears.forEach(yearData => {
            if (yearData.EUR) {
                const year = yearData.year;
                const eurValue = parseFloat(yearData.EUR) || 0;
                if (yearTotals.hasOwnProperty(year)) {
                    yearTotals[year] += eurValue;
                }
            }
        });
    });

    return yearTotals;
}

async function generateStatementOfAccounts(mongoDbData) {
    // console.log('Generating', mongoDbData);
    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Statement of Accounts');
    // console.log('Starting to process entities data');
    // console.log('Number of entities:', mongoDbData.premium_paid.entities.length);
    // Set up headers
    worksheet.columns = [
        { header: 'Entity', key: 'entity', width: 30 },
        { header: 'Total Transaction Amount (TA)', key: 'totalTransactionAmount', width: 25 },
        { header: 'Total Order Gross Premium (GP)', key: 'totalOrderGrossPremium', width: 25 },
        { header: '19/20 (TA)', key: 'uw1920', width: 15 },
        { header: '20/21 (TA)', key: 'uw2021', width: 15 },
        { header: '21/22 (TA)', key: 'uw2122', width: 15 },
        { header: '22/23 (TA)', key: 'uw2223', width: 15 },
        { header: '23/24 (TA)', key: 'uw2324', width: 15 },
        { header: '19/20 (GP)', key: 'GPuw1920', width: 15 },
        { header: '20/21 (GP)', key: 'GPuw2021', width: 15 },
        { header: '21/22 (GP)', key: 'GPuw2122', width: 15 },
        { header: '22/23 (GP)', key: 'GPuw2223', width: 15 },
        { header: '23/24 (GP)', key: 'GPuw2324', width: 15 }
    ];
    const usdEntities = mongoDbData.premium_paid.entities.filter(entity => entity.currency === 'USD');

    // Initialize totals object for both transaction amounts and gross premiums
    const yearlyTotals = {
        uw1920: 0, uw2021: 0, uw2122: 0, uw2223: 0, uw2324: 0,
        GPuw1920: 0, GPuw2021: 0, GPuw2122: 0, GPuw2223: 0, GPuw2324: 0
    };

    // Process entities data
    usdEntities.forEach((entity, index) => {
        const row = {
            entity: entity.entity,
            totalTransactionAmount: entity.totalTransactionAmount,
            totalOrderGrossPremium: entity.totalOrderGrossPremium,
            uw1920: "-", uw2021: "-", uw2122: "-", uw2223: "-", uw2324: "-",
            GPuw1920: "-", GPuw2021: "-", GPuw2122: "-", GPuw2223: "-", GPuw2324: "-"
        };

        // Fill in underwriting years for both transaction amounts and gross premiums
        entity.uwYears.forEach(uwYear => {
            const yearKey = `uw${uwYear.year.replace('/', '')}`;
            const gpYearKey = `GPuw${uwYear.year.replace('/', '')}`;
            console.log("YEAR KEY:", yearKey);
            console.log("GP YEAR KEY:", gpYearKey);
            // Process transaction amount
            const transactionAmount = parseFloat(uwYear.transactionAmount) || 0;
            row[yearKey] = transactionAmount !== 0 ? transactionAmount.toString() : "-";
            yearlyTotals[yearKey] += transactionAmount;

            // Process order gross premium
            const grossPremium = parseFloat(uwYear.orderGrossPremium) || 0;
            row[gpYearKey] = grossPremium !== 0 ? grossPremium.toString() : "-";
            yearlyTotals[gpYearKey] += grossPremium;
        }); console.log('Row data:', row); // Log each row for debugging


        worksheet.addRow(row);
    });

    // Create the totals row
    const totalsRow = {
        entity: 'TOTAL',
        totalTransactionAmount: parseFloat(mongoDbData.premium_paid.totals.USD.transactionAmount).toFixed(2),
        totalOrderGrossPremium: parseFloat(mongoDbData.premium_paid.totals.USD.orderGrossPremium).toFixed(2),
        uw1920: yearlyTotals.uw1920 !== 0 ? parseFloat(yearlyTotals.uw1920).toFixed(2) : "-",
        uw2021: yearlyTotals.uw2021 !== 0 ? parseFloat(yearlyTotals.uw2021).toFixed(2) : "-",
        uw2122: yearlyTotals.uw2122 !== 0 ? parseFloat(yearlyTotals.uw2122).toFixed(2) : "-",
        uw2223: yearlyTotals.uw2223 !== 0 ? parseFloat(yearlyTotals.uw2223).toFixed(2) : "-",
        uw2324: yearlyTotals.uw2324 !== 0 ? parseFloat(yearlyTotals.uw2324).toFixed(2) : "-",
        GPuw1920: yearlyTotals.GPuw1920 !== 0 ? parseFloat(yearlyTotals.GPuw1920).toFixed(2) : "-",
        GPuw2021: yearlyTotals.GPuw2021 !== 0 ? parseFloat(yearlyTotals.GPuw2021).toFixed(2) : "-",
        GPuw2122: yearlyTotals.GPuw2122 !== 0 ? parseFloat(yearlyTotals.GPuw2122).toFixed(2) : "-",
        GPuw2223: yearlyTotals.GPuw2223 !== 0 ? parseFloat(yearlyTotals.GPuw2223).toFixed(2) : "-",
        GPuw2324: yearlyTotals.GPuw2324 !== 0 ? parseFloat(yearlyTotals.GPuw2324).toFixed(2) : "-"
    };
    worksheet.addRow(totalsRow);


    // console.log('Calculating claims by year');
    const claimsByYear = calculateUSDClaimsByYear(mongoDbData.claims_paid);
    // console.log('Claims by year:', claimsByYear);

    const claimsRow = {
        entity: 'Claims',
        totalTransactionAmount: -mongoDbData.claims_paid.totals.USD,
        totalOrderGrossPremium: "-",
        uw1920: -claimsByYear['19/20'],
        uw2021: -claimsByYear['20/21'],
        uw2122: -claimsByYear['21/22'],
        uw2223: -claimsByYear['22/23'],
        uw2324: -claimsByYear['23/24']
    };
    // console.log('Claims row:', claimsRow);



    worksheet.addRow(claimsRow);


    // Create the grand total row
    const grandTotalRow = {
        entity: 'GRAND TOTAL',
        totalTransactionAmount: '',
        totalOrderGrossPremium: '-',
        uw1920: '-',
        uw2021: '-',
        uw2122: '-',
        uw2223: '-',
        uw2324: '-',
        GPuw1920: '-',
        GPuw2021: '-',
        GPuw2122: '-',
        GPuw2223: '-',
        GPuw2324: '-'
    };

    // Calculate grand totals
    const calculateGrandTotal = (totalValue, claimValue) => {
        if (totalValue === '-' && claimValue === '-') return '-';
        const total = parseFloat(totalValue) || 0;
        const claim = parseFloat(claimValue) || 0;
        // Since claims are already negative, we add them
        return (total + claim).toFixed(2);
    };

    // Assuming 'totalsRow' is your total row and 'claimsRow' is your claims row
    grandTotalRow.totalTransactionAmount = calculateGrandTotal(totalsRow.totalTransactionAmount, claimsRow.totalTransactionAmount);
    // grandTotalRow.totalOrderGrossPremium = calculateGrandTotal(totalsRow.totalOrderGrossPremium, claimsRow.totalOrderGrossPremium);

    // Calculate grand totals for each year
    ['uw1920', 'uw2021', 'uw2122', 'uw2223', 'uw2324'].forEach(year => {
        grandTotalRow[year] = calculateGrandTotal(totalsRow[year], claimsRow[year]);
    });

    // Add the grand total row to the worksheet
    worksheet.addRow(grandTotalRow)


    // Apply some styling
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(worksheet.rowCount).font = { bold: true };

    // Generate Excel file
    // const buffer = await workbook.xlsx.writeBuffer();

    // Generate CSV content
    const csvContent = worksheet.getSheetValues().map(row => row.join(',')).join('\n');
    const csvBuffer = await generateCsvBufferFromString(csvContent);
    const statementCsvName = 'statement_of_accounts.csv';
    const statementNameEncoded = encodeURIComponent(mongoDbData._id.toString()) + '/' + encodeURIComponent(statementCsvName);
    await addFileToApp(
        consts.MDB_CSV,
        consts.BLOB_CSV,
        mongoDbData._id.toString(),
        statementNameEncoded,
        statementCsvName,
        csvBuffer
    );
    // console.log(csvContent, "::Contnet")
    return { message: "Statement of Accounts Generated Successfully" };
}
async function generateEURStatementOfAccounts(mongoDbData) {
    // console.log('Generating', mongoDbData);
    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Statement of Accounts');
    // console.log('Starting to process entities data');
    // console.log('Number of entities:', mongoDbData.premium_paid.entities.length);
    // Set up headers
    worksheet.columns = [
        { header: 'Entity', key: 'entity', width: 30 },
        { header: 'Total Transaction Amount (TA)', key: 'totalTransactionAmount', width: 25 },
        { header: 'Total Order Gross Premium (GP)', key: 'totalOrderGrossPremium', width: 25 },
        { header: '19/20 (TA)', key: 'uw1920', width: 15 },
        { header: '20/21 (TA)', key: 'uw2021', width: 15 },
        { header: '21/22 (TA)', key: 'uw2122', width: 15 },
        { header: '22/23 (TA)', key: 'uw2223', width: 15 },
        { header: '23/24 (TA)', key: 'uw2324', width: 15 },
        { header: '19/20 (GP)', key: 'GPuw1920', width: 15 },
        { header: '20/21 (GP)', key: 'GPuw2021', width: 15 },
        { header: '21/22 (GP)', key: 'GPuw2122', width: 15 },
        { header: '22/23 (GP)', key: 'GPuw2223', width: 15 },
        { header: '23/24 (GP)', key: 'GPuw2324', width: 15 }
    ];
    const usdEntities = mongoDbData.premium_paid.entities.filter(entity => entity.currency === 'EUR');

    // Initialize totals object for both transaction amounts and gross premiums
    const yearlyTotals = {
        uw1920: 0, uw2021: 0, uw2122: 0, uw2223: 0, uw2324: 0,
        GPuw1920: 0, GPuw2021: 0, GPuw2122: 0, GPuw2223: 0, GPuw2324: 0
    };

    // Process entities data
    usdEntities.forEach((entity, index) => {
        const row = {
            entity: entity.entity,
            totalTransactionAmount: entity.totalTransactionAmount,
            totalOrderGrossPremium: entity.totalOrderGrossPremium,
            uw1920: "-", uw2021: "-", uw2122: "-", uw2223: "-", uw2324: "-",
            GPuw1920: "-", GPuw2021: "-", GPuw2122: "-", GPuw2223: "-", GPuw2324: "-"
        };

        // Fill in underwriting years for both transaction amounts and gross premiums
        entity.uwYears.forEach(uwYear => {
            const yearKey = `uw${uwYear.year.replace('/', '')}`;
            const gpYearKey = `GPuw${uwYear.year.replace('/', '')}`;
            console.log("YEAR KEY:", yearKey);
            console.log("GP YEAR KEY:", gpYearKey);
            // Process transaction amount
            const transactionAmount = parseFloat(uwYear.transactionAmount) || 0;
            row[yearKey] = transactionAmount !== 0 ? transactionAmount.toString() : "-";
            yearlyTotals[yearKey] += transactionAmount;

            // Process order gross premium
            const grossPremium = parseFloat(uwYear.orderGrossPremium) || 0;
            row[gpYearKey] = grossPremium !== 0 ? grossPremium.toString() : "-";
            yearlyTotals[gpYearKey] += grossPremium;
        }); console.log('Row data:', row); // Log each row for debugging


        worksheet.addRow(row);
    });

    // Create the totals row
    const totalsRow = {
        entity: 'TOTAL',
        totalTransactionAmount: parseFloat(mongoDbData.premium_paid.totals.USD.transactionAmount).toFixed(2),
        totalOrderGrossPremium: parseFloat(mongoDbData.premium_paid.totals.USD.orderGrossPremium).toFixed(2),
        uw1920: yearlyTotals.uw1920 !== 0 ? parseFloat(yearlyTotals.uw1920).toFixed(2) : "-",
        uw2021: yearlyTotals.uw2021 !== 0 ? parseFloat(yearlyTotals.uw2021).toFixed(2) : "-",
        uw2122: yearlyTotals.uw2122 !== 0 ? parseFloat(yearlyTotals.uw2122).toFixed(2) : "-",
        uw2223: yearlyTotals.uw2223 !== 0 ? parseFloat(yearlyTotals.uw2223).toFixed(2) : "-",
        uw2324: yearlyTotals.uw2324 !== 0 ? parseFloat(yearlyTotals.uw2324).toFixed(2) : "-",
        GPuw1920: yearlyTotals.GPuw1920 !== 0 ? parseFloat(yearlyTotals.GPuw1920).toFixed(2) : "-",
        GPuw2021: yearlyTotals.GPuw2021 !== 0 ? parseFloat(yearlyTotals.GPuw2021).toFixed(2) : "-",
        GPuw2122: yearlyTotals.GPuw2122 !== 0 ? parseFloat(yearlyTotals.GPuw2122).toFixed(2) : "-",
        GPuw2223: yearlyTotals.GPuw2223 !== 0 ? parseFloat(yearlyTotals.GPuw2223).toFixed(2) : "-",
        GPuw2324: yearlyTotals.GPuw2324 !== 0 ? parseFloat(yearlyTotals.GPuw2324).toFixed(2) : "-"
    };
    worksheet.addRow(totalsRow);


    // console.log('Calculating claims by year');
    const claimsByYear = calculateEURClaimsByYear(mongoDbData.claims_paid);
    // console.log('Claims by year:', claimsByYear);

    const claimsRow = {
        entity: 'Claims',
        totalTransactionAmount: -mongoDbData.claims_paid.totals.USD,
        totalOrderGrossPremium: "-",
        uw1920: -claimsByYear['19/20'],
        uw2021: -claimsByYear['20/21'],
        uw2122: -claimsByYear['21/22'],
        uw2223: -claimsByYear['22/23'],
        uw2324: -claimsByYear['23/24']
    };
    // console.log('Claims row:', claimsRow);



    worksheet.addRow(claimsRow);


    // Create the grand total row
    const grandTotalRow = {
        entity: 'GRAND TOTAL',
        totalTransactionAmount: '',
        totalOrderGrossPremium: '-',
        uw1920: '-',
        uw2021: '-',
        uw2122: '-',
        uw2223: '-',
        uw2324: '-',
        GPuw1920: '-',
        GPuw2021: '-',
        GPuw2122: '-',
        GPuw2223: '-',
        GPuw2324: '-'
    };

    // Calculate grand totals
    const calculateGrandTotal = (totalValue, claimValue) => {
        if (totalValue === '-' && claimValue === '-') return '-';
        const total = parseFloat(totalValue) || 0;
        const claim = parseFloat(claimValue) || 0;
        // Since claims are already negative, we add them
        return (total + claim).toFixed(2);
    };

    // Assuming 'totalsRow' is your total row and 'claimsRow' is your claims row
    grandTotalRow.totalTransactionAmount = calculateGrandTotal(totalsRow.totalTransactionAmount, claimsRow.totalTransactionAmount);
    // grandTotalRow.totalOrderGrossPremium = calculateGrandTotal(totalsRow.totalOrderGrossPremium, claimsRow.totalOrderGrossPremium);

    // Calculate grand totals for each year
    ['uw1920', 'uw2021', 'uw2122', 'uw2223', 'uw2324'].forEach(year => {
        grandTotalRow[year] = calculateGrandTotal(totalsRow[year], claimsRow[year]);
    });

    // Add the grand total row to the worksheet
    worksheet.addRow(grandTotalRow)


    // Apply some styling
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(worksheet.rowCount).font = { bold: true };

    // Generate Excel file
    // const buffer = await workbook.xlsx.writeBuffer();

    // Generate CSV content
    const csvContent = worksheet.getSheetValues().map(row => row.join(',')).join('\n');
    const csvBuffer = await generateCsvBufferFromString(csvContent);
    const statementCsvName = 'EUR_statement_of_accounts.csv';
    const statementNameEncoded = encodeURIComponent(mongoDbData._id.toString()) + '/' + encodeURIComponent(statementCsvName);
    await addFileToApp(
        consts.MDB_CSV,
        consts.BLOB_CSV,
        mongoDbData._id.toString(),
        statementNameEncoded,
        statementCsvName,
        csvBuffer
    );
    // console.log(csvContent, "::Contnet")
    return { message: "Statement of Accounts Generated Successfully" };
}

// export const premiumSummary = async (input) => {
//     let workbook;

//     if (input instanceof ExcelJS.Workbook) {
//         workbook = input;
//     } else if (typeof input === 'string' || input instanceof Buffer) {
//         workbook = new ExcelJS.Workbook();
//         if (typeof input === 'string') {
//             await workbook.xlsx.readFile(input);
//         } else {
//             await workbook.xlsx.load(input);
//         }
//     }

//     const entityData = new Map();
//     const mongoDbData = {
//         totals: {
//             USD: { orderGrossPremium: 0, transactionAmount: 0 },
//             EUR: { orderGrossPremium: 0, transactionAmount: 0 }
//         },
//         entities: []
//     };

//     const getCellValue = (cell) => {
//         if (cell && cell.formula) {
//             console.log(`Formula found: ${cell.formula}`);
//             return cell.result;
//         }
//         return cell && cell.value;
//     };

//     workbook.eachSheet((worksheet, sheetId) => {
//         const sheetName = worksheet.name;
//         console.log(`\nProcessing sheet: ${sheetName}`);

//         entityData.set(sheetName, new Map());

//         let lastRowWithFormulas = 0;
//         let headerRowIndex = 0;
//         let currencyType = '';

//         // Find the header row (first row with data in column Q) and the last row with formulas
//         // Also search for currency type in column N
//         worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
//             const qCellValue = getCellValue(row.getCell('Q'));
//             if (qCellValue && headerRowIndex === 0) {
//                 headerRowIndex = rowNumber;
//             }
//             if (row.values.some(cell => cell && cell.formula)) {
//                 lastRowWithFormulas = rowNumber;
//             }
//             if (!currencyType) {
//                 const nCellValue = getCellValue(row.getCell('N'));
//                 if (nCellValue) {
//                     if (nCellValue.toString().includes('EUR')) {
//                         currencyType = 'EUR';
//                     } else if (nCellValue.toString().includes('USD')) {
//                         currencyType = 'USD';
//                     }
//                 }
//             }
//         });

//         console.log(`Header row: ${headerRowIndex}`);
//         console.log(`Last row with formulas: ${lastRowWithFormulas}`);
//         console.log(`Currency type: ${currencyType}`);

//         if (lastRowWithFormulas > 0 && headerRowIndex > 0) {
//             const lastRow = worksheet.getRow(lastRowWithFormulas);
//             const headerRow = worksheet.getRow(headerRowIndex);
//             const uwYearRow = worksheet.getRow(headerRowIndex + 1);

//             let orderGrossPremiumColumn = '';
//             let transactionAmountColumn = '';

//             console.log(`Header row (${headerRowIndex}):`);
//             headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
//                 console.log(`  Column ${cell.address}: ${getCellValue(cell)}`);
//                 const cellValue = getCellValue(cell);
//                 if (typeof cellValue === 'string') {
//                     if (cellValue.toLowerCase().includes('order gross premium')) {
//                         orderGrossPremiumColumn = cell.address.replace(/\d+/, '');
//                     } else if (cellValue.toLowerCase().includes('transaction amount')) {
//                         transactionAmountColumn = cell.address.replace(/\d+/, '');
//                     }
//                 }
//             });

//             console.log(`Formula row (${lastRowWithFormulas}):`);
//             lastRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
//                 console.log(`  Column ${cell.address}: ${getCellValue(cell)}`);
//             });

//             console.log(`Identified columns:`);
//             console.log(`  Order Gross Premium: ${orderGrossPremiumColumn}`);
//             console.log(`  Transaction Amount: ${transactionAmountColumn}`);

//             if (!orderGrossPremiumColumn || !transactionAmountColumn) {
//                 console.log(`Warning: Could not find Order Gross Premium or Transaction Amount columns in ${sheetName}`);
//                 return; // Skip this sheet if we can't find the required columns
//             }

//             const orderGrossPremiumCell = lastRow.getCell(orderGrossPremiumColumn);
//             const transactionAmountCell = lastRow.getCell(transactionAmountColumn);

//             const sumOrderGrossPremium = parseFloat(getCellValue(orderGrossPremiumCell)) || 0;
//             const sumTransactionAmount = parseFloat(getCellValue(transactionAmountCell)) || 0;

//             console.log(`${sheetName} (${currencyType}) - Total Order Gross Premium: ${sumOrderGrossPremium.toFixed(2)}, Total Transaction Amount: ${sumTransactionAmount.toFixed(2)}`);

//             // Update the entityData with the total values
//             if (!entityData.get(sheetName).has('Total')) {
//                 entityData.get(sheetName).set('Total', { orderGrossPremium: sumOrderGrossPremium, transactionAmount: sumTransactionAmount, currency: currencyType });
//             } else {
//                 const totalData = entityData.get(sheetName).get('Total');
//                 totalData.orderGrossPremium += sumOrderGrossPremium;
//                 totalData.transactionAmount += sumTransactionAmount;
//             }

//             // Process data for individual UW years
//             let currentUWYear = '';

//             headerRow.eachCell({ includeEmpty: false }, (headerCell, colNumber) => {
//                 const headerValue = getCellValue(headerCell);
//                 const column = headerCell.address.replace(/\d+/, '');

//                 const uwYearCell = uwYearRow.getCell(column);
//                 const uwYearValue = getCellValue(uwYearCell);

//                 if (uwYearValue && uwYearValue.toString().startsWith('U/W')) {
//                     currentUWYear = uwYearValue.toString().replace('U/W', '').trim();
//                     if (!entityData.get(sheetName).has(currentUWYear)) {
//                         entityData.get(sheetName).set(currentUWYear, { orderGrossPremium: 0, transactionAmount: 0, currency: currencyType });
//                     }
//                 }

//                 const value = parseFloat(getCellValue(lastRow.getCell(column))) || 0;

//                 if (currentUWYear) {
//                     const data = entityData.get(sheetName).get(currentUWYear);
//                     if (headerValue.toLowerCase().includes('order gross premium')) {
//                         data.orderGrossPremium = value;
//                         console.log(`${sheetName} (${currencyType}) - Order Gross Premium ${currentUWYear}: ${value.toFixed(2)}`);
//                     } else if (headerValue.toLowerCase().includes('transaction amount')) {
//                         data.transactionAmount = value;
//                         console.log(`${sheetName} (${currencyType}) - Transaction Amount ${currentUWYear}: ${value.toFixed(2)}`);
//                     }
//                 }
//             });
//         }
//         console.log('---');
//     });

//     // Generate CSV content
//     let csvContent = [];

//     // Calculate totals for USD and EUR separately
//     let totalUSDOrderGrossPremium = 0;
//     let totalUSDTransactionAmount = 0;
//     let totalEUROrderGrossPremium = 0;
//     let totalEURTransactionAmount = 0;

//     entityData.forEach((entityYearData, entityName) => {
//         entityYearData.forEach((data, uwYear) => {
//             if (data.currency === 'USD') {
//                 totalUSDOrderGrossPremium += data.orderGrossPremium;
//                 totalUSDTransactionAmount += data.transactionAmount;
//             } else if (data.currency === 'EUR') {
//                 totalEUROrderGrossPremium += data.orderGrossPremium;
//                 totalEURTransactionAmount += data.transactionAmount;
//             }
//         });
//     });

//     // Update mongoDbData totals
//     mongoDbData.totals.USD.orderGrossPremium = totalUSDOrderGrossPremium.toFixed(2);
//     mongoDbData.totals.USD.transactionAmount = totalUSDTransactionAmount.toFixed(2);
//     mongoDbData.totals.EUR.orderGrossPremium = totalEUROrderGrossPremium.toFixed(2);
//     mongoDbData.totals.EUR.transactionAmount = totalEURTransactionAmount.toFixed(2);

//     // Populate the entities array
//     entityData.forEach((entityYearData, entityName) => {
//         const entity = {
//             entity: entityName,
//             currency: '',
//             uwYears: [],
//             totalOrderGrossPremium: '',
//             totalTransactionAmount: ''
//         };

//         entityYearData.forEach((data, uwYear) => {
//             if (uwYear === 'Total') {
//                 entity.totalOrderGrossPremium = data.orderGrossPremium.toFixed(2);
//                 entity.totalTransactionAmount = data.transactionAmount.toFixed(2);
//                 entity.currency = data.currency;
//             } else {
//                 entity.currency = data.currency;
//                 entity.uwYears.push({
//                     year: uwYear,
//                     orderGrossPremium: data.orderGrossPremium !== 0 ? data.orderGrossPremium.toFixed(2) : '-',
//                     transactionAmount: data.transactionAmount !== 0 ? data.transactionAmount.toFixed(2) : '-'
//                 });
//             }
//         });

//         mongoDbData.entities.push(entity);
//     });

//     // Add totals for USD and EUR
//     csvContent.push(['TOTAL (Order Gross Premium) USD', `USD ${totalUSDOrderGrossPremium.toFixed(2)}`]);
//     csvContent.push(['TOTAL (Transaction Amount) USD', `USD ${totalUSDTransactionAmount.toFixed(2)}`]);
//     csvContent.push(['TOTAL (Order Gross Premium) EUR', `EUR ${totalEUROrderGrossPremium.toFixed(2)}`]);
//     csvContent.push(['TOTAL (Transaction Amount) EUR', `EUR ${totalEURTransactionAmount.toFixed(2)}`]);
//     csvContent.push([]);

//     // Add headers
//     csvContent.push(['Entity', 'UW Year', 'Order Gross Premium', 'Transaction Amount']);

//     // Fill data
//     entityData.forEach((entityYearData, entityName) => {
//         const sortedYears = Array.from(entityYearData.keys()).filter(year => year !== 'Total').sort();
//         sortedYears.forEach(uwYear => {
//             const data = entityYearData.get(uwYear);
//             csvContent.push([
//                 entityName,
//                 uwYear,
//                 data.orderGrossPremium !== 0 ? `${data.currency} ${data.orderGrossPremium.toFixed(2)}` : '-',
//                 data.transactionAmount !== 0 ? `${data.currency} ${data.transactionAmount.toFixed(2)}` : '-'
//             ]);
//         });
//     });

//     // Add Grand Total
//     csvContent.push(['Grand Total', '', `USD ${totalUSDOrderGrossPremium.toFixed(2)} / EUR ${totalEUROrderGrossPremium.toFixed(2)}`, `USD ${totalUSDTransactionAmount.toFixed(2)} / EUR ${totalEURTransactionAmount.toFixed(2)}`]);

//     // Convert CSV content to string
//     const csvString = csvContent.map(row => row.join(';')).join('\n');

//     // Return both the CSV string and the MongoDB-ready data
//     return { csvString, mongoDbData };
// };


async function generateCsvBufferFromString(csvString) {
    // Split the CSV string into rows
    const rows = csvString.split('\n').map(row => row.trim()).filter(row => row.length > 0);

    let processedCsvContent = '';

    // Process each row
    rows.forEach((row, rowIndex) => {
        // Parse the row, handling quoted values correctly
        const values = row.match(/(".*?"|[^",]+)(?=\s*,|\s*$)/g).map(value =>
            value.startsWith('"') && value.endsWith('"') ? value.slice(1, -1).replace(/""/g, '"') : value
        );

        // Process the values as needed
        const processedValues = values.map(value => {
            if (value === null || value === undefined) {
                return '';
            }
            value = value.toString().replace(/"/g, '""');
            return `"${value}"`;
        });

        processedCsvContent += processedValues.join(',') + '\n';
    });

    // Convert the processed CSV content to a buffer
    return Buffer.from(processedCsvContent, 'utf8');
}


async function processExcelFile(input) {
    let workbook;

    if (input instanceof ExcelJS.Workbook) {
        workbook = input;
    } else if (typeof input === 'string' || input instanceof Buffer) {
        workbook = new ExcelJS.Workbook();
        if (typeof input === 'string') {
            await workbook.xlsx.readFile(input);
        } else {
            await workbook.xlsx.load(input);
        }
    } else {
        throw new Error('Invalid input type. Expected Workbook, file path string, or Buffer.');
    }

    // Initialize arrays for CSV content and UW values
    let csvContent = [];
    const uwValues = new Set();

    // Process each worksheet to collect data
    let summaryData = [];
    workbook.eachSheet((worksheet, sheetId) => {
        if (['Claims', 'Summary USD', 'Claim Summary', 'Automation'].includes(worksheet.name)) return;
        if (sheetId > 4) return; // Skip sheets with ID greater than 4

        const sheetName = worksheet.name;
        let sumTransaction = 0;
        let sumGrossPremium = 0;
        let uwData = {};

        // worksheet.eachRow((row, rowNumber) => {
        //     const usdCell = row.getCell('A');
        //     if (usdCell.value === 'USD' && usdCell.fill && usdCell.fill.fgColor && usdCell.fill.fgColor.argb === 'FFFFFF00') {
        //         sumTransaction += row.getCell(2).value || 0;
        //         sumGrossPremium += row.getCell(3).value || 0;
        //     }

        //     const uwCell = row.getCell('R');
        //     if (uwCell.value && uwCell.value.toString().startsWith('U/W')) {
        //         const uwValue = uwCell.value.toString().replace('U/W', '').trim();
        //         uwValues.add(uwValue);
        //         if (row.getCell('O').fill && row.getCell('O').fill.fgColor && row.getCell('O').fill.fgColor.argb === 'FFFFFF00') {
        //             uwData[uwValue] = {
        //                 transaction: (uwData[uwValue]?.transaction || 0) + (row.getCell('O').value || 0),
        //                 grossPremium: (uwData[uwValue]?.grossPremium || 0) + (row.getCell('P').value || 0)
        //             };
        //         }
        //     }
        // })
        const uniqueColors = new Set();
        let yellowCellCount = 0;

        worksheet.eachRow(row => {
            row.eachCell(cell => {
                if (cell.fill && cell.fill.fgColor) {
                    uniqueColors.add(cell.fill.fgColor.argb);
                    if (cell.fill.fgColor.argb === 'FFFFFF00') {
                        yellowCellCount++;
                    }
                }
            });
        });

        console.log('Unique colors:', uniqueColors);
        console.log('Number of yellow cells:', yellowCellCount);

        worksheet.eachRow((row, rowNumber) => {
            const accountCell = row.getCell(1); // Column A
            if (accountCell.fill && accountCell.fill.fgColor && accountCell.fill.fgColor.argb === 'FFFFFF00') {
                // Check if this is a total row (contains 'Total' or 'Grand Total')
                if (accountCell.value && (accountCell.value.toString().includes('Total') || accountCell.value.toString().includes('Grand Total'))) {
                    sumTransaction += row.getCell(2).value || 0; // Column B
                    sumGrossPremium += row.getCell(3).value || 0; // Column C

                    // Collect U/W data
                    for (let colIndex = 4; colIndex <= row.cellCount; colIndex++) {
                        const cell = row.getCell(colIndex);
                        if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb === 'FFFFFF00') {
                            const cellValue = cell.value;
                            if (cellValue !== null && cellValue !== undefined) {
                                const columnHeader = worksheet.getRow(1).getCell(colIndex).value;
                                if (columnHeader.startsWith('TRANSACTION AMOUNT')) {
                                    const uwYear = columnHeader.split(' ').pop();
                                    uwValues.add(uwYear);
                                    uwData[uwYear] = {
                                        ...(uwData[uwYear] || {}),
                                        transaction: (uwData[uwYear]?.transaction || 0) + cellValue
                                    };
                                } else if (columnHeader.startsWith('ORDER GROSS PREMIUM')) {
                                    const uwYear = columnHeader.split(' ').pop();
                                    uwValues.add(uwYear);
                                    uwData[uwYear] = {
                                        ...(uwData[uwYear] || {}),
                                        grossPremium: (uwData[uwYear]?.grossPremium || 0) + cellValue
                                    };
                                }
                            }
                        }
                    }
                }
            }

        });

        //         console.log('Sum Transaction:', sumTransaction);
        // console.log('Sum Gross Premium:', sumGrossPremium);
        // console.log('UW Values:', uwValues);
        // console.log('UW Data:', uwData);
        console.log('Worksheet name:', worksheet.name);
        console.log('Row count:', worksheet.rowCount);
        console.log('Column count:', worksheet.columnCount);
        summaryData.push({ sheetName, sumTransaction, sumGrossPremium, uwData });
    });
    uwValues.add('19/20')

    // Sort UW values
    const sortedUWValues = Array.from(uwValues).sort();
    // Create headers
    csvContent.push([
        'Account',
        'Sum of Transaction Amount',
        'Sum of Order Gross Premium',
        ...sortedUWValues.map(uw => `Transaction Amount ${uw}`),
        ...sortedUWValues.map(uw => `Order Gross Premium ${uw}`)
    ]);

    // Add data rows
    summaryData.forEach(data => {
        const row = [
            data.sheetName,
            data.sumTransaction,
            data.sumGrossPremium,
            ...sortedUWValues.map(uw => data.uwData[uw]?.transaction || ''),
            ...sortedUWValues.map(uw => data.uwData[uw]?.grossPremium || '')
        ];
        csvContent.push(row);
    });

    // Add Total row
    const totalRow = ['Total'];
    for (let i = 1; i < csvContent[0].length; i++) {
        const sum = summaryData.reduce((acc, curr) => acc + (Number(csvContent[csvContent.length - 1][i]) || 0), 0);
        totalRow.push(sum);
    }
    csvContent.push(totalRow);

    // Process Claims sheet
    const claimsSheet = workbook.getWorksheet('Claims');
    if (claimsSheet) {
        let claimsTransaction = 0;
        let claimsGrossPremium = 0;
        let claimsUWData = {};

        claimsSheet.eachRow((row, rowNumber) => {
            const usdCell = row.getCell('A');
            if (usdCell.value === 'USD' && usdCell.fill && usdCell.fill.fgColor && usdCell.fill.fgColor.argb === 'FFFFFF00') {
                claimsTransaction += row.getCell(2).value || 0;
                claimsGrossPremium += row.getCell(3).value || 0;
            }

            const uwCell = row.getCell('R');
            if (uwCell.value && uwCell.value.toString().startsWith('U/W')) {
                const uwValue = uwCell.value.toString().replace('U/W', '').trim();
                if (row.getCell('O').fill && row.getCell('O').fill.fgColor && row.getCell('O').fill.fgColor.argb === 'FFFFFF00') {
                    claimsUWData[uwValue] = {
                        transaction: (claimsUWData[uwValue]?.transaction || 0) + (row.getCell('O').value || 0),
                        grossPremium: (claimsUWData[uwValue]?.grossPremium || 0) + (row.getCell('P').value || 0)
                    };
                }
            }
        });

        const claimsRow = [
            'Claims',
            claimsTransaction,
            claimsGrossPremium,
            ...sortedUWValues.map(uw => claimsUWData[uw]?.transaction || ''),
            ...sortedUWValues.map(uw => claimsUWData[uw]?.grossPremium || '')
        ];
        csvContent.push(claimsRow);
    }

    // Add Grand Total row
    const grandTotalRow = ['Grand Total'];
    for (let i = 1; i < csvContent[0].length; i++) {
        const sum = Number(csvContent[csvContent.length - 2][i]) + Number(csvContent[csvContent.length - 1][i]);
        grandTotalRow.push(sum);
    }
    csvContent.push(grandTotalRow);

    // Convert CSV content to string
    const csvString = csvContent.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');

    return csvString;
}



async function claimsSummary(input) {
    let workbook;

    if (input instanceof ExcelJS.Workbook) {
        workbook = input;
    } else if (typeof input === 'string' || input instanceof Buffer) {
        workbook = new ExcelJS.Workbook();
        if (typeof input === 'string') {
            await workbook.xlsx.readFile(input);
        } else {
            await workbook.xlsx.load(input);
        }
    }

    let totalUSD = 0;
    let totalEUR = 0;
    const entityData = {
        SQ: new Map(),
        PICC: new Map()
    };

    // Define all UW years that should be included
    const allUWYears = ['UW19/20', 'UW20/21', 'UW21/22', 'UW22/23', 'UW23/24'];

    // Initialize entityData with all UW years
    ['SQ', 'PICC'].forEach(entity => {
        allUWYears.forEach(uwYear => {
            entityData[entity].set(uwYear, { usd: 0, eur: 0 });
        });
    });
    let claims = new Map(); // Using a Map to store unique claims by claimID

    workbook.eachSheet((worksheet, sheetId) => {
        let currentEntity = '';
        let currentUWYear = '';
        let currentCurrency = '';
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const entityCell = row.getCell('A');
            const claimIDCell = row.getCell('C');
            const uwYearCell = row.getCell('R');
            const claimAmountCell = row.getCell('O');
            const imoNumberCell = row.getCell('D');
            const insuredNameCell = row.getCell('F');
            const vesselNameCell = row.getCell('G');
            const lossDescriptionCell = row.getCell('L');

            // Detect new entity
            if (entityCell.value && typeof entityCell.value === 'string') {
                if (entityCell.value.includes('SQ')) {
                    currentEntity = 'SQ';
                } else if (entityCell.value.includes('PICC')) {
                    currentEntity = 'PICC';
                }
            }

            // Detect UW Year
            if (uwYearCell.value && typeof uwYearCell.value === 'string') {
                currentUWYear = 'UW' + uwYearCell.value.trim();
            }

            // Detect currency from the header row
            if (rowNumber > 1) {
                const headerCell = worksheet.getCell(`O${rowNumber - 1}`);
                if (headerCell.value && typeof headerCell.value === 'string') {
                    if (headerCell.value.includes('USD')) {
                        currentCurrency = 'USD';
                    } else if (headerCell.value.includes('EUR')) {
                        currentCurrency = 'EUR';
                    }
                }
            }

            // Process claim amount
            if (claimAmountCell.value && currentCurrency) {
                let amount = 0;

                if (typeof claimAmountCell.value === 'number') {
                    amount = claimAmountCell.value;
                } else if (typeof claimAmountCell.value === 'string') {
                    amount = parseFloat(claimAmountCell.value.replace(/[^\d.-]/g, ''));
                }

                if (currentEntity && currentUWYear && !isNaN(amount)) {
                    const data = entityData[currentEntity].get(currentUWYear);
                    if (data) {
                        if (currentCurrency === 'USD') {
                            data.usd += amount;
                            totalUSD += amount;
                        } else if (currentCurrency === 'EUR') {
                            data.eur += amount;
                            totalEUR += amount;
                        }
                    }

                    // Check if insuredName is not null, claimAmount is a valid number, and claimID exists
                    if (insuredNameCell.value && !isNaN(amount) && amount !== 0 && claimIDCell.value) {
                        const claimID = claimIDCell.value.toString();

                        // Only add the claim if it's not already in the Map
                        if (!claims.has(claimID)) {
                            claims.set(claimID, {
                                claimID: claimID,
                                entity: currentEntity,
                                imoNumber: imoNumberCell.value,
                                insuredName: insuredNameCell.value,
                                vesselName: vesselNameCell.value,
                                lossDescription: lossDescriptionCell.value,
                                dateOfLoss: row.getCell('H').value, // Assuming date of loss is in column H
                                underwritingYear: uwYearCell.value,
                                claimAmount: amount,
                                currency: currentCurrency
                            });
                        } else {
                            console.log(`Duplicate claim ID found: ${claimID}. Skipping this entry.`);
                        }
                    }
                }
            }
        });
    });

    // Convert the Map to an array for further processing
    const uniqueClaims = Array.from(claims.values());
    console.log('Unique claims:', uniqueClaims);


    await saveClaimsToMongoDB(uniqueClaims);

    // Generate CSV content
    let csvContent = [];

    // Add totals
    csvContent.push(['TOTAL (USD)', `USD ${totalUSD.toFixed(2)}`]);
    csvContent.push(['TOTAL (EUR)', `EUR ${totalEUR.toFixed(2)}`]);
    csvContent.push([]);

    // Add headers
    csvContent.push(['Breakdown', '', 'Claim Payments', '']);
    csvContent.push(['Entity', 'UW Year', '(USD)', '(EUR)']);

    // Fill data
    ['SQ', 'PICC'].forEach(entity => {
        allUWYears.forEach(uwYear => {
            const data = entityData[entity].get(uwYear);
            csvContent.push([
                entity,
                uwYear,
                data.usd > 0 ? `USD ${data.usd.toFixed(2)}` : '-',
                data.eur > 0 ? `EUR ${data.eur.toFixed(2)}` : '-'
            ]);
        });
    });

    // Add Grand Total
    csvContent.push(['Grand Total', '', `USD ${totalUSD.toFixed(2)}`, `EUR ${totalEUR.toFixed(2)}`]);
    const csvString = csvContent.map(row => row.join(';')).join('\n');
    console.log(csvString, "::STRING");

    // Create the structured object
    const structuredObject = {
        totals: { USD: "0", EUR: "0" },
        entities: []
    };

    let currentEntity = null;

    csvContent.forEach((row, index) => {
        if (index === 0) {
            structuredObject.totals.USD = row[1].split(' ')[1];
        } else if (index === 1) {
            structuredObject.totals.EUR = row[1].split(' ')[1];
        } else if (index > 4 && row[0] !== "Grand Total") {
            if (row[0] !== currentEntity) {
                currentEntity = row[0];
                structuredObject.entities.push({
                    entity: currentEntity,
                    uwYears: []
                });
            }

            const entityIndex = structuredObject.entities.length - 1;
            structuredObject.entities[entityIndex].uwYears.push({
                year: row[1].replace('UW', ''), // Remove 'UW' prefix
                USD: row[2] !== '-' ? row[2].split(' ')[1] : null,
                EUR: row[3] !== '-' ? row[3].split(' ')[1] : null
            });
        }
    });


    return { csvString, structuredObject };

}

export const saveClaimsToMongoDB = async (claims) => {
    console.log("CLAIM::::::", claims);

    const entitiesCollection = db.collection(consts.MDB_MARINE_ENTITIES);
    const vesselsCollection = db.collection(consts.MDB_MARINE_VESSELS);
    const insuredsCollection = db.collection(consts.MDB_MARINE_INSURED);
    const claimsCollection = db.collection(consts.MDB_MARINE_CLAIMS);
    for (const claim of claims) {
        // Save or update entity
        const entityResult = await entitiesCollection.updateOne(
            { entityName: claim.entity },
            { $setOnInsert: { entityName: claim.entity } },
            { upsert: true }
        );
        const entityId = entityResult.upsertedId || (await entitiesCollection.findOne({ entityName: claim.entity }))._id;

        // Save or update vessel
        const vesselResult = await vesselsCollection.updateOne(
            { imoNumber: claim.imoNumber },
            { $setOnInsert: { imoNumber: claim.imoNumber, vesselName: claim.vesselName } },
            { upsert: true }
        );

        // Save or update insured
        const insuredResult = await insuredsCollection.updateOne(
            { insuredName: claim.insuredName },
            { $setOnInsert: { insuredName: claim.insuredName } },
            { upsert: true }
        );
        const insuredId = insuredResult.upsertedId || (await insuredsCollection.findOne({ insuredName: claim.insuredName }))._id;

        // // Save claim
        // await claimsCollection.insertOne({
        //     claimId: claim.claimID,
        //     entityId: entityId,
        //     imoNumber: claim.imoNumber,
        //     insuredId: insuredId,
        //     lossDescription: claim.lossDescription,
        //     dateOfLoss: new Date(claim.dateOfLoss),
        //     underwritingYear: claim.underwritingYear,
        //     claimAmount: claim.claimAmount,
        //     currency: claim.currency
        // });
        // Save or update claim
        await claimsCollection.updateOne(
            { claimId: claim.claimID },
            {
                $set: {
                    claimId: claim.claimID,
                    entityId: entityId,
                    imoNumber: claim.imoNumber,
                    insuredId: insuredId,
                    lossDescription: claim.lossDescription,
                    dateOfLoss: new Date(claim.dateOfLoss),
                    underwritingYear: claim.underwritingYear,
                    claimAmount: claim.claimAmount,
                    currency: claim.currency
                }
            },
            { upsert: true }
        );
    }




}

export const getCsvFiles = async () => {
    const collection = db.collection(consts.MDB_CSV);

    // // You might want to add pagination here in the future
    // const files = await collection.find({
    //     // Add any filters here, e.g., by user or organization
    // }).sort({ uploadedAt: -1 }).toArray();

    // return files.map(file => ({
    //     id: file._id,
    //     name: file.name,
    //     uploadDate: file.uploadedAt,
    //     status: file.status,
    //     rowCount: file.rowCount
    // }));

    //return all documents in the collection
    return await collection.find({}).toArray();
};

// trackVesselData
export const trackVesselData = async (body) => {
    // const collection = db.collection(consts.MDB_VESSELS);
    // const result = await collection.insertOne(body);
    // return result;
    console.log(body.imoNumber, 'body')


    const apiKey = '06423d5d-779e-4344-994e-bb06d87df1d6'
    const imoNumber = body.imoNumber
    const url = `https://api.datalastic.com/api/v0/vessel_history?api-key=${apiKey}&imo=${imoNumber}&days=1`

    // const url = `https://api.datalastic.com/api/v0/vessel_pro?api-key=${apiKey}&imo=${imoNumber}`;

    let response = await fetch(url, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json'
        },
        redirect: 'follow'
    })
    let data = await response.json()
    if (!data) {
        return []
    }
    console.log(data, 'data')
    return data
}


export const getCSVFileContent = async (id, fileName) => {
    let fileName2;
    if (!id || id === '') {
        fileName2 = fileName;
    } else {
        fileName2 = encodeURI(id) + '/' + encodeURI(fileName);
    }

    let stream = await getBlobAsStream(consts.BLOB_CSV, fileName2);

    return new Promise((resolve, reject) => {
        const chunks = [];
        stream.on('data', (chunk) => chunks.push(chunk));
        stream.on('error', reject);
        stream.on('end', () => resolve(Buffer.concat(chunks)));
    });
}
