const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");
const createFileName = require("./utils/createFileName");
const {
    COLUMNS,
    PRODUCT,
    HEADERS,
    FILE,
    PRODUCTS_LIST_START_ROW,
} = require("./common/dataVars");
const {styleSheet} = require("./utils/style");
const structureSheet = require("./utils/structureSheet");
const listItemCreator = require("./utils/listItemCreator");
const addPriceLogo = require("./utils/addPriceLogo");
const getAllProducts = require("./utils/getFileFromFTP");

const priceBuilder = async () => {
    const newWorkbook = new ExcelJS.Workbook();

    fs.readFile("./source/offers.txt", "utf8", async (err, data) => {
        let newArr = [];
        let newProd = [];
        data.split("\n").forEach((i) => {
            if (i.trim()) newArr.push(i.trim());
        });
        await getAllProducts()
        const workbook = XLSX.readFile("source/All.xlsx");
        const worksheet = workbook.Sheets["Trade_CRM"];
        const jsa = XLSX.utils.sheet_to_json(worksheet);
        jsa.forEach((e) => {
            if (newArr.includes(e[PRODUCT.name])) {
                e[PRODUCT.imageUrl] = e[PRODUCT.imageUrl].split("\r\r\n").shift();
                newProd.push(e);
            }
        });

        // add file creator info
        newWorkbook.creator = HEADERS.storeName;
        newWorkbook.created = new Date(Date.now());
        let sheet = newWorkbook.addWorksheet(HEADERS.newArrivals, {
            properties: {tabColor: {argb: "FFC0000"}},
        });

        // add styles for cells, rows and columns except product list
        sheet = styleSheet(sheet);

        // get named structure of excel sheet
        const {STRUCTURE_HEADERS_LIST, STRUCTURE_LIST_TITLE, STRUCTURE_STORE} =
            structureSheet(sheet);

        // add only 'PNG' store logo to price list named as logo.png
        sheet = await addPriceLogo(newWorkbook, sheet);

        // naming of price headers
        STRUCTURE_STORE.title.value = HEADERS.storeTitle;
        STRUCTURE_STORE.phone.value = HEADERS.storePhone;
        STRUCTURE_STORE.email.value = HEADERS.storeEmail;
        STRUCTURE_STORE.url.value = {
            text: HEADERS.storeName,
            hyperlink: HEADERS.storeUrl,
            tooltip: HEADERS.storeName,
            color: {argb: "29a8ff"},
        };

        // naming price list columns and list header
        STRUCTURE_LIST_TITLE.name.value = HEADERS.newArrivals;
        STRUCTURE_HEADERS_LIST.imageRow.value = COLUMNS.image;
        STRUCTURE_HEADERS_LIST.nameRow.value = COLUMNS.name;
        STRUCTURE_HEADERS_LIST.modelRow.value = COLUMNS.model;
        STRUCTURE_HEADERS_LIST.queRow.value = COLUMNS.quantity;
        STRUCTURE_HEADERS_LIST.priceRow.value = COLUMNS.price;

        // create list items
        let index = 0;
        for (
            let startCountList = PRODUCTS_LIST_START_ROW;
            startCountList !== newProd.length + PRODUCTS_LIST_START_ROW;
            startCountList++
        ) {
            sheet = await listItemCreator(
                sheet,
                startCountList,
                newProd,
                index,
                newWorkbook
            );
            index++;
        }

        // naming output file "company name + today date dd.mm.yyyy"
        let fileName = createFileName(FILE.name);

        // save output file
        await fs.stat(`build/${fileName}.xlsx`, (err, stats) => {
            if (stats?.isFile()) {
                fileName = fileName + "_";
            }
            if (!fs.existsSync("./build")) {
                fs.mkdirSync("./build");
            }

            newWorkbook.xlsx.writeFile(`build/${fileName}.xlsx`).then(() => {
                console.log(`
        
        **********************************
        saved to ${fileName}.xlsx
        **********************************
        
        
        `);
            });
        });
    });
};

priceBuilder();
