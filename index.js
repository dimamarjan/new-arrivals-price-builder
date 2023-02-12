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
const { styleSheet } = require("./utils/style");
const structureSheet = require("./utils/structureSheet");
const listItemCreator = require("./utils/listItemCreator");
const addPriceLogo = require("./utils/addPriceLogo");

const priceBuilder = async () => {
  const newWorkbook = new ExcelJS.Workbook();

  fs.readFile("./source/offers.txt", "utf8", async (err, data) => {
    let newArr = [];
    let newProd = [];
    data.split("\n").forEach((i) => {
      if (i.trim()) newArr.push(i.trim());
    });
    const workbook = XLSX.readFile("source/All.xlsx");
    const worksheet = workbook.Sheets["Trade_CRM"];
    const jsa = XLSX.utils.sheet_to_json(worksheet);
    jsa.forEach((e) => {
      if (newArr.includes(e[PRODUCT.name])) {
        e[PRODUCT.imageUrl] = e[PRODUCT.imageUrl].split("\r\r\n").shift();
        newProd.push(e);
      }
    });

    newWorkbook.creator = HEADERS.storeName;
    newWorkbook.created = new Date(Date.now());
    let sheet = newWorkbook.addWorksheet(HEADERS.newArrivals, {
      properties: { tabColor: { argb: "FFC0000" } },
    });
    sheet = styleSheet(sheet);
    const { STRUCTURE_HEADERS_LIST, STRUCTURE_LIST_TITLE, STRUCTURE_STORE } =
      structureSheet(sheet);

    sheet = await addPriceLogo(newWorkbook, sheet);

    STRUCTURE_STORE.title.value = HEADERS.storeTitle;
    STRUCTURE_STORE.phone.value = HEADERS.storePhone;
    STRUCTURE_STORE.email.value = HEADERS.storeEmail;
    STRUCTURE_STORE.url.value = {
      text: HEADERS.storeName,
      hyperlink: HEADERS.storeUrl,
      tooltip: HEADERS.storeName,
      color: { argb: "29a8ff" },
    };

    STRUCTURE_LIST_TITLE.name.value = HEADERS.newArrivals;
    STRUCTURE_HEADERS_LIST.imageRow.value = COLUMNS.image;
    STRUCTURE_HEADERS_LIST.nameRow.value = COLUMNS.name;
    STRUCTURE_HEADERS_LIST.modelRow.value = COLUMNS.model;
    STRUCTURE_HEADERS_LIST.queRow.value = COLUMNS.quantity;
    STRUCTURE_HEADERS_LIST.priceRow.value = COLUMNS.price;

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

    let fileName = createFileName(FILE.name);

    await fs.stat(`build/${fileName}.xlsx`, (err, stats) => {
      if (stats?.isFile()) {
        fileName = fileName + "_";
      }
      if (!fs.existsSync("./build")) {
        fs.mkdirSync("./build");
      }

      newWorkbook.xlsx.writeFile(`build/${fileName}.xlsx`).then(() => {
        console.clear();
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
