const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");
const sharp = require("sharp");

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
      if (newArr.includes(e["Название"])) {
        e["URL Изображения"] = e["URL Изображения"].split("\r\r\n").shift();
        newProd.push(e);
      }
    });

    newWorkbook.creator = "megahertz.kiev.ua";
    newWorkbook.created = new Date(Date.now());
    const sheet = newWorkbook.addWorksheet("Новинки", {
      properties: { tabColor: { argb: "FFC0000" } },
    });

    sheet.getRow(1).height = 40;
    const editedLogo = await sharp("./common/logo.png")
      .resize(480, 50, {
        fit: "contain",
        background: { r: 255, g: 255, b: 255 },
      })
      .toBuffer();
    const logoBase64 = Buffer.from(editedLogo).toString("base64");
    const logo = newWorkbook.addImage({
      base64: logoBase64,
      extension: "png",
    });
    sheet.addImage(logo, {
      tl: { col: 1, row: 0 },
      ext: { width: 480, height: 50 },
    });

    sheet.getCell("B2").value = "Интернет магазин Megahertz";
    sheet.getCell("B2").font = {
      name: "BELL MT",
      size: 20,
    };
    sheet.getCell("B2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    sheet.getRow(2).height = 25;

    sheet.getCell("B4").value = "тел: +38(096) 803-00-33";
    sheet.getCell("B4").font = {
      name: "Cambria",
      size: 10,
      bold: true,
    };
    sheet.getCell("B4").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    sheet.getCell("B5").value = "e-mail:info@megahertz.com.ua";
    sheet.getCell("B5").font = {
      name: "Cambria",
      size: 10,
      color: { argb: "29a8ff" },
      bold: true,
    };
    sheet.getCell("B5").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    sheet.getCell("B6").value = {
      text: "megahertz.kiev.ua",
      hyperlink: "https://megahertz.kiev.ua/",
      tooltip: "www.megahertz.kiev.ua",
      color: { argb: "29a8ff" },
    };
    sheet.getCell("B6").font = {
      name: "Cambria",
      size: 10,
      bold: true,
      color: { argb: "29a8ff" },
    };
    sheet.getCell("B6").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    sheet.getColumn("A").width = 30;
    sheet.getColumn("B").width = 70;
    sheet.getColumn("C").width = 12;
    sheet.getColumn("D").width = 12;

    sheet.getColumn("E").width = 12;

    sheet.getRow(8).height = 23;
    sheet.getRow(10).height = 23;

    const cells = ["A", "B", "C", "D", "E"];
    cells.forEach((c) => {
      sheet.getCell(`${c}10`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "020000" },
      };
      sheet.getCell(`${c}9`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "878787" },
      };
    });

    const HEADER_ROWS = {
      imageRow: sheet.getCell("A8"),
      nameRow: sheet.getCell("B8"),
      modelRow: sheet.getCell("C8"),
      queRow: sheet.getCell("D8"),
      priceRow: sheet.getCell("E8"),
    };

    sheet.getCell("B10").font = {
      name: "BELL MT",
      size: 13,
      bold: true,
      color: { argb: "ffffff" },
    };
    sheet.getCell("B10").value = "Новинки";
    sheet.getCell("B10").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    HEADER_ROWS.imageRow.value = "изображение";
    HEADER_ROWS.nameRow.value = "наименование";
    HEADER_ROWS.modelRow.value = "модель";
    HEADER_ROWS.queRow.value = "кол-во";
    HEADER_ROWS.priceRow.value = "цена";

    Object.values(HEADER_ROWS).forEach((r) => {
      r.alignment = { vertical: "middle", horizontal: "center" };
      r.font = {
        name: "Arial",
        size: 8,
        bold: true,
      };
    });

    let index = 0;
    const dataCells = ["B", "C", "D", "E"];
    for (
      let startCountList = 11;
      startCountList !== newProd.length + 11;
      startCountList++
    ) {
      sheet.getRow(startCountList).height = 116;
      dataCells.forEach((c) => {
        sheet.getCell(c + startCountList).alignment = {
          vertical: "middle",
          horizontal: "center",
        };
        sheet.getCell(c + startCountList).font = {
          name: "Arial",
          size: 10,
        };
      });
      sheet.getCell(`B${startCountList}`).value = newProd[index]["Название"];
      sheet.getCell(`C${startCountList}`).value = newProd[index]["Код товара"];
      sheet.getCell(`D${startCountList}`).value = "В наличии";
      sheet.getCell(`E${startCountList}`).value =
        newProd[index]["Цена Розничная"];

      const url = newProd[index]["URL Изображения"];
      const response = await fetch(url);
      const blob = await response.blob();
      const arrayBuffer = await blob.arrayBuffer();
      const data = await sharp(Buffer.from(arrayBuffer))
        .resize(151, 151, {
          fit: "contain",
          background: { r: 255, g: 255, b: 255 },
        })
        .toBuffer();
      const imageBase64 = Buffer.from(data).toString("base64");

      const prodImage = newWorkbook.addImage({
        base64: imageBase64,
        extension: "jpeg",
      });

      sheet.addImage(prodImage, {
        tl: { col: 0.5, row: parseFloat(`${startCountList - 1}.1`) },
        ext: { height: 151, width: 151 },
      });
      index++;
    }

    const today = new Date();
    const yyyy = today.getFullYear();
    let mm = today.getMonth() + 1; // Months start at 0!
    let dd = today.getDate();

    if (dd < 10) dd = "0" + dd;
    if (mm < 10) mm = "0" + mm;

    let formattedToday = `MHz_new_${dd}.${mm}.${yyyy}`;
    await fs.stat(`build/${formattedToday}.xlsx`, (err, stats) => {
      if (stats?.isFile()) {
        formattedToday = formattedToday + "_";
      }
      newWorkbook.xlsx.writeFile(`build/${formattedToday}.xlsx`).then(() => {
        console.clear();
        console.log(`
        
        **********************************
        saved to ${formattedToday}.xlsx
        **********************************
        
        
        `);
      });
    });
  });
};

priceBuilder();
