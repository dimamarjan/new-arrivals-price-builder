const { PRODUCT } = require("../common/dataVars");
const sharp = require("sharp");
const { textAlignmentCenter, fontItemList } = require("./style");

async function listItemCreator(sheet, cellNumber, item, index, book) {
  const dataCells = ["B", "C", "D", "E"];
  dataCells.forEach((c) => {
    sheet.getCell(c + cellNumber).alignment = textAlignmentCenter;
    sheet.getCell(c + cellNumber).font = fontItemList;
  });

  sheet.getRow(cellNumber).height = 116;
  sheet.getCell(`${dataCells[0]}${cellNumber}`).value =
    item[index][PRODUCT.name];
  sheet.getCell(`${dataCells[1]}${cellNumber}`).value =
    item[index][PRODUCT.code];
  sheet.getCell(`${dataCells[2]}${cellNumber}`).value = PRODUCT.quantity;
  sheet.getCell(`${dataCells[3]}${cellNumber}`).value =
    item[index][PRODUCT.price];

  const url = item[index][PRODUCT.imageUrl];
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

  const prodImage = book.addImage({
    base64: imageBase64,
    extension: "jpeg",
  });

  sheet.addImage(prodImage, {
    tl: { col: 0.5, row: parseFloat(`${cellNumber - 1}.1`) },
    ext: { height: 151, width: 151 },
  });

  return sheet;
}

module.exports = listItemCreator;
