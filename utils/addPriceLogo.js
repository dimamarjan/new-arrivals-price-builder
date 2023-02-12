const sharp = require("sharp");

async function addPriceLogo(newWorkbook, sheet) {
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
  return sheet;
}

module.exports = addPriceLogo;
