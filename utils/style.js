const structureSheet = require("./structureSheet");

const descriptionFontBold = {
  name: "Cambria",
  size: 10,
  bold: true,
};

const fontItemList = {
  name: "Arial",
  size: 10,
};

const textAlignmentCenter = {
  vertical: "middle",
  horizontal: "center",
};

const fillPattern = {
  type: "pattern",
  pattern: "solid",
};

function styleSheet(sheet) {
  const { STRUCTURE_HEADERS_LIST } = structureSheet(sheet);

  sheet.getColumn("A").width = 30;
  sheet.getColumn("B").width = 70;
  sheet.getColumn("C").width = 12;
  sheet.getColumn("D").width = 12;
  sheet.getColumn("E").width = 12;

  sheet.getRow(1).height = 40;
  sheet.getRow(2).height = 25;
  sheet.getRow(8).height = 23;
  sheet.getRow(10).height = 23;

  sheet.getCell("B2").alignment = textAlignmentCenter;
  sheet.getCell("B4").alignment = textAlignmentCenter;
  sheet.getCell("B5").alignment = textAlignmentCenter;
  sheet.getCell("B6").alignment = textAlignmentCenter;
  sheet.getCell("B10").alignment = textAlignmentCenter;

  sheet.getCell("B4").font = descriptionFontBold;
  sheet.getCell("B5").font = {
    ...descriptionFontBold,
    color: { argb: "29a8ff" },
  };
  sheet.getCell("B6").font = {
    ...descriptionFontBold,
    color: { argb: "29a8ff" },
  };
  sheet.getCell("B10").font = {
    name: "BELL MT",
    size: 13,
    bold: true,
    color: { argb: "ffffff" },
  };
  sheet.getCell("B2").font = {
    name: "BELL MT",
    size: 20,
  };

  const cells = ["A", "B", "C", "D", "E"];
  cells.forEach((c) => {
    sheet.getCell(`${c}10`).fill = {
      ...fillPattern,
      fgColor: { argb: "020000" },
    };
    sheet.getCell(`${c}9`).fill = {
      ...fillPattern,
      fgColor: { argb: "878787" },
    };
  });

  Object.values(STRUCTURE_HEADERS_LIST).forEach((r) => {
    r.alignment = textAlignmentCenter;
    r.font = {
      name: "Arial",
      size: 8,
      bold: true,
    };
  });

  return sheet;
}

module.exports = { styleSheet, textAlignmentCenter, fontItemList };
