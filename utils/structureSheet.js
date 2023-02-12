function structureSheet(sheet) {
  const STRUCTURE_HEADERS_LIST = {
    imageRow: sheet.getCell("A8"),
    nameRow: sheet.getCell("B8"),
    modelRow: sheet.getCell("C8"),
    queRow: sheet.getCell("D8"),
    priceRow: sheet.getCell("E8"),
  };

  const STRUCTURE_LIST_TITLE = {
    name: sheet.getCell("B10"),
  };

  const STRUCTURE_STORE = {
    title: sheet.getCell("B2"),
    phone: sheet.getCell("B4"),
    email: sheet.getCell("B5"),
    url: sheet.getCell("B6"),
  };

  return { STRUCTURE_HEADERS_LIST, STRUCTURE_LIST_TITLE, STRUCTURE_STORE };
}

module.exports = structureSheet;
