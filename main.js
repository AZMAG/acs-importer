const XLSX = require("xlsx");

const { get5yrDataByTableId, get5yrGeoTable, writeToSql } = require("./data");

const { year, excelFile, outputDatabase, chunkSize } = require("./config");

const workbook = XLSX.readFile(excelFile);
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

(async () => {
  // Import GEO Tables
  console.log("Loading 5yr Geo Tables");
  const geos5Yr = await get5yrGeoTable();
  const tblName5Yr = `${outputDatabase}.dbo.G${year}5YR`;
  await writeToSql(geos5Yr, {
    tblName: tblName5Yr,
    chunkSize,
    idColumn: "DADSID",
    overwrite: true,
  });

  for (let i = 0; i < xlData.length; i++) {
    const row = xlData[i];
    const tableId = row["Table ID"];
    console.log(`Loading table file (${i} / ${xlData.length}) -- ${tableId}`);
    if (tableId) {
      const fiveYrData = await get5yrDataByTableId(tableId);
      console.log(`Table Loaded. Writing to db... (${fiveYrData.length} rows)`);
      const tblName = `${outputDatabase}.dbo.${tableId}`;
      await writeToSql(fiveYrData, {
        tblName,
        chunkSize,
        idColumn: "GEO_ID",
        overwrite: true,
      });
      console.log(`${tableId} - written to database.`);
    }
  }
  // await createJoinedViewByYear(xlData, '5');
})();
