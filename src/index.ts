Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      console.log('Excel is ready');
    }
  });
  
  export async function run() {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [["Hello World"]];
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
  