const SPREADSHEET_ID = "1iDj5bdilqofk663MOJRwIQrwtGcuqRGx2Yyb46RSKXs";

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Manpower & Workload Planning')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Utility to include separate HTML/CSS/JS files into Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
