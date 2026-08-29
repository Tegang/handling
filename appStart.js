// =========================================================================
// APPSTART.GS - Blixtsnabb startsekvens
// =========================================================================

var APP_VERSION = "1.2"; 
var APP_BUILD_TIME = "2026-08-29 14:18"; 

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  template.appVersion = "v" + APP_VERSION + " • " + APP_BUILD_TIME;

  try {
    template.scriptUrl = ScriptApp.getService().getUrl();
  } catch(err) {
    template.scriptUrl = "";
  }

  return template.evaluate()
      .setTitle('Handling')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDriveLastUpdated() {
  return "v" + APP_VERSION + " • " + APP_BUILD_TIME;
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "<!-- Filen " + filename + " hittades inte -->";
  }
}