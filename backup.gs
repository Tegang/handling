// =========================================================================
// BACKUP.GS - Källkodsexport till Google Drive
// =========================================================================
 
var MIN_KOMMENTAR = "innan GitHub repository";

function exporteraKodTillDrive() {
  var exaktTidsStämpel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var filNamn = exaktTidsStämpel + ".txt";
  var textFilInnehall = hamtaAllKällkod();

  var kommentar = MIN_KOMMENTAR.trim();
  if (kommentar === "") {
    kommentar = "Manuell backup";
  }

  var filInnehallMedKommentar = 
    "========================================\n" +
    "BACKUP DATUM: " + exaktTidsStämpel + "\n" +
    "KOMMENTAR: " + kommentar + "\n" +
    "========================================\n\n" +
    textFilInnehall;

  var f1 = DriveApp.getFoldersByName("AppsScript");
  var folderAppsScript = f1.hasNext() ? f1.next() : DriveApp.createFolder("AppsScript");

  var f2 = folderAppsScript.getFoldersByName("Handling");
  var folderHandling = f2.hasNext() ? f2.next() : folderAppsScript.createFolder("Handling");

  var f3 = folderHandling.getFoldersByName("Backup");
  var targetFolder = f3.hasNext() ? f3.next() : folderHandling.createFolder("Backup");

  targetFolder.createFile(filNamn, filInnehallMedKommentar);

  if (kommentar !== "") {
    Logger.log("💬 Kommentar: " + kommentar);
  }
  Logger.log("✅ Backup sparad i: AppsScript > Handling > Backup > " + filNamn);
}

function hamtaAllKällkod() {
  try {
    var scriptId = ScriptApp.getScriptId();
    var url = "https://script.google.com/feeds/download/export?id=" + scriptId + "&format=json";
    
    var options = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    var samladKod = "";

    if (json.files && json.files.length > 0) {
      json.files.forEach(function(file) {
        samladKod += "--------------------------------------------------\n";
        samladKod += "// FIL: " + file.name + "\n";
        samladKod += "--------------------------------------------------\n\n";
        samladKod += file.source + "\n\n\n";
      });
      return samladKod;
    }
  } catch (e) {
    Logger.log("Kunde inte hämta källkod: " + e.toString());
  }
  return "// Kunde inte läsa källkoden automatiskt.";
}