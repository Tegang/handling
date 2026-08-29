// =========================================================================
// SERVERHANDLING.GS - Serverkod för listor, butiker, AI, bilder och databas
// =========================================================================
 
function capitalizeFirstLetter(str) {
  str = String(str || "").trim();
  if (!str) return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function normalizeAvdName(str) {
  return String(str || "")
    .toLowerCase()
    .replace(/&/g, "och")
    .replace(/[^a-z0-9åäö]/gi, "")
    .trim();
}

function normalizeItemName(str) {
  str = String(str || "").toLowerCase().trim().replace(/[^a-z0-9åäö]/gi, "");
  if (str.length > 4) {
    str = str.replace(/(er|ar|or|en)$/i, "");
  }
  return str;
}

// FAS 1: LÄS ENDAST DET ABSOLUT NÖDVÄNDIGASTE FÖR SKÄRM 1 (BLIXTSNABB START)
function getFastInitialAppData(butikId, listaId) {
  const lists = getLists();
  
  if (!listaId && lists.length > 0) {
    listaId = lists[0].listaId;
  }

  const stores = getStores();
  
  let items = [];
  if (listaId && lists.some(l => l.listaId === listaId)) {
    items = getShoppingList(butikId, listaId);
  } else {
    listaId = "";
  }

  return {
    lists: lists,
    stores: stores,
    items: items,
    activeListId: listaId
  };
}

// FAS 2: HÄMTA VARUDATABAS, AVDELNINGAR & ENHETER I BAKGRUNDEN
function getDeferredAppData() {
  return {
    departments: getDepartmentsServer(),
    units: getUnitsServer(),
    masterVaror: getMasterVarorServer()
  };
}

function getMasterVarorServer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Vara");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  data.shift();

  const map = {};
  data.forEach(r => {
    const vId = String(r[0]);
    const vNamn = String(r[1] || "").trim();
    const nameKey = vNamn.toLowerCase();
    const avdId = String(r[2] || "0");
    const enhetId = r[3] ? String(r[3]) : "";

    if (nameKey) {
      map[nameKey] = { 
        varaId: vId, 
        varaNamn: capitalizeFirstLetter(vNamn), 
        avdId: avdId, 
        enhetId: enhetId 
      };
    }
  });
  return map;
}

function getStores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Butik");
  let butikAvdSheet = ss.getSheetByName("Butik_Avdelning");

  if (!sheet) return [];

  const bData = sheet.getDataRange().getValues();
  const baData = butikAvdSheet ? butikAvdSheet.getDataRange().getValues() : [];

  if (bData.length <= 1) return [];

  bData.shift();
  if (baData.length > 0) baData.shift();

  const storeSortMaps = new Map();
  baData.forEach(r => {
    const bId = String(r[0]);
    const aId = String(r[1]);
    const rawSort = r[2];
    const sortNum = (rawSort !== "" && !isNaN(rawSort)) ? Number(rawSort) : 999;

    if (!storeSortMaps.has(bId)) {
      storeSortMaps.set(bId, {});
    }
    storeSortMaps.get(bId)[aId] = sortNum;
  });

  return bData.map(r => {
    const bId = String(r[0]);
    const bNamn = String(r[1]).trim();
    const bPlats = r[2] ? String(r[2]).trim() : "";
    const fulltNamn = bPlats ? (bNamn + " • " + bPlats) : bNamn;
    const avdSort = storeSortMaps.get(bId) || {};

    return { 
      butikId: bId, 
      butikNamn: fulltNamn, 
      rawNamn: bNamn, 
      butikPlats: bPlats, 
      avdSort: avdSort 
    };
  });
}

function getDepartmentsServer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Avdelning");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift();

  return data.map(r => ({
    avdId: String(r[0]),
    avdNamn: String(r[1]).trim()
  }));
}

function getUnitsServer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Enhet");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift();

  return data.map(r => ({
    enhetId: String(r[0]),
    enhetsnamn: String(r[1]).trim()
  }));
}

function getLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Lista");
  
  if (!sheet) {
    sheet = ss.insertSheet("Lista");
    sheet.appendRow(["Lista_ID", "Lista_Namn", "Lista_Klar", "Skapad_Datum"]);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  data.shift(); 
  return data.map(r => ({
    listaId: String(r[0]),
    listaNamn: String(r[1]),
    listaKlar: r[2] === true || r[2] === "TRUE"
  }));
}

function createNewList(listaNamn) {
  listaNamn = capitalizeFirstLetter(listaNamn);
  if (!listaNamn) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Lista");
  if (!sheet) {
    sheet = ss.insertSheet("Lista");
    sheet.appendRow(["Lista_ID", "Lista_Namn", "Lista_Klar", "Skapad_Datum"]);
  }

  const listaId = String(new Date().getTime());
  sheet.appendRow([listaId, listaNamn, false, new Date()]);
  return listaId;
}

function deleteDriveFileById(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(String(fileId)).setTrashed(true);
  } catch (e) {
    // Om filen inte hittas eller redan raderats
  }
}

function deleteList(listaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listaSheet = ss.getSheetByName("Lista");
  const handlaSheet = ss.getSheetByName("Handla");

  if (listaSheet) {
    const listaData = listaSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < listaData.length; i++) {
      if (String(listaData[i][0]).trim() === String(listaId).trim()) {
        listaSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  if (handlaSheet) {
    const handlaData = handlaSheet.getDataRange().getDisplayValues();
    for (let i = handlaData.length - 1; i >= 1; i--) {
      if (String(handlaData[i][1]).trim() === String(listaId).trim()) {
        const bildId = handlaData[i][6];
        deleteDriveFileById(bildId);
        handlaSheet.deleteRow(i + 1);
      }
    }
  }
}

function predictAvdelningForText(itemName) {
  if (!itemName) return "";
  const normName = itemName.toLowerCase().trim();

  const quickKeywords = [
    { keywords: ["korv", "falukorv", "kött", "skinka", "bacon", "kyckling", "färs", "köttfärs", "salami", "leverpastej", "fläsk", "nötkött", "biff"], avd: "Kött & Chark" },
    { keywords: ["mjölk", "ost", "smör", "grädde", "fil", "yoghurt", "kvarg", "fraiche", "margarin"], avd: "Mejeri" },
    { keywords: ["äpple", "banan", "tomat", "gurka", "sallad", "potatis", "lök", "citron", "morot", "frukt", "bär", "vindruvor", "avokado", "paprika"], avd: "Frukt & Grönt" },
    { keywords: ["bröd", "limpa", "franska", "korvbröd", "hamburgerbröd", "kaka", "bulle"], avd: "Bröd" },
    { keywords: ["pasta", "nudlar", "nudel", "ris", "mjöl", "socker", "konserv", "olja", "krydda", "salt", "peppar", "buljong", "sås", "soppa", "taco", "müsli", "flingor", "havregryn", "krossade tomater"], avd: "Skafferi" },
    { keywords: ["fisk", "lax", "torsk", "räkor", "sill", "krabba"], avd: "Fisk & Skaldjur" },
    { keywords: ["läsk", "juice", "saft", "vatten", "öl", "cider", "energidryck"], avd: "Dryck" },
    { keywords: ["tvål", "shampoo", "schampo", "toalettpapper", "toapapper", "diskmedel", "tvättmedel", "tandkräm"], avd: "Hygien & Städa" }
  ];

  for (let group of quickKeywords) {
    for (let kw of group.keywords) {
      if (normName.includes(kw)) {
        return group.avd;
      }
    }
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avdSheet = ss.getSheetByName("Avdelning");
  if (!avdSheet) return "";
  const avdData = avdSheet.getDataRange().getValues();
  avdData.shift();
  const avdNamesList = avdData.map(r => String(r[1]).trim()).filter(n => n).join(", ");

  const validModels = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-2.5-flash"];

  const prompt = `Vilken avdelning passar bäst för matvaran "${itemName}"?
TILLGÄNGLIGA AVDELNINGAR: [${avdNamesList}]
Svara ENDAST med avdelningsnamnet exakt från listan. Om du är osäker svara "Övrigt".`;

  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": 0.1 }
  };

  for (let model of validModels) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    try {
      const response = UrlFetchApp.fetch(url, {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });

      if (response.getResponseCode() === 200) {
        const result = JSON.parse(response.getContentText());
        return result.candidates[0].content.parts[0].text.trim();
      }
    } catch (e) {
      // Försök nästa modell
    }
  }

  return "";
}

function getShoppingList(butikId, listaId) {
  butikId = String(butikId || "1");
  listaId = String(listaId || "");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handlaSheet = ss.getSheetByName("Handla");
  const varaSheet = ss.getSheetByName("Vara");
  const avdSheet = ss.getSheetByName("Avdelning");
  const enhetSheet = ss.getSheetByName("Enhet");
  const butikAvdSheet = ss.getSheetByName("Butik_Avdelning");

  if (!handlaSheet || !varaSheet || !avdSheet) return [];

  const handlaData = handlaSheet.getDataRange().getDisplayValues();
  const varaData = varaSheet.getDataRange().getValues();
  const avdData = avdSheet.getDataRange().getValues();
  const enhetData = enhetSheet ? enhetSheet.getDataRange().getValues() : [];
  const butikAvdData = butikAvdSheet ? butikAvdSheet.getDataRange().getValues() : [];

  if (handlaData.length <= 1) return [];

  handlaData.shift(); 
  if (varaData.length > 0) varaData.shift();
  if (avdData.length > 0) avdData.shift();
  if (enhetData.length > 0) enhetData.shift();
  if (butikAvdData.length > 0) butikAvdData.shift();

  const avdNamesMap = new Map();
  avdData.forEach(r => avdNamesMap.set(String(r[0]), String(r[1]).trim()));

  const enhetNamesMap = new Map();
  enhetData.forEach(r => enhetNamesMap.set(String(r[0]), String(r[1]).trim()));

  const avdSortMap = new Map();
  butikAvdData.forEach(r => {
    const bId = String(r[0]);
    const aId = String(r[1]);
    const rawSort = r[2];
    const sortNum = (rawSort !== "" && !isNaN(rawSort)) ? Number(rawSort) : 999;

    if (bId === butikId) {
      avdSortMap.set(aId, sortNum);
    }
  });

  const varaMap = new Map();
  varaData.forEach(r => varaMap.set(String(r[0]), { 
    namn: capitalizeFirstLetter(r[1]), 
    avdId: String(r[2]),
    enhetId: r[3] ? String(r[3]) : ""
  }));

  let list = [];
  handlaData.forEach(r => {
    const handlaId = String(r[0]).trim();
    const itemListaId = String(r[1]).trim();
    const varaId = String(r[2]).trim();
    const antal = (r[3] !== undefined && r[3] !== null) ? String(r[3]).trim() : ""; 
    const info = r[4] || "";                          
    const klar = String(r[5]).toUpperCase() === "TRUE";    
    const bildId = r[6] ? String(r[6]).trim() : "";
    const bildUrl = bildId ? ("https://lh3.googleusercontent.com/d/" + bildId) : "";

    if (itemListaId === listaId) {
      const vara = varaMap.get(varaId) || { namn: "Okänd vara (" + varaId + ")", avdId: "0", enhetId: "" };
      const avdNamn = avdNamesMap.get(vara.avdId) || "Övrigt";
      const enhetNamn = vara.enhetId ? (enhetNamesMap.get(vara.enhetId) || vara.enhetId) : "";

      let sortering = avdSortMap.has(vara.avdId) ? avdSortMap.get(vara.avdId) : (vara.avdId === "0" ? 0 : 999);

      list.push({
        handlaId: handlaId,
        listaId: itemListaId,
        varaId: varaId,
        varaNamn: vara.namn,
        avdId: vara.avdId,
        avdelningNamn: avdNamn,
        enhetId: vara.enhetId,
        enhetNamn: enhetNamn,
        bildUrl: bildUrl,
        antal: antal, 
        sortering: sortering,
        klar: klar,
        info: info
      });
    }
  });

  list.sort((a, b) => {
    if (a.klar !== b.klar) return a.klar ? 1 : -1;
    return a.sortering - b.sortering;
  });

  return list;
}

function toggleItemDone(handlaId, isDone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Handla");
  if (!sheet) return;
  const data = sheet.getDataRange().getDisplayValues();
  const searchId = String(handlaId).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      sheet.getRange(i + 1, 6).setValue(isDone);
      break;
    }
  }
}

function deleteSingleItem(handlaId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Handla");
  if (!sheet) return;
  const data = sheet.getDataRange().getDisplayValues();
  const searchId = String(handlaId).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      const bildId = data[i][6];
      deleteDriveFileById(bildId);
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function addSingleItem(itemName, itemInfo, listaId) {
  itemName = capitalizeFirstLetter(itemName);
  itemInfo = itemInfo ? String(itemInfo).trim() : "";
  if (!itemName) return null;

  const createdItems = addMultipleItems([{ vara: itemName, info: itemInfo }], listaId);
  return (createdItems && createdItems.length > 0) ? createdItems[0] : null;
}

function findAvdelningFromExistingVara(cleanName, varaData) {
  if (!cleanName || !varaData) return null;
  const normNew = normalizeItemName(cleanName);

  for (let r of varaData) {
    const existingName = String(r[1] || "").trim();
    const existingAvdId = r[2];

    if (existingName && existingAvdId && String(existingAvdId) !== "0") {
      if (normalizeItemName(existingName) === normNew) {
        return String(existingAvdId);
      }
    }
  }
  return null;
}

function addMultipleItems(itemsArray, listaId) {
  if (!itemsArray || itemsArray.length === 0) return [];
  listaId = String(listaId || "");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let varaSheet = ss.getSheetByName("Vara");
  let handlaSheet = ss.getSheetByName("Handla");
  let avdSheet = ss.getSheetByName("Avdelning");
  let enhetSheet = ss.getSheetByName("Enhet");

  if (!varaSheet) {
    varaSheet = ss.insertSheet("Vara");
    varaSheet.appendRow(["Vara_ID", "Vara_Namn", "Avdelning_ID", "Enhet_ID"]);
  }

  if (!handlaSheet) {
    handlaSheet = ss.insertSheet("Handla");
    handlaSheet.appendRow(["Handla_ID", "Lista_ID", "Vara_ID", "Antal", "Handla_Info", "Handla_Klar", "Bild_ID"]);
  }

  const varaData = varaSheet.getDataRange().getValues();
  if (varaData.length > 0) varaData.shift();

  const handlaData = handlaSheet.getDataRange().getValues();
  if (handlaData.length > 0) handlaData.shift();

  const avdData = avdSheet ? avdSheet.getDataRange().getValues() : [];
  if (avdData.length > 0) avdData.shift();

  const enhetData = enhetSheet ? enhetSheet.getDataRange().getValues() : [];
  if (enhetData.length > 0) enhetData.shift();

  const avdNameToIdMap = new Map();
  const avdIdToNameMap = new Map();
  avdData.forEach(r => {
    const avdId = String(r[0]);
    const avdName = String(r[1] || "").trim();
    const normKey = normalizeAvdName(avdName);
    if (normKey) avdNameToIdMap.set(normKey, avdId);
    avdIdToNameMap.set(avdId, avdName);
  });

  const enhetIdToNameMap = new Map();
  enhetData.forEach(r => {
    enhetIdToNameMap.set(String(r[0]), String(r[1] || "").trim());
  });

  const varaMap = new Map();
  const varaNormMap = new Map();
  varaData.forEach(r => {
    const vId = String(r[0]);
    const nameKey = String(r[1] || "").trim().toLowerCase();
    const normKey = normalizeItemName(r[1]);
    if (nameKey) varaMap.set(nameKey, vId);
    if (normKey) varaNormMap.set(normKey, vId);
  });

  const activeVaraIds = new Set();
  handlaData.forEach(r => {
    const itemListaId = String(r[1]);
    const varaId = String(r[2]);
    const klar = r[5] === true || r[5] === "TRUE";
    if (itemListaId === listaId && !klar) activeVaraIds.add(varaId);
  });

  const newVaraRows = [];
  const newHandlaRows = [];
  const createdItemsList = [];
  const now = new Date().getTime();

  itemsArray.forEach((itemObj, index) => {
    let rawName = (typeof itemObj === 'object' && itemObj.vara) ? itemObj.vara : (typeof itemObj === 'string' ? itemObj : "");
    let rawAvd = (typeof itemObj === 'object' && (itemObj.avdelning || itemObj.kategori)) ? (itemObj.avdelning || itemObj.kategori) : "";
    const itemInfo = (typeof itemObj === 'object' && itemObj.info) ? String(itemObj.info).trim() : "";

    let cleanName = capitalizeFirstLetter(String(rawName).trim());
    let suggestedAvd = String(rawAvd).trim();

    if (!cleanName) return;

    const normName = normalizeItemName(cleanName);
    let varaId = varaMap.get(cleanName.toLowerCase()) || varaNormMap.get(normName);
    let assignedAvdId = "0";
    let assignedEnhetId = "";

    if (!varaId) {
      varaId = String(now + index);
      varaMap.set(cleanName.toLowerCase(), varaId);
      varaNormMap.set(normName, varaId);

      const inheritedAvdId = findAvdelningFromExistingVara(cleanName, varaData);
      if (inheritedAvdId) {
        assignedAvdId = String(inheritedAvdId);
      } else {
        if (!suggestedAvd) {
          suggestedAvd = predictAvdelningForText(cleanName);
        }

        const normSuggested = normalizeAvdName(suggestedAvd);
        if (avdNameToIdMap.has(normSuggested)) {
          assignedAvdId = String(avdNameToIdMap.get(normSuggested));
        } else {
          for (let [normKey, id] of avdNameToIdMap.entries()) {
            if (normKey.length > 2 && (normKey.includes(normSuggested) || normSuggested.includes(normKey))) {
              assignedAvdId = String(id);
              break;
            }
          }
        }
      }

      newVaraRows.push([varaId, cleanName, assignedAvdId, ""]);
    } else {
      const existingVara = varaData.find(r => String(r[0]) === String(varaId));
      if (existingVara) {
        assignedAvdId = String(existingVara[2] || "0");
        assignedEnhetId = existingVara[3] ? String(existingVara[3]) : "";
      }
    }

    if (!activeVaraIds.has(varaId)) {
      activeVaraIds.add(varaId);
      const handlaId = String(now + index) + String(Math.floor(Math.random() * 90) + 10);
      
      newHandlaRows.push([handlaId, listaId, varaId, "", itemInfo, false, ""]);

      createdItemsList.push({
        handlaId: handlaId,
        listaId: listaId,
        varaId: varaId,
        varaNamn: cleanName,
        avdId: assignedAvdId,
        avdelningNamn: avdIdToNameMap.get(assignedAvdId) || "Övrigt",
        enhetId: assignedEnhetId,
        enhetNamn: enhetIdToNameMap.get(assignedEnhetId) || "",
        bildUrl: "",
        antal: "",
        info: itemInfo
      });
    }
  });

  if (newVaraRows.length > 0) {
    varaSheet.getRange(varaSheet.getLastRow() + 1, 1, newVaraRows.length, 4).setValues(newVaraRows);
  }

  if (newHandlaRows.length > 0) {
    handlaSheet.getRange(handlaSheet.getLastRow() + 1, 1, newHandlaRows.length, 7).setValues(newHandlaRows);
  }

  return createdItemsList;
}

function processWhiteboardImage(base64Image, listaId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API-nyckel saknas i Skriptegenskaper!");

  const validModels = ["gemini-3.5-flash-lite", "gemini-flash-latest", "gemini-2.5-flash"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avdSheet = ss.getSheetByName("Avdelning");
  if (!avdSheet) throw new Error("Avdelnings-fliken saknas!");
  const avdData = avdSheet.getDataRange().getValues();
  avdData.shift();
  const avdNamesList = avdData.map(r => String(r[1]).trim()).filter(n => n).join(", ");

  const prompt = `Identifiera matvarorna på bilden av den handskrivna whiteboard-listan.

TILLGÄNGLIGA AVDELNINGAR I BUTIKEN:
[${avdNamesList}]

STRIKTA REGLER:
1. GRUNDFORM & SINGULARIS: Skriv ALLA varor i sin GRUNDFORM i SINGULARIS på svenska med stor begynnelsebokstav (t.ex. skriv "Äpple" ISTÄLLET FÖR "Äpplen", "Banan" ISTÄLLET FÖR "Bananer", "Tomat" ISTÄLLET FÖR "Tomater").
2. PLUSTECKEN (+): Om en rad innehåller ett plustecken (+) som skiljer varor åt (t.ex. "Äpplen + Smör"), MÅSTE du dela upp dem i två separata objekt!
3. UTAN PLUSTECKEN: Om en rad INTE har ett plustecken, behåll hela raden som ett enstaka objekt (t.ex. ska "Fransk senap Dijon" förbli ETT objekt).
4. AVDELNING: Koppla varje enskild vara till den MEST PASSANDE avdelningen från listan ovan.
5. SVARSFORMAT: Svara ENDAST med en ren JSON-array av objekt med egenskaperna "vara" och "avdelning".`;
  
  const cleanBase64 = base64Image.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, "");

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": "image/jpeg", "data": cleanBase64 } }
      ]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  let lastErrorMsg = "";

  for (let model of validModels) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    for (let attempt = 1; attempt <= 2; attempt++) {
      const response = UrlFetchApp.fetch(url, options);
      const resCode = response.getResponseCode();

      if (resCode === 200) {
        const result = JSON.parse(response.getContentText());
        try {
          const rawText = result.candidates[0].content.parts[0].text.trim();
          const jsonString = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
          const itemsArray = JSON.parse(jsonString);

          if (Array.isArray(itemsArray)) {
            const created = addMultipleItems(itemsArray, listaId);
            return {
              parsedCount: itemsArray.length,
              createdCount: created ? created.length : 0
            };
          }
          return { parsedCount: 0, createdCount: 0 };
        } catch (e) {
          throw new Error("Kunde inte tolka AI-svaret: " + e.message);
        }
      } else if (resCode === 503) {
        lastErrorMsg = "Googles servrar har hög belastning just nu (503).";
        Utilities.sleep(1500);
      } else {
        const errJson = JSON.parse(response.getContentText());
        lastErrorMsg = (errJson.error && errJson.error.message) ? errJson.error.message : response.getContentText();
        break; 
      }
    }
  }

  throw new Error("Tillfälligt hög belastning hos AI-tjänsten (" + lastErrorMsg + "). Vänta några sekunder och försök igen!");
}

function updateItemAndVaraServer(handlaId, varaId, newVaraNamn, newAvdId, newEnhetId, newAntal, newInfo) {
  newVaraNamn = capitalizeFirstLetter(newVaraNamn);
  newInfo = newInfo ? String(newInfo).trim() : "";
  newAvdId = String(newAvdId);
  newEnhetId = String(newEnhetId || "");
  newAntal = String(newAntal || "").trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const varaSheet = ss.getSheetByName("Vara");
  const handlaSheet = ss.getSheetByName("Handla");

  if (varaSheet) {
    const vData = varaSheet.getDataRange().getDisplayValues();
    const searchName = newVaraNamn.toLowerCase();

    for (let i = 1; i < vData.length; i++) {
      const currentVaraId = String(vData[i][0]).trim();
      const currentVaraNamn = String(vData[i][1] || "").trim().toLowerCase();

      if ((varaId && currentVaraId === String(varaId).trim()) || (currentVaraNamn === searchName)) {
        varaSheet.getRange(i + 1, 2).setValue(newVaraNamn);
        varaSheet.getRange(i + 1, 3).setValue(newAvdId);
        varaSheet.getRange(i + 1, 4).setValue(newEnhetId);
        break;
      }
    }
  }

  if (handlaSheet && handlaId) {
    const hData = handlaSheet.getDataRange().getDisplayValues();
    const searchHandlaId = String(handlaId).trim();

    for (let i = 1; i < hData.length; i++) {
      if (String(hData[i][0]).trim() === searchHandlaId) {
        handlaSheet.getRange(i + 1, 4).setValue(newAntal);
        handlaSheet.getRange(i + 1, 5).setValue(newInfo);
        break;
      }
    }
  }
}

function saveVaraImageServer(handlaId, base64Image) {
  if (!handlaId || !base64Image) throw new Error("Handla-ID eller bilddata saknas.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handlaSheet = ss.getSheetByName("Handla");
  if (!handlaSheet) throw new Error("Fliken Handla saknas.");

  const data = handlaSheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  let oldBildId = "";

  const searchId = String(handlaId).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      rowIndex = i + 1;
      oldBildId = data[i][6] ? String(data[i][6]).trim() : "";
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("Varan hittades inte i handlingslistan.");
  }

  if (oldBildId) {
    deleteDriveFileById(oldBildId);
  }

  const cleanBase64 = base64Image.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, "");
  const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/jpeg", "handla_" + searchId + ".jpg");

  const f1 = DriveApp.getFoldersByName("AppsScript");
  const folderAppsScript = f1.hasNext() ? f1.next() : DriveApp.createFolder("AppsScript");

  const f2 = folderAppsScript.getFoldersByName("Handling");
  const folderHandling = f2.hasNext() ? f2.next() : folderAppsScript.createFolder("Handling");

  const f3 = folderHandling.getFoldersByName("Varubilder");
  const targetFolder = f3.hasNext() ? f3.next() : folderHandling.createFolder("Varubilder");

  const file = targetFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const newFileId = file.getId();

  handlaSheet.getRange(rowIndex, 7).setValue(newFileId);

  return "https://lh3.googleusercontent.com/d/" + newFileId;
}

function deleteVaraImageServer(handlaId) {
  if (!handlaId) throw new Error("Handla-ID saknas.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handlaSheet = ss.getSheetByName("Handla");
  if (!handlaSheet) throw new Error("Fliken Handla saknas.");

  const data = handlaSheet.getDataRange().getDisplayValues();
  const searchId = String(handlaId).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      const bildId = data[i][6] ? String(data[i][6]).trim() : "";
      if (bildId) {
        deleteDriveFileById(bildId);
      }
      handlaSheet.getRange(i + 1, 7).setValue("");
      break;
    }
  }
}

// BUTIKSHANTERING & AVDELNINGSSORTERING
function saveStoreServer(butikId, butikNamn, butikPlats) {
  butikNamn = capitalizeFirstLetter(butikNamn);
  butikPlats = capitalizeFirstLetter(butikPlats);
  if (!butikNamn) throw new Error("Butiksnamn får inte vara tomt.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Butik");
  if (!sheet) {
    sheet = ss.insertSheet("Butik");
    sheet.appendRow(["Butik_ID", "Butik_Namn", "Butik_Plats"]);
  }

  const data = sheet.getDataRange().getDisplayValues();
  const searchId = String(butikId || "").trim();

  if (searchId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        sheet.getRange(i + 1, 2).setValue(butikNamn);
        sheet.getRange(i + 1, 3).setValue(butikPlats);
        return searchId;
      }
    }
  }

  const newButikId = String(new Date().getTime());
  sheet.appendRow([newButikId, butikNamn, butikPlats]);
  return newButikId;
}

function deleteStoreServer(butikId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const butikSheet = ss.getSheetByName("Butik");
  const butikAvdSheet = ss.getSheetByName("Butik_Avdelning");
  const searchId = String(butikId).trim();

  if (butikSheet) {
    const data = butikSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        butikSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  if (butikAvdSheet) {
    const baData = butikAvdSheet.getDataRange().getDisplayValues();
    for (let i = baData.length - 1; i >= 1; i--) {
      if (String(baData[i][0]).trim() === searchId) {
        butikAvdSheet.deleteRow(i + 1);
      }
    }
  }
}

function getStoreDepartmentsServer(butikId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avdSheet = ss.getSheetByName("Avdelning");
  const butikAvdSheet = ss.getSheetByName("Butik_Avdelning");

  if (!avdSheet) return [];

  const avdData = avdSheet.getDataRange().getValues();
  if (avdData.length <= 1) return [];
  avdData.shift();

  const sortMap = new Map();
  if (butikAvdSheet) {
    const baData = butikAvdSheet.getDataRange().getDisplayValues();
    if (baData.length > 1) {
      baData.shift();
      baData.forEach(r => {
        if (String(r[0]).trim() === String(butikId).trim()) {
          sortMap.set(String(r[1]).trim(), Number(r[2]) || 999);
        }
      });
    }
  }

  const result = avdData.map(r => {
    const aId = String(r[0]);
    const aNamn = String(r[1]).trim();
    const sortVal = sortMap.has(aId) ? sortMap.get(aId) : 999;
    return { avdId: aId, avdNamn: aNamn, sortering: sortVal };
  });

  result.sort((a, b) => a.sortering - b.sortering);
  return result;
}

function saveStoreDepartmentOrderServer(butikId, orderedAvdIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let butikAvdSheet = ss.getSheetByName("Butik_Avdelning");

  if (!butikAvdSheet) {
    butikAvdSheet = ss.insertSheet("Butik_Avdelning");
    butikAvdSheet.appendRow(["Butik_ID", "Avdelning_ID", "Avdelning_Sortering"]);
  }

  const data = butikAvdSheet.getDataRange().getDisplayValues();
  const searchId = String(butikId).trim();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === searchId) {
      butikAvdSheet.deleteRow(i + 1);
    }
  }

  const newRows = orderedAvdIds.map((avdId, index) => [String(butikId), String(avdId), index + 1]);
  if (newRows.length > 0) {
    butikAvdSheet.getRange(butikAvdSheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }
}

// AVDELNINGSHANTERING
function saveDepartmentServer(avdId, avdNamn) {
  avdNamn = capitalizeFirstLetter(avdNamn);
  if (!avdNamn) throw new Error("Avdelningsnamn får inte vara tomt.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Avdelning");
  if (!sheet) {
    sheet = ss.insertSheet("Avdelning");
    sheet.appendRow(["Avdelning_ID", "Avdelning_Namn"]);
  }

  const data = sheet.getDataRange().getDisplayValues();
  const searchId = String(avdId || "").trim();

  if (searchId && !searchId.startsWith("temp_")) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        sheet.getRange(i + 1, 2).setValue(avdNamn);
        return searchId;
      }
    }
  }

  const newAvdId = String(new Date().getTime());
  sheet.appendRow([newAvdId, avdNamn]);
  return newAvdId;
}

function checkDepartmentReferencesServer(avdId) {
  avdId = String(avdId).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const varaSheet = ss.getSheetByName("Vara");
  let count = 0;
  let examples = [];

  if (varaSheet) {
    const data = varaSheet.getDataRange().getDisplayValues();
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]).trim() === avdId) {
          count++;
          if (examples.length < 3) {
            examples.push(String(data[i][1]).trim());
          }
        }
      }
    }
  }

  return {
    isReferenced: count > 0,
    count: count,
    examples: examples
  };
}

function deleteDepartmentServer(avdId) {
  const refCheck = checkDepartmentReferencesServer(avdId);
  if (refCheck.isReferenced) {
    throw new Error("Avdelningen kan inte raderas eftersom det finns " + refCheck.count + " vara/varor i databasen som är kopplade till den.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avdSheet = ss.getSheetByName("Avdelning");
  const butikAvdSheet = ss.getSheetByName("Butik_Avdelning");
  const searchId = String(avdId).trim();

  if (avdSheet) {
    const aData = avdSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < aData.length; i++) {
      if (String(aData[i][0]).trim() === searchId) {
        avdSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  if (butikAvdSheet) {
    const baData = butikAvdSheet.getDataRange().getDisplayValues();
    for (let i = baData.length - 1; i >= 1; i--) {
      if (String(baData[i][1]).trim() === searchId) {
        butikAvdSheet.deleteRow(i + 1);
      }
    }
  }
}

// VARUDATABAS
function saveMasterVaraServer(varaId, varaNamn, avdId, enhetId) {
  varaNamn = capitalizeFirstLetter(varaNamn);
  if (!varaNamn) throw new Error("Varunamn får inte vara tomt.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Vara");
  if (!sheet) {
    sheet = ss.insertSheet("Vara");
    sheet.appendRow(["Vara_ID", "Vara_Namn", "Avdelning_ID", "Enhet_ID"]);
  }

  const data = sheet.getDataRange().getDisplayValues();
  avdId = String(avdId || "0");
  enhetId = String(enhetId || "");
  const searchId = String(varaId || "").trim();

  if (searchId && !searchId.startsWith("temp_")) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        sheet.getRange(i + 1, 2).setValue(varaNamn);
        sheet.getRange(i + 1, 3).setValue(avdId);
        sheet.getRange(i + 1, 4).setValue(enhetId);
        return searchId;
      }
    }
  }

  const newVaraId = String(new Date().getTime());
  sheet.appendRow([newVaraId, varaNamn, avdId, enhetId]);
  return newVaraId;
}

function checkVaraReferencesServer(varaId) {
  varaId = String(varaId).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handlaSheet = ss.getSheetByName("Handla");
  const listaSheet = ss.getSheetByName("Lista");

  if (!handlaSheet) return { isReferenced: false, count: 0, listNames: [] };

  const handlaData = handlaSheet.getDataRange().getDisplayValues();
  if (handlaData.length <= 1) return { isReferenced: false, count: 0, listNames: [] };

  const listMap = new Map();
  if (listaSheet) {
    const listaData = listaSheet.getDataRange().getDisplayValues();
    if (listaData.length > 1) {
      listaData.shift();
      listaData.forEach(r => listMap.set(String(r[0]).trim(), String(r[1]).trim()));
    }
  }

  let count = 0;
  const referencedListNames = new Set();

  for (let i = 1; i < handlaData.length; i++) {
    if (String(handlaData[i][2]).trim() === varaId) {
      count++;
      const lId = String(handlaData[i][1]).trim();
      const lName = listMap.get(lId) || "Okänd lista";
      referencedListNames.add(lName);
    }
  }

  return {
    isReferenced: count > 0,
    count: count,
    listNames: Array.from(referencedListNames)
  };
}

function deleteMasterVaraServer(varaId) {
  const refCheck = checkVaraReferencesServer(varaId);
  if (refCheck.isReferenced) {
    throw new Error("Varan kan inte raderas eftersom den finns med på " + refCheck.count + " rad(er) i handlingslistor.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const varaSheet = ss.getSheetByName("Vara");

  if (varaSheet) {
    const data = varaSheet.getDataRange().getDisplayValues();
    const searchId = String(varaId).trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        varaSheet.deleteRow(i + 1);
        break;
      }
    }
  }
}