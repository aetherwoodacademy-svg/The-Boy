const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const FAMILY_CODE = "dadrew";

function doGet(e) {
  const action = e.parameter.action || "list";
  const callback = e.parameter.callback || "callback";
  const code = e.parameter.code || "";

  if (code !== FAMILY_CODE) {
    return jsonp(callback, { ok: false, error: "Wrong family code" });
  }

  if (action === "list") {
    return jsonp(callback, {
      ok: true,
      adventures: getRows("Adventure Log"),
      photos: getRows("Memory Photos"),
      suggestions: getRows("Quest Suggestions"),
      rewards: getRows("Rewards"),
      rallies: getRows("Rallies"),
      locations: latestLocations()
    });
  }

  return jsonp(callback, { ok: false, error: "Unknown action" });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    if (data.familyCode !== FAMILY_CODE) {
      return json({ ok: false, error: "Wrong family code" });
    }

    if (data.type === "adventure") {
      appendRow("Adventure Log", [new Date(), data.person, data.questId, data.questTitle, data.status, data.loved, data.rating, data.notes]);
    }

    if (data.type === "photo") {
      appendRow("Memory Photos", [new Date(), data.person, data.questId, data.questTitle, data.caption, data.imageUrl, data.publicId]);
    }

    if (data.type === "suggestion") {
      appendRow("Quest Suggestions", [data.suggestionId || Utilities.getUuid(), new Date(), data.person, data.title, data.notes, 0, 0, "{}"]);
    }

    if (data.type === "vote") {
      saveSuggestionVote(data.suggestionId, data.person, data.vote);
    }

    if (data.type === "reward") {
      appendRow("Rewards", [new Date(), data.person, data.reward, data.reason, data.claimed || "No"]);
    }

    if (data.type === "rally") {
      appendRow("Rallies", [data.rallyId || Utilities.getUuid(), new Date(), data.person, data.title, data.when, data.where, data.message, "{}"]);
    }

    if (data.type === "rsvp") {
      saveRsvp(data.rallyId, data.person, data.response);
    }

    if (data.type === "location") {
      appendRow("Locations", [new Date(), data.person, data.mode, data.lat, data.lng, data.accuracy, data.timestamp]);
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function saveSuggestionVote(suggestionId, person, vote) {
  const sheet = getSheet("Quest Suggestions");
  const values = sheet.getDataRange().getValues();
  if (!suggestionId || values.length < 2) return;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(suggestionId)) {
      let yay = Number(values[r][5] || 0);
      let nay = Number(values[r][6] || 0);
      let voters = {};
      try { voters = JSON.parse(values[r][7] || "{}"); } catch (err) { voters = {}; }

      const previous = voters[person];
      if (previous === "yay") yay = Math.max(0, yay - 1);
      if (previous === "nay") nay = Math.max(0, nay - 1);

      voters[person] = vote;
      if (vote === "yay") yay++;
      if (vote === "nay") nay++;

      sheet.getRange(r + 1, 6).setValue(yay);
      sheet.getRange(r + 1, 7).setValue(nay);
      sheet.getRange(r + 1, 8).setValue(JSON.stringify(voters));
      return;
    }
  }
}

function saveRsvp(rallyId, person, response) {
  const sheet = getSheet("Rallies");
  const values = sheet.getDataRange().getValues();
  if (!rallyId || values.length < 2) return;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(rallyId)) {
      let responses = {};
      try { responses = JSON.parse(values[r][7] || "{}"); } catch (err) { responses = {}; }
      responses[person] = response;
      sheet.getRange(r + 1, 8).setValue(JSON.stringify(responses));
      return;
    }
  }
}

function latestLocations() {
  const rows = getRows("Locations");
  const latest = {};
  rows.forEach(row => {
    latest[row.person] = row;
  });
  return Object.keys(latest).map(k => latest[k]);
}

function appendRow(tabName, values) {
  const sheet = getSheet(tabName);
  sheet.appendRow(values);
}

function getRows(tabName) {
  const sheet = getSheet(tabName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const item = {};
    headers.forEach((h, i) => item[h] = row[i]);
    return item;
  });
}

function getSheet(tabName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    if (tabName === "Adventure Log") sheet.appendRow(["timestamp", "person", "questId", "questTitle", "status", "loved", "rating", "notes"]);
    if (tabName === "Memory Photos") sheet.appendRow(["timestamp", "person", "questId", "questTitle", "caption", "imageUrl", "publicId"]);
    if (tabName === "Quest Suggestions") sheet.appendRow(["suggestionId", "timestamp", "person", "title", "notes", "yay", "nay", "voters"]);
    if (tabName === "Rewards") sheet.appendRow(["timestamp", "person", "reward", "reason", "claimed"]);
    if (tabName === "Rallies") sheet.appendRow(["rallyId", "timestamp", "person", "title", "when", "where", "message", "responses"]);
    if (tabName === "Locations") sheet.appendRow(["timestamp", "person", "mode", "lat", "lng", "accuracy", "deviceTimestamp"]);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp(callback, obj) {
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(obj) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
