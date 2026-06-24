const FAMILY_CODE = "dadrew";

function doGet(e) {
  setupSheets();

  const action = e.parameter.action || "list";
  const callback = e.parameter.callback || "";
  const code = e.parameter.code || "";

  if (code !== FAMILY_CODE) {
    return returnData(callback, {
      ok: false,
      success: false,
      error: "Wrong family code"
    });
  }

  if (action === "list") {
    return returnData(callback, {
      ok: true,
      success: true,
      adventures: getRows("Adventure Log"),
      photos: getRows("Memory Photos"),
      suggestions: getRows("Quest Suggestions"),
      rewards: getRows("Rewards"),
      rallies: getRows("Rallies"),
      locations: latestLocations()
    });
  }

  return returnData(callback, {
    ok: false,
    success: false,
    error: "Unknown action"
  });
}

function doPost(e) {
  setupSheets();

  try {
    const data = JSON.parse(e.postData.contents || "{}");

    if (data.familyCode !== FAMILY_CODE) {
      return json({
        ok: false,
        success: false,
        error: "Wrong family code"
      });
    }

    if (data.type === "adventure") {
      appendRow("Adventure Log", [
        new Date(),
        data.person || "",
        data.questId || "",
        data.questTitle || "",
        data.status || "",
        data.loved || "",
        data.rating || "",
        data.notes || ""
      ]);
    }

    if (data.type === "photo") {
      appendRow("Memory Photos", [
        new Date(),
        data.person || "",
        data.questId || "",
        data.questTitle || "",
        data.caption || "",
        data.imageUrl || "",
        data.publicId || ""
      ]);
    }

    if (data.type === "suggestion") {
      appendRow("Quest Suggestions", [
        data.suggestionId || Utilities.getUuid(),
        new Date(),
        data.person || "",
        data.title || "",
        data.notes || "",
        0,
        0,
        "{}"
      ]);
    }

    if (data.type === "vote") {
      saveSuggestionVote(data.suggestionId, data.person, data.vote);
    }

    if (data.type === "reward") {
      appendRow("Rewards", [
        new Date(),
        data.person || "",
        data.reward || "",
        data.reason || "",
        data.claimed || "No"
      ]);
    }

    if (data.type === "rally") {
      appendRow("Rallies", [
        data.rallyId || Utilities.getUuid(),
        new Date(),
        data.familyCode || FAMILY_CODE,
        data.person || "",
        data.title || "",
        data.when || "",
        data.where || "",
        data.message || "",
        JSON.stringify(data.responses || {})
      ]);
    }

    if (data.type === "rsvp") {
      saveRsvp(data.rallyId, data.person, data.response);
    }

    if (data.type === "closeRally") {
      closeRally(data.rallyId);
    }

    if (data.type === "deleteRally") {
      deleteRally(data.rallyId);
    }

    if (data.type === "location") {
      appendRow("Locations", [
        new Date(),
        data.familyCode || FAMILY_CODE,
        data.person || "",
        data.mode || "",
        data.lat || "",
        data.lng || "",
        data.accuracy || "",
        data.timestamp || new Date().toISOString()
      ]);
    }

    return json({
      ok: true,
      success: true
    });

  } catch (err) {
    return json({
      ok: false,
      success: false,
      error: String(err)
    });
  }
}

function setupSheets() {
  getSheet("Adventure Log");
  getSheet("Memory Photos");
  getSheet("Quest Suggestions");
  getSheet("Rewards");
  getSheet("Rallies");
  getSheet("Locations");
}

function getSheet(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }

  if (sheet.getLastRow() === 0) {
    if (tabName === "Adventure Log") {
      sheet.appendRow(["timestamp", "person", "questId", "questTitle", "status", "loved", "rating", "notes"]);
    }

    if (tabName === "Memory Photos") {
      sheet.appendRow(["timestamp", "person", "questId", "questTitle", "caption", "imageUrl", "publicId"]);
    }

    if (tabName === "Quest Suggestions") {
      sheet.appendRow(["suggestionId", "timestamp", "person", "title", "notes", "yay", "nay", "voters"]);
    }

    if (tabName === "Rewards") {
      sheet.appendRow(["timestamp", "person", "reward", "reason", "claimed"]);
    }

    if (tabName === "Rallies") {
      sheet.appendRow(["rallyId", "timestamp", "familyCode", "person", "title", "when", "where", "message", "responses", "status"]);
    }

    if (tabName === "Locations") {
      sheet.appendRow(["timestamp", "familyCode", "person", "mode", "lat", "lng", "accuracy", "deviceTimestamp"]);
    }
  }

  return sheet;
}

function appendRow(tabName, values) {
  const sheet = getSheet(tabName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (tabName === "Rallies" && headers.indexOf("status") !== -1 && values.length === 9) {
    values.push("open");
  }

  sheet.appendRow(values);
}

function getRows(tabName) {
  const sheet = getSheet(tabName);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];

  return values.slice(1)
    .map(function(row) {
      const item = {};
      headers.forEach(function(header, index) {
        item[header] = row[index];
      });
      return item;
    })
    .filter(function(item) {
      if (tabName === "Rallies") {
        return !item.status || item.status === "open";
      }
      return true;
    });
}

function saveSuggestionVote(suggestionId, person, vote) {
  const sheet = getSheet("Quest Suggestions");
  const values = sheet.getDataRange().getValues();

  if (!suggestionId || values.length < 2) {
    return;
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(suggestionId)) {
      let yay = Number(values[r][5] || 0);
      let nay = Number(values[r][6] || 0);
      let voters = {};

      try {
        voters = JSON.parse(values[r][7] || "{}");
      } catch (err) {
        voters = {};
      }

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

  if (!rallyId || values.length < 2) {
    return;
  }

  const headers = values[0];
  const rallyIdCol = headers.indexOf("rallyId");
  const responsesCol = headers.indexOf("responses");

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][rallyIdCol]) === String(rallyId)) {
      let responses = {};

      try {
        responses = JSON.parse(values[r][responsesCol] || "{}");
      } catch (err) {
        responses = {};
      }

      responses[person] = response;

      sheet.getRange(r + 1, responsesCol + 1).setValue(JSON.stringify(responses));
      return;
    }
  }
}

function closeRally(rallyId) {
  const sheet = getSheet("Rallies");
  const values = sheet.getDataRange().getValues();

  if (!rallyId || values.length < 2) {
    return;
  }

  const headers = values[0];
  const rallyIdCol = headers.indexOf("rallyId");
  const statusCol = headers.indexOf("status");

  if (statusCol === -1) {
    return;
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][rallyIdCol]) === String(rallyId)) {
      sheet.getRange(r + 1, statusCol + 1).setValue("closed");
      return;
    }
  }
}

function deleteRally(rallyId) {
  const sheet = getSheet("Rallies");
  const values = sheet.getDataRange().getValues();

  if (!rallyId || values.length < 2) {
    return;
  }

  const headers = values[0];
  const rallyIdCol = headers.indexOf("rallyId");

  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][rallyIdCol]) === String(rallyId)) {
      sheet.deleteRow(r + 1);
      return;
    }
  }
}

function latestLocations() {
  const rows = getRows("Locations");
  const latest = {};

  rows.forEach(function(row) {
    latest[row.person] = row;
  });

  return Object.keys(latest).map(function(key) {
    return latest[key];
  });
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

function returnData(callback, obj) {
  if (callback) {
    return jsonp(callback, obj);
  }

  return json(obj);
}
