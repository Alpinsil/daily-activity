const DAILY_ACTIVITY_HEADERS = [
  "ID",
  "Nama",
  "Kategori",
  "Catatan",
  "Waktu",
  "Pengulangan",
  "Mulai",
  "Sampai",
  "Hari",
  "Tanggal_Bulanan",
  "Reminder_Menit",
  "Aktif",
  "Dibuat",
  "Diubah",
];
const DAILY_ACTIVITY_LOG_HEADERS = [
  "Activity_ID",
  "Occurrence_Date",
  "Event_Type",
  "Recipient_ID",
  "Status",
  "Updated_At",
  "Error",
];

function doGet(event) {
  if (event && event.parameter && event.parameter.bridge === "1") {
    const template = HtmlService.createTemplateFromFile(
      "daily_activity_bridge",
    );
    template.allowedOrigin =
      PropertiesService.getScriptProperties().getProperty(
        "DAILY_ACTIVITY_ALLOWED_ORIGIN",
      ) || "";
    const bridgeNonce = String(event.parameter.bridgeNonce || "");
    template.bridgeNonce = /^[a-f0-9]{32}$/.test(bridgeNonce)
      ? bridgeNonce
      : "";
    return template
      .evaluate()
      .setTitle("Daily Activity API")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return dailyActivityJson_({
    status: "ok",
    message: "Daily Activity API is ready.",
  });
}

function doPost(event) {
  try {
    const request = JSON.parse(event?.postData?.contents || "{}");
    const properties = PropertiesService.getScriptProperties();
    const expectedToken = properties.getProperty("DAILY_ACTIVITY_BOT_TOKEN");
    if (!expectedToken || request.token !== expectedToken) {
      return dailyActivityJson_({
        status: "error",
        message: "Akses daily activity ditolak.",
      });
    }

    if (
      request.source !== "daily_activity" &&
      request.source !== "daily_activity_scheduler"
    ) {
      throw new Error("Sumber permintaan tidak valid.");
    }
    return dailyActivityJson_({
      status: "success",
      data: handleDailyActivityRequest_(request),
    });
  } catch (error) {
    console.error("Daily activity request gagal: " + error.stack);
    return dailyActivityJson_({
      status: "error",
      message: error.message || "Permintaan daily activity gagal.",
    });
  }
}

function dailyActivityWebBridgeRequest(request) {
  try {
    if (!request || typeof request !== "object" || Array.isArray(request)) {
      throw new Error("Permintaan website tidak valid.");
    }
    if (JSON.stringify(request).length > 32768) {
      throw new Error("Ukuran permintaan terlalu besar.");
    }

    const allowedActions = [
      "set_sheet",
      "get_data",
      "list_instances",
      "create_activity",
      "update_activity",
      "delete_activity",
      "set_completion",
    ];
    const action = String(request.action || "");
    if (allowedActions.indexOf(action) < 0) {
      throw new Error("Aksi website tidak dikenal.");
    }
    return {
      status: "success",
      data: handleDailyActivityRequest_({
        source: "daily_activity",
        action: action,
        data: request.data || {},
      }),
    };
  } catch (error) {
    console.error("Daily activity web bridge gagal: " + error.stack);
    return {
      status: "error",
      message: error.message || "Permintaan website gagal.",
    };
  }
}

function dailyActivityJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function handleDailyActivityRequest_(request) {
  const action = String(request.action || "");
  const data = request.data || {};
  if (request.source === "daily_activity_scheduler") {
    if (action === "list_due") return listDailyActivityDue_(data);
    if (action === "mark_delivery") return markDailyActivityDelivery_(data);
    throw new Error("Aksi scheduler daily activity tidak dikenal.");
  }

  if (action === "set_sheet") return setDailyActivitySheet_(data.sheetId);
  if (action === "get_data") return getDailyActivityData_();
  if (action === "list_instances") return listDailyActivityInstances_(data.date);
  if (action === "create_activity") return createDailyActivity_(data);
  if (action === "update_activity") return updateDailyActivity_(data);
  if (action === "delete_activity") return deleteDailyActivity_(data.id);
  if (action === "set_completion") return setDailyActivityCompletion_(data);
  throw new Error("Aksi daily activity tidak dikenal.");
}

function setDailyActivitySheet_(sheetId) {
  const cleanId = String(sheetId || "").trim();
  if (!/^[A-Za-z0-9_-]{20,}$/.test(cleanId)) {
    throw new Error("ID Google Sheet tidak valid.");
  }

  return withDailyActivityLock_(function () {
    const spreadsheet = SpreadsheetApp.openById(cleanId);
    verifyDedicatedDailyActivitySpreadsheet_(spreadsheet);
    ensureDailyActivitySheet_(spreadsheet, "Daily_Activities", DAILY_ACTIVITY_HEADERS);
    ensureDailyActivitySheet_(
      spreadsheet,
      "Daily_Activity_Log",
      DAILY_ACTIVITY_LOG_HEADERS,
    );
    PropertiesService.getScriptProperties().setProperty(
      "DAILY_ACTIVITY_SHEET_ID",
      cleanId,
    );
    return {
      sheetId: cleanId,
      sheetName: spreadsheet.getName(),
    };
  });
}

function verifyDedicatedDailyActivitySpreadsheet_(spreadsheet) {
  const allowedTabNames = ["Daily_Activities", "Daily_Activity_Log"];
  const unrelatedData = spreadsheet.getSheets().some(function (sheet) {
    return (
      allowedTabNames.indexOf(sheet.getName()) < 0 &&
      sheet.getLastRow() > 0
    );
  });
  if (unrelatedData) {
    throw new Error(
      "Gunakan Google Sheet khusus daily activity. Spreadsheet berisi tab lain yang memiliki data.",
    );
  }
}

function openDailyActivitySpreadsheet_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty(
    "DAILY_ACTIVITY_SHEET_ID",
  );
  if (!sheetId) {
    throw new Error("Belum ada Google Sheet yang dipilih.");
  }
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const activities = ensureDailyActivitySheet_(
    spreadsheet,
    "Daily_Activities",
    DAILY_ACTIVITY_HEADERS,
  );
  const log = ensureDailyActivitySheet_(
    spreadsheet,
    "Daily_Activity_Log",
    DAILY_ACTIVITY_LOG_HEADERS,
  );
  return { spreadsheet: spreadsheet, activities: activities, log: log };
}

function ensureDailyActivitySheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  const lastRow = sheet.getLastRow();
  const actualHeaders =
    lastRow === 0
      ? []
      : sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (actualHeaders.length === 0 || actualHeaders.every(function (value) {
    return !value;
  })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (headers.some(function (header, index) {
    return actualHeaders[index] !== header;
  })) {
    throw new Error(
      "Header pada tab " + name + " tidak sesuai. Jangan ubah urutan kolom.",
    );
  }
  return sheet;
}

function getDailyActivityData_() {
  const selectedSheet = PropertiesService.getScriptProperties().getProperty(
    "DAILY_ACTIVITY_SHEET_ID",
  );
  if (!selectedSheet) return { sheetId: "", sheetName: "", activities: [] };

  const opened = openDailyActivitySpreadsheet_();
  return {
    sheetId: selectedSheet,
    sheetName: opened.spreadsheet.getName(),
    activities: readDailyActivities_(opened.activities),
  };
}

function readDailyActivities_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, DAILY_ACTIVITY_HEADERS.length)
    .getDisplayValues()
    .map(activityFromRow_);
}

function activityFromRow_(row) {
  return {
    id: row[0],
    title: row[1],
    category: row[2],
    notes: row[3],
    time: row[4],
    recurrence: row[5],
    startDate: row[6],
    endDate: row[7],
    weekdays: row[8]
      ? row[8].split(",").map(function (value) {
          return Number(value);
        })
      : [],
    dayOfMonth: row[9] ? Number(row[9]) : "",
    reminderMinutes: Number(row[10]),
    active: row[11] === "TRUE",
    createdAt: row[12],
    updatedAt: row[13],
  };
}

function validateDailyActivity_(input) {
  if (!input || typeof input !== "object") {
    throw new Error("Data aktivitas tidak valid.");
  }
  const title = String(input.title || "").trim();
  const category = String(input.category || "").trim();
  const notes = String(input.notes || "").trim();
  const time = String(input.time || "").trim();
  const recurrence = String(input.recurrence || "").trim();
  const startDate = String(input.startDate || "").trim();
  const endDate = String(input.endDate || "").trim();
  const allowedRecurrence = ["once", "daily", "weekly", "monthly"];
  const weekdays = Array.isArray(input.weekdays)
    ? [...new Set(input.weekdays.map(Number))]
    : [];
  const reminderMinutes = Number(input.reminderMinutes);
  const dayOfMonth = Number(input.dayOfMonth);

  if (!title || title.length > 100) {
    throw new Error("Nama kegiatan wajib diisi (maksimal 100 karakter).");
  }
  if (category.length > 60 || notes.length > 500) {
    throw new Error("Kategori maksimal 60 dan catatan maksimal 500 karakter.");
  }
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) {
    throw new Error("Waktu kegiatan tidak valid.");
  }
  if (allowedRecurrence.indexOf(recurrence) < 0) {
    throw new Error("Pilih pengulangan yang tersedia.");
  }
  if (!isDailyActivityDate_(startDate)) {
    throw new Error("Tanggal mulai tidak valid.");
  }
  if (endDate && (!isDailyActivityDate_(endDate) || endDate < startDate)) {
    throw new Error("Tanggal selesai harus valid dan tidak sebelum tanggal mulai.");
  }
  if ([10, 15].indexOf(reminderMinutes) < 0) {
    throw new Error("Pengingat harus 10 atau 15 menit sebelumnya.");
  }
  if (
    recurrence === "weekly" &&
    (!weekdays.length ||
      weekdays.some(function (day) {
        return !Number.isInteger(day) || day < 1 || day > 7;
      }))
  ) {
    throw new Error("Pilih minimal satu hari untuk jadwal mingguan.");
  }
  if (
    recurrence === "monthly" &&
    (!Number.isInteger(dayOfMonth) || dayOfMonth < 1 || dayOfMonth > 31)
  ) {
    throw new Error("Tanggal bulanan harus antara 1 dan 31.");
  }

  return {
    title: title,
    category: category,
    notes: notes,
    time: time,
    recurrence: recurrence,
    startDate: startDate,
    endDate: endDate,
    weekdays: recurrence === "weekly" ? weekdays.sort() : [],
    dayOfMonth: recurrence === "monthly" ? dayOfMonth : "",
    reminderMinutes: reminderMinutes,
    active: input.active !== false,
  };
}

function isDailyActivityDate_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return false;
  const date = new Date(value + "T00:00:00Z");
  return !isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function activityRow_(activity, id, createdAt, updatedAt) {
  return [
    id,
    activity.title,
    activity.category,
    activity.notes,
    activity.time,
    activity.recurrence,
    activity.startDate,
    activity.endDate,
    activity.weekdays.join(","),
    activity.dayOfMonth,
    activity.reminderMinutes,
    activity.active ? "TRUE" : "FALSE",
    createdAt,
    updatedAt,
  ];
}

function createDailyActivity_(input) {
  return withDailyActivityLock_(function () {
    const opened = openDailyActivitySpreadsheet_();
    const activity = validateDailyActivity_(input);
    const now = new Date().toISOString();
    const id = Utilities.getUuid();
    opened.activities.appendRow(activityRow_(activity, id, now, now));
    return { id: id };
  });
}

function updateDailyActivity_(input) {
  return withDailyActivityLock_(function () {
    const opened = openDailyActivitySpreadsheet_();
    const rowIndex = findDailyActivityRow_(opened.activities, input.id);
    if (rowIndex < 2) throw new Error("Kegiatan tidak ditemukan.");
    const oldRow = opened.activities
      .getRange(rowIndex, 1, 1, DAILY_ACTIVITY_HEADERS.length)
      .getDisplayValues()[0];
    const activity = validateDailyActivity_(input);
    opened.activities
      .getRange(rowIndex, 1, 1, DAILY_ACTIVITY_HEADERS.length)
      .setValues([
        activityRow_(
          activity,
          oldRow[0],
          oldRow[12],
          new Date().toISOString(),
        ),
      ]);
    return { id: oldRow[0] };
  });
}

function deleteDailyActivity_(id) {
  return withDailyActivityLock_(function () {
    const opened = openDailyActivitySpreadsheet_();
    const rowIndex = findDailyActivityRow_(opened.activities, id);
    if (rowIndex < 2) throw new Error("Kegiatan tidak ditemukan.");
    opened.activities.deleteRow(rowIndex);
    deleteDailyActivityLogs_(opened.log, String(id));
    return { id: String(id) };
  });
}

function findDailyActivityRow_(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = ids.findIndex(function (row) {
    return row[0] === String(id || "");
  });
  return index < 0 ? -1 : index + 2;
}

function deleteDailyActivityLogs_(sheet, activityId) {
  for (let row = sheet.getLastRow(); row >= 2; row -= 1) {
    if (sheet.getRange(row, 1).getDisplayValue() === activityId) {
      sheet.deleteRow(row);
    }
  }
}

function listDailyActivityInstances_(date) {
  if (!isDailyActivityDate_(date)) throw new Error("Tanggal tidak valid.");
  const opened = openDailyActivitySpreadsheet_();
  const activities = readDailyActivities_(opened.activities).filter(function (activity) {
    return activity.active && isDailyActivityOccurrence_(activity, date);
  });
  const completed = readDailyActivityLogs_(opened.log).filter(function (event) {
    return (
      event.occurrenceDate === date &&
      event.eventType === "COMPLETED" &&
      event.status === "DONE"
    );
  });
  return {
    date: date,
    instances: activities.map(function (activity) {
      return {
        ...activity,
        completed: completed.some(function (event) {
          return event.activityId === activity.id;
        }),
      };
    }),
  };
}

function setDailyActivityCompletion_(data) {
  const activityId = String(data.id || "");
  const date = String(data.date || "");
  if (!isDailyActivityDate_(date)) throw new Error("Tanggal tidak valid.");
  return withDailyActivityLock_(function () {
    const opened = openDailyActivitySpreadsheet_();
    const rowIndex = findDailyActivityRow_(opened.activities, activityId);
    if (rowIndex < 2) throw new Error("Kegiatan tidak ditemukan.");
    const activity = activityFromRow_(
      opened.activities
        .getRange(rowIndex, 1, 1, DAILY_ACTIVITY_HEADERS.length)
        .getDisplayValues()[0],
    );
    if (!isDailyActivityOccurrence_(activity, date)) {
      throw new Error("Kegiatan tidak dijadwalkan pada tanggal tersebut.");
    }
    setDailyActivityLog_(
      opened.log,
      activityId,
      date,
      "COMPLETED",
      "",
      data.completed === true ? "DONE" : "",
      "",
    );
    return { id: activityId, date: date, completed: data.completed === true };
  });
}

function isDailyActivityOccurrence_(activity, date) {
  if (date < activity.startDate || (activity.endDate && date > activity.endDate)) {
    return false;
  }
  if (activity.recurrence === "once") return date === activity.startDate;
  if (activity.recurrence === "daily") return true;
  if (activity.recurrence === "weekly") {
    return activity.weekdays.indexOf(isoWeekday_(date)) >= 0;
  }
  if (activity.recurrence === "monthly") {
    return Number(date.slice(8, 10)) === Number(activity.dayOfMonth);
  }
  return false;
}

function isoWeekday_(date) {
  const day = new Date(date + "T00:00:00Z").getUTCDay();
  return day === 0 ? 7 : day;
}

function readDailyActivityLogs_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, DAILY_ACTIVITY_LOG_HEADERS.length)
    .getDisplayValues()
    .map(function (row) {
      return {
        activityId: row[0],
        occurrenceDate: row[1],
        eventType: row[2],
        recipientId: row[3],
        status: row[4],
      };
    });
}

function setDailyActivityLog_(
  sheet,
  activityId,
  date,
  eventType,
  recipientId,
  status,
  error,
) {
  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  if (lastRow >= 2) {
    const rows = sheet
      .getRange(2, 1, lastRow - 1, 4)
      .getDisplayValues();
    const offset = rows.findIndex(function (row) {
      return (
        row[0] === activityId &&
        row[1] === date &&
        row[2] === eventType &&
        row[3] === recipientId
      );
    });
    if (offset >= 0) rowIndex = offset + 2;
  }

  if (!status) {
    if (rowIndex >= 2) sheet.deleteRow(rowIndex);
    return;
  }
  const values = [
    activityId,
    date,
    eventType,
    recipientId,
    status,
    new Date().toISOString(),
    String(error || "").slice(0, 200),
  ];
  if (rowIndex >= 2) {
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function listDailyActivityDue_(data) {
  const now = data.now ? new Date(data.now) : new Date();
  if (isNaN(now.getTime())) throw new Error("Waktu scheduler tidak valid.");
  const timeZone = Session.getScriptTimeZone();
  const nowKey = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm");
  const today = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  const tomorrow = addDailyActivityDays_(today, 1);
  const targets = Array.isArray(data.targets)
    ? [...new Set(data.targets.map(String).filter(Boolean))]
    : [];
  if (!targets.length) return { deliveries: [] };

  const opened = openDailyActivitySpreadsheet_();
  const activities = readDailyActivities_(opened.activities);
  const logs = readDailyActivityLogs_(opened.log);
  const deliveries = [];
  activities.forEach(function (activity) {
    if (!activity.active) return;
    [today, tomorrow].forEach(function (occurrenceDate) {
      if (!isDailyActivityOccurrence_(activity, occurrenceDate)) return;
      if (
        !isDailyActivityReminderDue_(
          nowKey,
          occurrenceDate,
          activity.time,
          activity.reminderMinutes,
        )
      ) {
        return;
      }

      targets.forEach(function (targetJid) {
        const alreadySent = logs.some(function (event) {
          return (
            event.activityId === activity.id &&
            event.occurrenceDate === occurrenceDate &&
            event.eventType === "REMINDER" &&
            event.recipientId === targetJid &&
            event.status === "SENT"
          );
        });
        if (alreadySent) return;
        deliveries.push({
          activityId: activity.id,
          occurrenceDate: occurrenceDate,
          targetJid: targetJid,
          message:
            "⏰ *Pengingat kegiatan*\n" +
            activity.title +
            "\nMulai pukul " +
            activity.time +
            " (sekitar " +
            activity.reminderMinutes +
            " menit lagi)." +
            (activity.category ? "\nKategori: " + activity.category : "") +
            (activity.notes ? "\nCatatan: " + activity.notes : ""),
        });
      });
    });
  });
  return { deliveries: deliveries };
}

function isDailyActivityReminderDue_(nowKey, occurrenceDate, time, minutes) {
  const startMinute =
    Number(time.slice(0, 2)) * 60 + Number(time.slice(3, 5));
  let dueDate = occurrenceDate;
  let dueMinute = startMinute - minutes;
  if (dueMinute < 0) {
    dueDate = addDailyActivityDays_(occurrenceDate, -1);
    dueMinute += 1440;
  }
  const dueTime =
    String(Math.floor(dueMinute / 60)).padStart(2, "0") +
    ":" +
    String(dueMinute % 60).padStart(2, "0");
  return (
    nowKey >= dueDate + "T" + dueTime &&
    nowKey < occurrenceDate + "T" + time
  );
}

function markDailyActivityDelivery_(data) {
  const activityId = String(data.activityId || "");
  const occurrenceDate = String(data.occurrenceDate || "");
  const targetJid = String(data.targetJid || "");
  if (!isDailyActivityDate_(occurrenceDate) || !targetJid) {
    throw new Error("Data tujuan reminder tidak valid.");
  }
  return withDailyActivityLock_(function () {
    const opened = openDailyActivitySpreadsheet_();
    if (findDailyActivityRow_(opened.activities, activityId) < 2) {
      throw new Error("Kegiatan reminder tidak ditemukan.");
    }
    setDailyActivityLog_(
      opened.log,
      activityId,
      occurrenceDate,
      "REMINDER",
      targetJid,
      data.delivered === true ? "SENT" : "FAILED",
      data.errorMessage || "",
    );
    return { recorded: true };
  });
}

function addDailyActivityDays_(date, amount) {
  const value = new Date(date + "T00:00:00Z");
  value.setUTCDate(value.getUTCDate() + amount);
  return value.toISOString().slice(0, 10);
}

function withDailyActivityLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}
