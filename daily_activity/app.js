const API_BASE = "/api/daily-activity";
const GAS_BRIDGE_URL = window.DAILY_ACTIVITY_GAS_BRIDGE_URL || "";
const SHEET_STORAGE_KEY = "ritme.dailyActivity.sheetId";
const apiBridgeRequests = new Map();
const apiBridgeReadyWaiters = [];
let apiBridgeFrame = null;
let apiBridgeReady = false;
let apiBridgeConfigurationError = "";
let apiBridgeTargetWindow = null;
let apiBridgeTargetOrigin = "";
let apiBridgeNonce = "";
let apiBridgeRequestId = 0;
const state = {
  sheetId: "",
  sheetName: "",
  activities: [],
  instances: [],
};
const dayLabels = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"];

document.addEventListener("DOMContentLoaded", () => {
  document.querySelector("#today-label").textContent = new Intl.DateTimeFormat(
    "id-ID",
    { weekday: "long", day: "numeric", month: "long", year: "numeric" },
  ).format(new Date());
  document.querySelector("#start-date-input").value = localDate();
  document.querySelector("#day-of-month-input").value = Number(
    localDate().slice(8, 10),
  );
  restoreSheetInput();
  bindForms();
  bindListActions();
  updateRepeatFields();
  initializeApiBridge();
  loadData().catch(showError);
  registerServiceWorker();
});

function localDate(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function restoreSheetInput() {
  try {
    document.querySelector("#sheet-input").value =
      window.localStorage.getItem(SHEET_STORAGE_KEY) || "";
  } catch (error) {
    console.error("Tidak dapat membaca ID Sheet dari penyimpanan lokal:", error);
    showMessage(
      "Browser tidak dapat membaca penyimpanan lokal. ID Google Sheet tidak bisa dipulihkan otomatis.",
    );
  }
}

function persistSheetId(sheetId) {
  try {
    window.localStorage.setItem(SHEET_STORAGE_KEY, String(sheetId));
    return true;
  } catch (error) {
    console.error("Tidak dapat menyimpan ID Sheet secara lokal:", error);
    showMessage(
      "Google Sheet terhubung, tetapi browser tidak dapat menyimpan ID Sheet di perangkat ini.",
    );
    return false;
  }
}

async function apiRequest(path = "", options = {}) {
  if (GAS_BRIDGE_URL || apiBridgeConfigurationError) {
    const request = createBridgeRequest(path, options);
    const result = await callApiBridge(request);
    if (result?.status !== "success") {
      throw new Error(result?.message || "Permintaan Google Sheet gagal.");
    }
    return result.data;
  }

  const response = await fetch(`${API_BASE}${path}`, {
    method: options.method || "GET",
    headers: options.body
      ? { "Content-Type": "application/json" }
      : undefined,
    body: options.body ? JSON.stringify(options.body) : undefined,
    cache: "no-store",
  });
  let result;
  try {
    result = await response.json();
  } catch {
    throw new Error(unavailableApiMessage());
  }
  if (!response.ok) {
    throw new Error(result.error || "Permintaan gagal. Coba lagi.");
  }
  return result;
}

function createBridgeRequest(path, options) {
  const method = options.method || "GET";
  const body = options.body || {};
  if (method === "GET" && !path) return { action: "get_data", data: {} };
  if (method === "POST" && path === "/sheet") {
    return {
      action: "set_sheet",
      data: { sheetId: parseSpreadsheetId(body.sheet) },
    };
  }
  if (method === "GET" && path.startsWith("/instances?")) {
    const url = new URL(path, location.origin);
    return {
      action: "list_instances",
      data: { date: url.searchParams.get("date") || "" },
    };
  }
  if (method === "POST" && path === "/activities") {
    return { action: "create_activity", data: body };
  }

  const activityMatch = path.match(
    /^\/activities\/([^/?]+)(\/completion)?(?:\?.*)?$/,
  );
  if (!activityMatch) throw new Error("Aksi aplikasi tidak dikenal.");
  const id = decodeURIComponent(activityMatch[1]);
  if (method === "PUT" && !activityMatch[2]) {
    return {
      action: "update_activity",
      data: { ...body, id },
    };
  }
  if (method === "DELETE" && !activityMatch[2]) {
    return { action: "delete_activity", data: { id } };
  }
  if (method === "POST" && activityMatch[2]) {
    return {
      action: "set_completion",
      data: { id, date: body.date, completed: body.completed === true },
    };
  }
  throw new Error("Metode permintaan tidak didukung.");
}

function parseSpreadsheetId(value) {
  const input = String(value || "").trim();
  const match = input.match(
    /^(?:https:\/\/docs\.google\.com\/spreadsheets\/d\/)?([A-Za-z0-9_-]{20,})(?:\/.*)?$/,
  );
  if (!match) throw new Error("URL atau ID Google Sheet tidak valid.");
  return match[1];
}

function initializeApiBridge() {
  if (!GAS_BRIDGE_URL) return;

  let bridgeUrl;
  try {
    bridgeUrl = new URL(GAS_BRIDGE_URL);
    if (
      bridgeUrl.protocol !== "https:" ||
      bridgeUrl.hostname !== "script.google.com" ||
      !/^\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(bridgeUrl.pathname)
    ) {
      throw new Error("URL deployment Google Apps Script tidak valid.");
    }
  } catch (error) {
    apiBridgeConfigurationError = error.message;
    return;
  }

  const nonceBytes = new Uint8Array(16);
  window.crypto.getRandomValues(nonceBytes);
  apiBridgeNonce = Array.from(nonceBytes, (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");

  apiBridgeFrame = document.createElement("iframe");
  apiBridgeFrame.title = "Penghubung ke Google Sheets";
  apiBridgeFrame.tabIndex = -1;
  apiBridgeFrame.setAttribute("aria-hidden", "true");
  apiBridgeFrame.style.cssText =
    "position:fixed;width:1px;height:1px;left:-2px;bottom:-2px;border:0;opacity:0;pointer-events:none;";

  window.addEventListener("message", (event) => {
    if (!isApiBridgeFrame(event.source, event.origin)) return;
    const message = event.data;
    if (message?.type === "daily-activity-bridge-ready") {
      if (message.nonce !== apiBridgeNonce) return;
      apiBridgeTargetWindow = event.source;
      apiBridgeTargetOrigin = event.origin;
      apiBridgeReady = true;
      apiBridgeReadyWaiters.splice(0).forEach((resolve) => resolve());
      return;
    }
    if (
      event.source !== apiBridgeTargetWindow ||
      event.origin !== apiBridgeTargetOrigin ||
      message?.type !== "daily-activity-bridge-response" ||
      message.nonce !== apiBridgeNonce
    ) {
      return;
    }

    const pending = apiBridgeRequests.get(message.id);
    if (!pending) return;
    clearTimeout(pending.timeout);
    apiBridgeRequests.delete(message.id);
    if (message.error) pending.reject(new Error(message.error));
    else pending.resolve(message.result);
  });

  bridgeUrl.searchParams.set("bridge", "1");
  bridgeUrl.searchParams.set("bridgeNonce", apiBridgeNonce);
  apiBridgeFrame.src = bridgeUrl.toString();
  document.body.appendChild(apiBridgeFrame);
}

function isApiBridgeFrame(source, origin) {
  if (!apiBridgeFrame || !source) return false;
  try {
    const sender = new URL(origin);
    return (
      sender.protocol === "https:" &&
      (sender.hostname === "script.google.com" ||
        sender.hostname === "script.googleusercontent.com" ||
        sender.hostname.endsWith(".script.googleusercontent.com") ||
        /^[a-z0-9-]+-script\.googleusercontent\.com$/.test(sender.hostname))
    );
  } catch {
    return false;
  }
}

function waitForApiBridge() {
  if (apiBridgeReady) return Promise.resolve();
  if (apiBridgeConfigurationError) {
    return Promise.reject(new Error(apiBridgeConfigurationError));
  }
  if (!apiBridgeFrame) {
    return Promise.reject(
      new Error("Google Apps Script bridge belum dikonfigurasi."),
    );
  }
  return new Promise((resolve, reject) => {
    const resolveWhenReady = () => {
      clearTimeout(timeout);
      resolve();
    };
    const timeout = setTimeout(() => {
      const index = apiBridgeReadyWaiters.indexOf(resolveWhenReady);
      if (index >= 0) apiBridgeReadyWaiters.splice(index, 1);
      reject(
        new Error(
          "Google Apps Script bridge belum merespons. Periksa deployment dan Script Property DAILY_ACTIVITY_ALLOWED_ORIGIN.",
        ),
      );
    }, 20000);
    apiBridgeReadyWaiters.push(resolveWhenReady);
  });
}

async function callApiBridge(request) {
  await waitForApiBridge();
  return new Promise((resolve, reject) => {
    apiBridgeRequestId += 1;
    const id = `${Date.now()}-${apiBridgeRequestId}`;
    const timeout = setTimeout(() => {
      apiBridgeRequests.delete(id);
      reject(new Error("Google Apps Script belum merespons. Coba lagi nanti."));
    }, 30000);
    apiBridgeRequests.set(id, { resolve, reject, timeout });
    apiBridgeTargetWindow.postMessage(
      {
        type: "daily-activity-bridge-request",
        id,
        nonce: apiBridgeNonce,
        request,
      },
      apiBridgeTargetOrigin,
    );
  });
}

function unavailableApiMessage() {
  if (location.hostname.endsWith(".github.io")) {
    return "Bridge Google Apps Script belum dikonfigurasi. Periksa DAILY_ACTIVITY_GAS_BRIDGE_URL pada deployment GitHub Pages.";
  }
  return "API backend tidak tersedia. Live Server hanya menyajikan file statis. Jalankan npm.cmd run start:daily-activity dari folder bot, lalu buka http://127.0.0.1:3001/daily-activity/.";
}

async function loadData() {
  setSyncStatus("Menyinkronkan...");
  const data = await apiRequest();
  state.sheetId = data.sheetId || "";
  state.sheetName = data.sheetName || "";
  state.activities = Array.isArray(data.activities) ? data.activities : [];
  if (state.sheetId) {
    document.querySelector("#sheet-input").value = state.sheetId;
    persistSheetId(state.sheetId);
  }
  document.querySelector("#sheet-name").textContent = state.sheetId
    ? `Terhubung: ${state.sheetName || "Google Sheet"}`
    : "Belum ada Sheet yang terhubung.";

  if (state.sheetId) {
    const instances = await apiRequest(`/instances?date=${localDate()}`);
    state.instances = Array.isArray(instances.instances)
      ? instances.instances
      : [];
    setSyncStatus("Tersinkron");
  } else {
    state.instances = [];
    setSyncStatus("Belum tersinkron");
  }
  render();
}

function bindForms() {
  document.querySelector("#sheet-form").addEventListener("submit", async (event) => {
    event.preventDefault();
    const input = document.querySelector("#sheet-input");
    const button = event.currentTarget.querySelector("button");
    button.disabled = true;
    clearMessage();
    try {
      const result = await apiRequest("/sheet", {
        method: "POST",
        body: { sheet: input.value.trim() },
      });
      state.sheetId = result.sheetId || "";
      state.sheetName = result.sheetName || "";
      if (!state.sheetId) {
        throw new Error("Google Sheet berhasil diproses, tetapi ID Sheet tidak diterima.");
      }
      input.value = state.sheetId;
      const savedLocally = persistSheetId(state.sheetId);
      await loadData();
      if (savedLocally) {
        showMessage("Google Sheet berhasil dihubungkan. Tab kegiatan dan progres sudah siap.");
      }
    } catch (error) {
      showError(error);
    } finally {
      button.disabled = false;
    }
  });

  document.querySelector("#activity-form").addEventListener("submit", saveActivity);
  document.querySelector("#repeat-input").addEventListener("change", updateRepeatFields);
  document.querySelector("#start-date-input").addEventListener("change", (event) => {
    if (document.querySelector("#repeat-input").value === "monthly") {
      document.querySelector("#day-of-month-input").value = Number(
        event.target.value.slice(8, 10),
      );
    }
  });
  document.querySelector("#cancel-edit").addEventListener("click", resetActivityForm);
}

async function saveActivity(event) {
  event.preventDefault();
  if (!state.sheetId) {
    showError(new Error("Hubungkan Google Sheet sebelum menyimpan kegiatan."));
    return;
  }
  const form = event.currentTarget;
  const id = document.querySelector("#activity-id").value;
  const recurrence = document.querySelector("#repeat-input").value;
  const weekdays = [...form.querySelectorAll('input[name="weekday"]:checked')]
    .map((input) => Number(input.value));
  if (recurrence === "weekly" && weekdays.length === 0) {
    showError(new Error("Pilih setidaknya satu hari untuk jadwal mingguan."));
    return;
  }

  const data = {
    title: form.elements.title.value.trim(),
    category: form.elements.category.value.trim(),
    notes: form.elements.notes.value.trim(),
    time: form.elements.time.value,
    recurrence,
    startDate: form.elements.startDate.value,
    endDate: form.elements.endDate.value,
    weekdays,
    dayOfMonth: Number(form.elements.dayOfMonth.value),
    reminderMinutes: Number(form.elements.reminderMinutes.value),
    active: form.elements.active.checked,
  };
  const saveButton = document.querySelector("#save-button");
  saveButton.disabled = true;
  clearMessage();
  try {
    await apiRequest(id ? `/activities/${encodeURIComponent(id)}` : "/activities", {
      method: id ? "PUT" : "POST",
      body: data,
    });
    resetActivityForm();
    await loadData();
    showMessage(id ? "Kegiatan berhasil diperbarui." : "Kegiatan berhasil ditambahkan.");
  } catch (error) {
    showError(error);
  } finally {
    saveButton.disabled = false;
  }
}

function updateRepeatFields() {
  const recurrence = document.querySelector("#repeat-input").value;
  document.querySelector("#weekdays-field").hidden = recurrence !== "weekly";
  document.querySelector("#monthly-field").hidden = recurrence !== "monthly";
  document.querySelector("#end-date-field").hidden = recurrence === "once";
}

function bindListActions() {
  document.querySelector("#all-list").addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const activity = state.activities.find((item) => item.id === button.dataset.id);
    if (!activity) return;
    if (button.dataset.action === "edit") {
      fillActivityForm(activity);
      window.scrollTo({ top: document.querySelector(".form-panel").offsetTop - 78, behavior: "smooth" });
      return;
    }
    if (button.dataset.action === "delete") {
      if (!window.confirm(`Hapus kegiatan "${activity.title}" dan progresnya?`)) return;
      button.disabled = true;
      try {
        await apiRequest(`/activities/${encodeURIComponent(activity.id)}`, {
          method: "DELETE",
        });
        await loadData();
        showMessage("Kegiatan dan progresnya sudah dihapus.");
      } catch (error) {
        showError(error);
      } finally {
        button.disabled = false;
      }
    }
  });

  document.querySelector("#today-list").addEventListener("change", async (event) => {
    const checkbox = event.target.closest('input[type="checkbox"][data-id]');
    if (!checkbox) return;
    checkbox.disabled = true;
    try {
      await apiRequest(
        `/activities/${encodeURIComponent(checkbox.dataset.id)}/completion`,
        {
          method: "POST",
          body: { date: localDate(), completed: checkbox.checked },
        },
      );
      await loadData();
    } catch (error) {
      checkbox.checked = !checkbox.checked;
      showError(error);
    } finally {
      checkbox.disabled = false;
    }
  });
}

function fillActivityForm(activity) {
  const form = document.querySelector("#activity-form");
  document.querySelector("#activity-id").value = activity.id;
  form.elements.title.value = activity.title;
  form.elements.category.value = activity.category || "";
  form.elements.notes.value = activity.notes || "";
  form.elements.time.value = activity.time;
  form.elements.recurrence.value = activity.recurrence;
  form.elements.startDate.value = activity.startDate;
  form.elements.endDate.value = activity.endDate || "";
  form.elements.dayOfMonth.value = activity.dayOfMonth || "";
  form.elements.reminderMinutes.value = String(activity.reminderMinutes || 10);
  form.elements.active.checked = activity.active;
  form.querySelectorAll('input[name="weekday"]').forEach((input) => {
    input.checked = (activity.weekdays || []).includes(Number(input.value));
  });
  document.querySelector("#form-title").textContent = "Ubah kegiatan";
  document.querySelector("#save-button").textContent = "Simpan perubahan";
  document.querySelector("#cancel-edit").hidden = false;
  updateRepeatFields();
}

function resetActivityForm() {
  const form = document.querySelector("#activity-form");
  form.reset();
  document.querySelector("#activity-id").value = "";
  document.querySelector("#form-title").textContent = "Tambah kegiatan";
  document.querySelector("#save-button").textContent = "Simpan kegiatan";
  document.querySelector("#cancel-edit").hidden = true;
  document.querySelector("#start-date-input").value = localDate();
  document.querySelector("#day-of-month-input").value = Number(
    localDate().slice(8, 10),
  );
  updateRepeatFields();
}

function render() {
  renderToday();
  renderActivities();
}

function renderToday() {
  const container = document.querySelector("#today-list");
  container.replaceChildren();
  const completed = state.instances.filter((item) => item.completed).length;
  document.querySelector("#progress-count").textContent =
    `${completed} / ${state.instances.length}`;
  if (!state.sheetId) {
    container.append(createEmptyState("Hubungkan Google Sheet untuk melihat jadwal."));
    return;
  }
  if (!state.instances.length) {
    container.append(
      createEmptyState(
        state.activities.length
          ? "Tidak ada kegiatan terjadwal hari ini."
          : "Belum ada kegiatan. Tambahkan kegiatan pertama di formulir.",
      ),
    );
    return;
  }
  state.instances
    .slice()
    .sort((a, b) => a.time.localeCompare(b.time))
    .forEach((activity) => {
      const card = document.createElement("article");
      card.className = `activity-card${activity.completed ? " is-completed" : ""}`;
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "activity-check";
      checkbox.checked = activity.completed;
      checkbox.dataset.id = activity.id;
      checkbox.setAttribute("aria-label", `Tandai ${activity.title} selesai`);
      card.append(checkbox, createActivityMain(activity, true));
      container.append(card);
    });
}

function renderActivities() {
  const container = document.querySelector("#all-list");
  container.replaceChildren();
  const count = state.activities.length;
  document.querySelector("#activity-count").textContent =
    `${count} kegiatan tersimpan`;
  if (!state.sheetId || !count) {
    container.append(
      createEmptyState(
        state.sheetId
          ? "Belum ada kegiatan. Tambahkan kegiatan pertama di formulir."
          : "Hubungkan Google Sheet untuk memulai.",
      ),
    );
    return;
  }
  state.activities
    .slice()
    .sort((a, b) => a.time.localeCompare(b.time))
    .forEach((activity) => {
      const card = document.createElement("article");
      card.className = `activity-card${activity.active ? "" : " inactive-card"}`;
      card.append(createActivityMain(activity, false));
      const actions = document.createElement("div");
      actions.className = "activity-actions";
      actions.append(
        createActionButton("Ubah", "edit", activity.id),
        createActionButton("Hapus", "delete", activity.id, true),
      );
      card.append(actions);
      container.append(card);
    });
}

function createActivityMain(activity, includeRepeat) {
  const main = document.createElement("div");
  main.className = "activity-main";
  const title = document.createElement("p");
  title.className = "activity-title";
  title.textContent = activity.title;
  const meta = document.createElement("div");
  meta.className = "activity-meta";
  const time = document.createElement("span");
  time.className = "time-tag";
  time.textContent = activity.time;
  meta.append(time);
  if (activity.category) {
    const category = document.createElement("span");
    category.className = "category-tag";
    category.textContent = activity.category;
    meta.append(category);
  }
  if (includeRepeat) {
    const repeat = document.createElement("span");
    repeat.textContent = recurrenceLabel(activity);
    meta.append(repeat);
  }
  const reminder = document.createElement("span");
  reminder.textContent = `Pengingat ${activity.reminderMinutes} menit`;
  meta.append(reminder);
  main.append(title, meta);
  return main;
}

function recurrenceLabel(activity) {
  if (activity.recurrence === "once") return `Satu kali · ${activity.startDate}`;
  if (activity.recurrence === "daily") return `Harian · mulai ${activity.startDate}`;
  if (activity.recurrence === "monthly") {
    return `Bulanan · tanggal ${activity.dayOfMonth}`;
  }
  const days = (activity.weekdays || [])
    .map((day) => dayLabels[day - 1])
    .join(", ");
  return `Mingguan · ${days}`;
}

function createActionButton(label, action, id, danger = false) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `icon-button${danger ? " danger" : ""}`;
  button.dataset.action = action;
  button.dataset.id = id;
  button.textContent = label;
  return button;
}

function createEmptyState(message) {
  const empty = document.createElement("div");
  empty.className = "empty-state";
  empty.textContent = message;
  return empty;
}

function setSyncStatus(message) {
  document.querySelector("#sync-label").textContent = message;
}

function showMessage(message) {
  const element = document.querySelector("#app-message");
  element.className = "app-message";
  element.textContent = message;
}

function showError(error) {
  const element = document.querySelector("#app-message");
  element.className = "app-message is-error";
  element.textContent = error.message || "Terjadi kesalahan.";
  setSyncStatus("Gagal tersinkron");
}

function clearMessage() {
  const element = document.querySelector("#app-message");
  element.className = "app-message";
  element.textContent = "";
}

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./service-worker.js").catch((error) => {
      console.error("Service worker tidak dapat didaftarkan:", error);
    });
  }
}
