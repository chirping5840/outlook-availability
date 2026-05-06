/* =====================================================================
 * CONFIG — set these before deploying
 * ===================================================================== */

// Application (client) ID from your Azure AD app registration
const CLIENT_ID = "2c0383d0-39d9-4a01-8c6e-487037cbfcce";

// Use "common" for any work/school account, "organizations" for any work account,
// or your tenant ID (GUID) to lock to one organization.
const TENANT = "c54b4401-f6b5-4b8f-b5eb-497292265d18";

const SCOPES = ["Calendars.Read", "MailboxSettings.Read"];

/* =====================================================================
 * STATE
 * ===================================================================== */

let msalInstance = null;
let accessToken = null;
let selectedDates = new Set();             // Set of "YYYY-MM-DD"
let calendarMonth = null;                  // Date pinned to 1st of displayed month
const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

/* =====================================================================
 * INIT
 * ===================================================================== */

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage" }
  });

  document.getElementById("signin-btn").addEventListener("click", signIn);
  document.getElementById("generate-btn").addEventListener("click", () => generate(true));
  document.getElementById("copy-btn").addEventListener("click", copyToClipboard);

  calendarMonth = startOfMonth(new Date());
  renderCalendar();

  trySilentSignIn();
});

/* =====================================================================
 * AUTH
 * ===================================================================== */

async function trySilentSignIn() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return;
  try {
    const result = await msalInstance.acquireTokenSilent({
      account: accounts[0],
      scopes: SCOPES
    });
    accessToken = result.accessToken;
    showMain();
  } catch (e) {
    // User needs to interactively sign in
    console.log("Silent sign-in failed; waiting for click.", e);
  }
}

async function signIn() {
  setStatus("auth-status", "Signing in...");
  try {
    const result = await msalInstance.loginPopup({ scopes: SCOPES });
    accessToken = result.accessToken;
    showMain();
  } catch (e) {
    setStatus("auth-status", "Sign-in failed: " + e.message, "error");
  }
}

function showMain() {
  document.getElementById("auth-section").classList.add("hidden");
  document.getElementById("main-section").classList.remove("hidden");
}

/* =====================================================================
 * DATE PICKER
 * ===================================================================== */

function renderCalendar() {
  const grid = document.getElementById("calendar-grid");
  grid.innerHTML = "";

  // Header with prev / next month
  const header = document.createElement("div");
  header.className = "cal-header-bar";
  header.innerHTML = `
    <button type="button" id="prev-month" aria-label="Previous month">‹</button>
    <span class="month-label">${calendarMonth.toLocaleDateString(undefined, { month: "long", year: "numeric" })}</span>
    <button type="button" id="next-month" aria-label="Next month">›</button>
  `;
  grid.appendChild(header);

  // Day-of-week labels (Sun first to match Outlook default)
  ["S", "M", "T", "W", "T", "F", "S"].forEach(d => {
    const el = document.createElement("div");
    el.className = "cal-day label";
    el.textContent = d;
    grid.appendChild(el);
  });

  const year = calendarMonth.getFullYear();
  const month = calendarMonth.getMonth();
  const firstWeekday = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const today = stripTime(new Date());

  // Empty leading cells
  for (let i = 0; i < firstWeekday; i++) {
    const el = document.createElement("div");
    el.className = "cal-day empty";
    grid.appendChild(el);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month, d);
    const key = isoDateKey(date);
    const el = document.createElement("div");
    el.className = "cal-day";
    el.textContent = d;
    if (date.getTime() === today.getTime()) el.classList.add("today");
    if (date < today) el.classList.add("past");
    if (selectedDates.has(key)) el.classList.add("selected");
    el.addEventListener("click", () => {
      if (selectedDates.has(key)) selectedDates.delete(key);
      else selectedDates.add(key);
      renderCalendar();
      updateSummary();
    });
    grid.appendChild(el);
  }

  document.getElementById("prev-month").addEventListener("click", () => {
    calendarMonth = new Date(year, month - 1, 1);
    renderCalendar();
  });
  document.getElementById("next-month").addEventListener("click", () => {
    calendarMonth = new Date(year, month + 1, 1);
    renderCalendar();
  });
}

function updateSummary() {
  const el = document.getElementById("selected-summary");
  if (selectedDates.size === 0) {
    el.textContent = "No days selected";
  } else {
    el.textContent = `${selectedDates.size} day${selectedDates.size === 1 ? "" : "s"} selected`;
  }
}

/* =====================================================================
 * GRAPH
 * ===================================================================== */

async function graphFetch(path, extraHeaders = {}) {
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, {
    headers: {
      Authorization: "Bearer " + accessToken,
      ...extraHeaders
    }
  });
  if (!r.ok) {
    const body = await r.text();
    throw new Error(`Graph ${r.status}: ${body}`);
  }
  return r.json();
}

async function getWorkingHours() {
  try {
    const settings = await graphFetch("/me/mailboxSettings/workingHours");
    if (!settings || !settings.daysOfWeek || !settings.startTime || !settings.endTime) return null;
    return settings;
  } catch (e) {
    console.warn("Could not fetch working hours:", e);
    return null;
  }
}

async function getCalendarView(startISO, endISO) {
  // Asking Graph to return event times in the user's local time zone makes
  // free-time math simple — the strings parse as local with `new Date()`.
  const path =
    `/me/calendarView` +
    `?startDateTime=${encodeURIComponent(startISO)}` +
    `&endDateTime=${encodeURIComponent(endISO)}` +
    `&$select=start,end,subject,showAs,isAllDay` +
    `&$top=500` +
    `&$orderby=start/dateTime`;
  const json = await graphFetch(path, {
    Prefer: `outlook.timezone="${userTimeZone}"`
  });
  return json.value || [];
}

/* =====================================================================
 * FREE-TIME COMPUTATION
 * ===================================================================== */

async function generate(insertIntoBody) {
  if (selectedDates.size === 0) {
    setStatus("status", "Pick at least one day.", "error");
    return;
  }
  setStatus("status", "Working...");

  try {
    const useMailboxHours = document.getElementById("use-mailbox-hours").checked;
    const workingHours = useMailboxHours ? await getWorkingHours() : null;
    const fallbackStart = document.getElementById("work-start").value || "09:00";
    const fallbackEnd = document.getElementById("work-end").value || "17:00";
    const minSlotMs = parseInt(document.getElementById("min-slot").value, 10) * 60 * 1000;

    const sorted = [...selectedDates].sort();
    const startBoundary = new Date(sorted[0] + "T00:00:00");
    const endBoundary = new Date(sorted[sorted.length - 1] + "T00:00:00");
    endBoundary.setDate(endBoundary.getDate() + 1);

    const events = await getCalendarView(startBoundary.toISOString(), endBoundary.toISOString());

    // Build one line per selected day
    const lines = [];
    for (const dateKey of sorted) {
      const date = new Date(dateKey + "T00:00:00");

      const window = computeWorkWindow(date, workingHours, fallbackStart, fallbackEnd);
      if (!window) continue; // not a working day

      const dayBusy = events
        .filter(ev => ev.showAs && ev.showAs !== "free")
        .map(ev => ({
          start: parseGraphLocal(ev.start.dateTime),
          end: parseGraphLocal(ev.end.dateTime),
          isAllDay: !!ev.isAllDay
        }))
        .filter(ev => ev.end > window.start && ev.start < window.end)
        .sort((a, b) => a.start - b.start);

      if (dayBusy.some(ev => ev.isAllDay)) continue;

      const slots = subtractBusy(window.start, window.end, dayBusy)
        .filter(s => (s.end - s.start) >= minSlotMs);

      if (slots.length === 0) continue;

      const label = date.toLocaleDateString(undefined, { weekday: "long", month: "short", day: "numeric" });
      const ranges = slots.map(s => `${fmtTime(s.start)}–${fmtTime(s.end)}`).join(", ");
      lines.push({ label, ranges });
    }

    if (lines.length === 0) {
      setStatus("status", "No free time on the selected days within working hours.", "error");
      document.getElementById("preview").textContent = "(no availability)";
      return;
    }

    // Plain-text preview
    const previewText =
      "How are these times for you?\n" +
      lines.map(l => `• ${l.label}: ${l.ranges}`).join("\n");
    document.getElementById("preview").textContent = previewText;

    if (insertIntoBody) {
      // HTML version inserts as a real bullet list at the cursor
      const html =
        "<p>How are these times for you?</p>" +
        "<ul>" +
        lines.map(l => `<li><strong>${escapeHtml(l.label)}:</strong> ${escapeHtml(l.ranges)}</li>`).join("") +
        "</ul>";

      Office.context.mailbox.item.body.setSelectedDataAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            setStatus("status", "Inserted into your message.", "success");
          } else {
            setStatus("status", "Insert failed: " + result.error.message, "error");
          }
        }
      );
    } else {
      setStatus("status", "Generated. Click Insert to add it to the message.", "success");
    }
  } catch (e) {
    console.error(e);
    setStatus("status", "Error: " + e.message, "error");
  }
}

function computeWorkWindow(date, workingHours, fallbackStart, fallbackEnd) {
  if (workingHours) {
    const dayName = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"][date.getDay()];
    if (!workingHours.daysOfWeek.includes(dayName)) return null;
    return {
      start: combineDateAndTime(date, workingHours.startTime),
      end: combineDateAndTime(date, workingHours.endTime)
    };
  }
  return {
    start: combineDateAndTime(date, fallbackStart + ":00"),
    end: combineDateAndTime(date, fallbackEnd + ":00")
  };
}

function subtractBusy(windowStart, windowEnd, busyEvents) {
  const slots = [];
  let cursor = new Date(windowStart);
  for (const ev of busyEvents) {
    const s = ev.start < windowStart ? new Date(windowStart) : new Date(ev.start);
    const e = ev.end > windowEnd ? new Date(windowEnd) : new Date(ev.end);
    if (e <= cursor) continue;
    if (s >= windowEnd) break;
    if (s > cursor) slots.push({ start: new Date(cursor), end: new Date(s) });
    if (e > cursor) cursor = new Date(e);
  }
  if (cursor < windowEnd) slots.push({ start: new Date(cursor), end: new Date(windowEnd) });
  return slots;
}

/* =====================================================================
 * COPY TO CLIPBOARD (fallback if Office.body insert fails)
 * ===================================================================== */

async function copyToClipboard() {
  await generate(false);
  const text = document.getElementById("preview").textContent;
  try {
    await navigator.clipboard.writeText(text);
    setStatus("status", "Copied to clipboard.", "success");
  } catch (e) {
    setStatus("status", "Could not copy: " + e.message, "error");
  }
}

/* =====================================================================
 * UTILITIES
 * ===================================================================== */

function isoDateKey(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function startOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function stripTime(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Graph returns "2026-05-07T10:00:00.0000000" with no offset when Prefer is set.
// Treat the wall-clock time as local.
function parseGraphLocal(dt) {
  const cleaned = dt.split(".")[0];
  return new Date(cleaned);
}

function combineDateAndTime(date, timeStr) {
  // timeStr like "09:00:00" or "09:00:00.0000000"
  const parts = timeStr.split(":");
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  const s = parts[2] ? parseInt(parts[2], 10) : 0;
  const out = new Date(date);
  out.setHours(h, m, s, 0);
  return out;
}

function fmtTime(d) {
  const sel = document.getElementById("time-format")?.value || "auto";
  if (sel === "24h") {
    return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  }
  if (sel === "12h") {
    return d.toLocaleTimeString("en-US", { hour: "numeric", minute: d.getMinutes() ? "2-digit" : undefined });
  }
  return d.toLocaleTimeString(undefined, { hour: "numeric", minute: d.getMinutes() ? "2-digit" : undefined });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;"
  }[c]));
}

function setStatus(id, msg, kind) {
  const el = document.getElementById(id);
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}
