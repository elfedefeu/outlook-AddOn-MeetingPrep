/* Meeting Prep - Outlook task pane (Office.js, no dependencies). */
const CHECKS = [
  "Know the goal of this meeting",
  "Read the agenda / related emails",
  "Know who decides what",
  "Prepare my questions or updates",
  "Open documents ready to share",
];

let item;
let state = { subject: "", start: null, end: null, location: "", body: "", people: [], notesKey: "" };

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  item = Office.context.mailbox.item;
  if (!item) { show("Open a calendar event to see its prep."); return; }
  load().catch((e) => show("Could not read this event: " + e.message));
});

// Read-mode properties are plain values; compose-mode ones expose getAsync.
function get(prop) {
  return new Promise((resolve) => {
    if (prop && typeof prop.getAsync === "function") {
      prop.getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null));
    } else resolve(prop ?? null);
  });
}

function bodyText() {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""));
  });
}

async function load() {
  const [subject, start, end, location, req, opt, body] = await Promise.all([
    get(item.subject), get(item.start), get(item.end), get(item.location),
    get(item.requiredAttendees), get(item.optionalAttendees), bodyText(),
  ]);
  const organizer = item.organizer && !item.organizer.getAsync ? item.organizer : null;
  const me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();

  state.subject = subject || "(no subject)";
  state.start = start ? new Date(start) : null;
  state.end = end ? new Date(end) : null;
  state.location = location || "";
  state.body = body || "";

  const seen = new Set();
  const people = [];
  const add = (p, role) => {
    if (!p || !p.emailAddress) return;
    const email = p.emailAddress.toLowerCase();
    if (seen.has(email)) return;
    seen.add(email);
    people.push({ name: p.displayName || email, email, role, isMe: email === me });
  };
  add(organizer, "Organizer");
  (req || []).forEach((p) => add(p, "Required"));
  (opt || []).forEach((p) => add(p, "Optional"));
  state.people = people;
  state.notesKey = "prep_" + hash(state.subject + (state.start ? state.start.toISOString() : ""));

  render(me);
  document.getElementById("msg").hidden = true;
  document.getElementById("app").hidden = false;
  setInterval(() => updateCountdown(), 30000);
}

function render(me) {
  const $ = (id) => document.getElementById(id);
  $("subject").textContent = state.subject;

  if (state.start) {
    const opts = { weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" };
    const mins = state.end ? Math.round((state.end - state.start) / 60000) : null;
    $("when").textContent = state.start.toLocaleString([], opts) + (mins ? ` (${mins} min)` : "");
  }
  updateCountdown();

  const join = findJoinLink(state.location + "\n" + state.body);
  if (join) { $("joinLink").href = join; $("join").hidden = false; }
  $("where").textContent = state.location && !/^https?:/i.test(state.location) ? "Location: " + state.location : "";

  // Agenda
  const lines = extractAgenda(state.body);
  const ul = $("agenda");
  ul.innerHTML = "";
  lines.forEach((l) => { const li = document.createElement("li"); li.textContent = "• " + l; ul.appendChild(li); });
  $("agendaWarn").hidden = lines.length > 0;

  // People
  const myDomain = me.split("@")[1] || "";
  const ext = new Set();
  const pl = $("people");
  pl.innerHTML = "";
  state.people.forEach((p) => {
    const domain = p.email.split("@")[1] || "";
    const external = myDomain && domain !== myDomain;
    if (external && !p.isMe) ext.add(domain);
    const li = document.createElement("li");
    li.className = "person";
    const left = document.createElement("span");
    const n = document.createElement("span"); n.className = "name"; n.textContent = p.name + (p.isMe ? " (you)" : "");
    left.appendChild(n);
    left.appendChild(tag(p.role));
    if (external) left.appendChild(tag("External", "ext"));
    li.appendChild(left);
    if (!p.isMe) {
      const a = document.createElement("a");
      a.textContent = "Past emails"; a.className = "small"; a.target = "_blank"; a.rel = "noopener";
      a.href = searchUrl(p.email);
      li.appendChild(a);
    }
    pl.appendChild(li);
  });
  $("peopleCount").textContent = `(${state.people.length})`;
  if (ext.size) { $("extWarn").hidden = false; $("extWarn").textContent = "External attendees from: " + [...ext].join(", "); }

  $("searchSubject").href = searchUrl(cleanSubject(state.subject));

  // Checklist (persisted)
  const saved = getSetting(state.notesKey + "_c") || [];
  const cl = $("checklist");
  cl.innerHTML = "";
  CHECKS.forEach((c, i) => {
    const li = document.createElement("li");
    const label = document.createElement("label");
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = saved.includes(i);
    cb.addEventListener("change", () => {
      const cur = new Set(getSetting(state.notesKey + "_c") || []);
      cb.checked ? cur.add(i) : cur.delete(i);
      setSetting(state.notesKey + "_c", [...cur]);
    });
    label.appendChild(cb); label.appendChild(document.createTextNode(c));
    li.appendChild(label); cl.appendChild(li);
  });

  // Notes (debounced save)
  const notes = $("notes");
  notes.value = getSetting(state.notesKey + "_n") || "";
  let t;
  notes.addEventListener("input", () => { clearTimeout(t); t = setTimeout(() => setSetting(state.notesKey + "_n", notes.value), 600); });

  $("copy").addEventListener("click", copyBrief);
}

function updateCountdown() {
  const el = document.getElementById("countdown");
  if (!state.start) return;
  const diff = Math.round((state.start - Date.now()) / 60000);
  if (diff > 1440) el.textContent = `In ${Math.round(diff / 1440)} day(s)`;
  else if (diff > 60) el.textContent = `In ${Math.floor(diff / 60)}h ${diff % 60}min`;
  else if (diff > 0) el.textContent = `Starts in ${diff} min`;
  else if (state.end && Date.now() < state.end) el.textContent = "In progress";
  else el.textContent = "Ended";
}

// Agenda = bullet/numbered/"Agenda:" lines. Skips Teams/Zoom boilerplate.
function extractAgenda(text) {
  const out = [];
  let inAgenda = false;
  for (const raw of (text || "").split(/\r?\n/)) {
    const l = raw.trim();
    if (!l) continue;
    if (/^_{5,}|^-{5,}/.test(l) || /microsoft teams|join (the )?meeting|meeting id|passcode|need help\?/i.test(l)) { inAgenda = false; continue; }
    if (/^agenda\b|^ordre du jour|^programme\b/i.test(l)) { inAgenda = true; const rest = l.replace(/^[^:]*:?/, "").trim(); if (rest) out.push(rest); continue; }
    if (/^([-*•●]|\d+[.)])\s+/.test(l)) out.push(l.replace(/^([-*•●]|\d+[.)])\s+/, ""));
    else if (inAgenda && l.length < 140) out.push(l);
  }
  return out.slice(0, 12);
}

function findJoinLink(text) {
  const m = (text || "").match(/https:\/\/(teams\.microsoft\.com\/l\/meetup-join|[\w.-]*zoom\.us\/j|meet\.google\.com)[^\s<>")]+/i);
  return m ? m[0] : null;
}

function searchUrl(q) { return "https://outlook.office.com/mail/search?q=" + encodeURIComponent(q); }
function cleanSubject(s) { return s.replace(/^(re|fw|fwd|tr)\s*:\s*/gi, "").trim(); }
function tag(text, cls) { const s = document.createElement("span"); s.className = "tag " + (cls || ""); s.textContent = text; return s; }
function hash(s) { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return String(h); }

function getSetting(k) { try { return Office.context.roamingSettings.get(k); } catch { return null; } }
function setSetting(k, v) {
  try { Office.context.roamingSettings.set(k, v); Office.context.roamingSettings.saveAsync(); } catch { /* ignore */ }
}

function copyBrief() {
  const agenda = extractAgenda(state.body).map((a) => "- " + a).join("\n") || "(none)";
  const people = state.people.map((p) => `- ${p.name} <${p.email}> [${p.role}]`).join("\n");
  const txt = `${state.subject}\n${document.getElementById("when").textContent}\n\nAgenda:\n${agenda}\n\nPeople:\n${people}\n\nNotes:\n${document.getElementById("notes").value}`;
  const done = () => { const c = document.getElementById("copied"); c.hidden = false; setTimeout(() => (c.hidden = true), 1500); };
  if (navigator.clipboard) navigator.clipboard.writeText(txt).then(done, () => fallbackCopy(txt, done));
  else fallbackCopy(txt, done);
}
function fallbackCopy(txt, done) {
  const ta = document.createElement("textarea"); ta.value = txt; document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); done(); } catch { /* ignore */ } ta.remove();
}

function show(m) { document.getElementById("msg").textContent = m; }
