(function (root) {
  var XLSX = root.XLSX;
  var SKIP = {
    "zaměstnanci": 1, zamestnanci: 1, parametry: 1, "(zam)": 1, list1: 1,
    "w_provoznýold": 1, w_provoznyold: 1, "m_hošková": 1, m_hoskova: 1
  };
  var ZARAZENI = {
    "učitelka": ["ucitelka", "pedagog"],
    "asist.ped.": ["asistentka", "pedagog"],
    "asistent pedagoga": ["asistentka", "pedagog"],
    "provozní": ["provoz", "provoz"],
    pradlena: ["provoz", "provoz"],
    "administr.": ["administrativa", "provoz"],
    administrativa: ["administrativa", "provoz"],
    "školní asist.": ["asistentka", "extra"],
    "školní asistent": ["asistentka", "extra"],
    šablony: ["ucitelka", "extra"]
  };

  function norm(s) {
    return String(s || "").toLowerCase().replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  }

  function pad(n) {
    return String(n).padStart(2, "0");
  }

  function asTime(v, formatted) {
    if (formatted && /^\d{1,2}:\d{2}/.test(String(formatted))) {
      var fp = String(formatted).split(":");
      return pad(Number(fp[0])) + ":" + pad(Number(fp[1]));
    }
    if (v == null || v === "") return "";
    if (typeof v === "number" && !isNaN(v)) {
      if (v >= 0 && v < 1.5) {
        var mins = Math.round((v % 1) * 24 * 60);
        return pad(Math.floor(mins / 60)) + ":" + pad(mins % 60);
      }
      return "";
    }
    if (v instanceof Date && !isNaN(v.getTime())) {
      return pad(v.getHours()) + ":" + pad(v.getMinutes());
    }
    var s = String(v).trim();
    if (!s || s === "0" || s === "0:00:00") return "";
    var m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return pad(Number(m[1])) + ":" + pad(Number(m[2]));
    return s;
  }

  function cell(ws, row, col) {
    return ws[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })] || null;
  }

  function cellTime(ws, row, col) {
    var x = cell(ws, row, col);
    if (!x) return "";
    return asTime(x.v, x.w);
  }

  function cellText(ws, row, col) {
    var x = cell(ws, row, col);
    if (!x || x.v == null) return "";
    return String(x.v).trim();
  }

  function cellNum(ws, row, col) {
    var x = cell(ws, row, col);
    if (!x || x.v == null || x.v === "") return null;
    var n = Number(x.v);
    return isNaN(n) ? null : n;
  }

  function asHoursDuration(ws, row, col) {
    var x = cell(ws, row, col);
    if (!x || x.v == null || x.v === "") return 0;
    if (typeof x.v === "number" && x.v > 1) return Math.round(x.v * 100) / 100;
    var t = asTime(x.v, x.w);
    if (!t || t.indexOf(":") < 0) return 0;
    var p = t.split(":");
    return Math.round((Number(p[0]) + Number(p[1]) / 60) * 100) / 100;
  }

  function asFte(v) {
    if (v == null || v === "") return 1;
    var n = Number(v);
    if (isNaN(n)) return 1;
    if (n > 1.5) return 1;
    return Math.round(n * 10000) / 10000;
  }

  function asOscis(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number") {
      if (!v) return "";
      if (v === Math.floor(v)) return String(Math.floor(v));
      return String(v).replace(/\.0+$/, "");
    }
    var s = String(v).trim();
    if (s === "0" || s === "0.0") return "";
    if (s.slice(-2) === ".0") s = s.slice(0, -2);
    return s;
  }

  function emptyPed() {
    return {
      n1Start: "", n1End: "", primaStart: "", primaEnd: "",
      n2Start: "", n2End: "", p2Start: "", p2End: "", n3Start: "", n3End: ""
    };
  }

  function emptyProvoz() {
    return { start: "", end: "", breakStart: "", breakEnd: "", breakMin: 0 };
  }

  function slotFilled(slot) {
    return Object.keys(slot).some(function (k) {
      return k !== "breakMin" && String(slot[k] || "").trim();
    });
  }

  function readRow(ws, row, kind) {
    if (kind === "provoz") {
      var start = cellTime(ws, row, 5);
      var end = cellTime(ws, row, 6);
      var brS = cellTime(ws, row, 13);
      var brE = cellTime(ws, row, 14);
      var mins = 0;
      if (brS && brE && brS < brE) {
        var a = brS.split(":").map(Number);
        var b = brE.split(":").map(Number);
        mins = Math.max(0, b[0] * 60 + b[1] - (a[0] * 60 + a[1]));
      }
      return { start: start, end: end, breakStart: brS, breakEnd: brE, breakMin: mins };
    }
    var slot = emptyPed();
    slot.n1Start = cellTime(ws, row, 3);
    slot.n1End = cellTime(ws, row, 4);
    slot.primaStart = cellTime(ws, row, 5);
    slot.primaEnd = cellTime(ws, row, 6);
    slot.n2Start = cellTime(ws, row, 7);
    slot.n2End = cellTime(ws, row, 8);
    slot.p2Start = cellTime(ws, row, 9);
    slot.p2End = cellTime(ws, row, 10);
    slot.n3Start = cellTime(ws, row, 11);
    slot.n3End = cellTime(ws, row, 12);
    return slot;
  }

  function readSchedule(ws, kind) {
    var ranni = {};
    var odpo = {};
    var i;
    for (i = 0; i < 5; i++) ranni[String(i + 1)] = readRow(ws, i + 2, kind);
    for (i = 0; i < 5; i++) odpo[String(i + 1)] = readRow(ws, i + 7, kind);
    var same = JSON.stringify(ranni) === JSON.stringify(odpo);
    var has = Object.keys(ranni).some(function (k) { return slotFilled(ranni[k]); }) ||
      Object.keys(odpo).some(function (k) { return slotFilled(odpo[k]); });
    return { templates: { ranni: ranni, odpoledni: odpo }, alternatesShifts: !same, hasSchedule: has };
  }

  function detectKind(fileName, wb) {
    var n = norm(fileName);
    if (n.indexOf("provoz") >= 0) return "provoz";
    if (n.indexOf("extra") >= 0) return "extra";
    if (n.indexOf("pedagog") >= 0) return "pedagog";
    var names = (wb.SheetNames || []).map(norm);
    if (names.indexOf("zaměstnanci") >= 0 || names.indexOf("zamestnanci") >= 0) return "pedagog";
    return "pedagog";
  }

  function loadSchedules(wb, kind) {
    var out = {};
    (wb.SheetNames || []).forEach(function (name) {
      var n = norm(name);
      if (SKIP[n] || n.indexOf("m_") === 0 || n.indexOf("w_") === 0) return;
      out[n] = readSchedule(wb.Sheets[name], kind);
      out[n].sheet = name;
    });
    return out;
  }

  function readDirectory(wb) {
    var sheetName = wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    var ref = ws && ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { e: { r: 1 } };
    var grouped = [];
    var map = {};
    var r;
    for (r = 2; r <= ref.e.r + 1; r++) {
      var cele = cellText(ws, r, 2);
      var prijmeni = cellText(ws, r, 4);
      var jmeno = cellText(ws, r, 5);
      if (!prijmeni && cele) {
        var parts = cele.split(" ");
        prijmeni = parts[0];
        jmeno = parts.slice(1).join(" ");
      }
      if (!prijmeni && !cele) continue;
      var key = prijmeni + "\t" + jmeno;
      if (!map[key]) {
        map[key] = [];
        grouped.push({ prijmeni: prijmeni, jmeno: jmeno, rows: map[key] });
      }
      map[key].push({
        cele: cele || (prijmeni + " " + jmeno).trim(),
        zarazeni: cellText(ws, r, 3),
        os: asOscis(cell(ws, r, 6) && cell(ws, r, 6).v),
        funkce: cellText(ws, r, 7),
        fte: asFte(cellNum(ws, r, 8)),
        korekce: asHoursDuration(ws, r, 9),
        workplace: cellText(ws, r, 10),
        classroom: cellText(ws, r, 11),
        note: cellText(ws, r, 13)
      });
    }
    return grouped;
  }

  function mapZarazeni(raw) {
    return ZARAZENI[norm(raw)] || ["ucitelka", "pedagog"];
  }

  function personId(os, prijmeni, jmeno) {
    if (os) return "os-" + os;
    return "nm-" + norm(prijmeni + "-" + jmeno).replace(/ /g, "-");
  }

  function emptyTemplates(form) {
    var t = { ranni: {}, odpoledni: {} };
    var d;
    for (d = 1; d <= 5; d++) {
      t.ranni[String(d)] = form === "provoz" ? emptyProvoz() : emptyPed();
      t.odpoledni[String(d)] = form === "provoz" ? emptyProvoz() : emptyPed();
    }
    return t;
  }

  function buildPeople(grouped, pedSched, provSched, extraSched, existing) {
    var byId = {};
    var byName = {};
    (existing || []).forEach(function (p) {
      byId[p.id] = p;
      byName[norm(p.name)] = p;
    });
    var people = grouped.map(function (g) {
      var key = norm(g.prijmeni + " " + g.jmeno);
      var fte = 0;
      var korekce = 0;
      var osReal = "";
      var workplace = "";
      var classroom = "";
      var notes = [];
      var role = "ucitelka";
      var form = "pedagog";
      var funkce = "";
      var contracts = [];
      g.rows.forEach(function (row) {
        var mapped = mapZarazeni(row.zarazeni);
        contracts.push({ zarazeni: row.zarazeni, role: mapped[0], fte: row.fte, personalNumber: row.os });
        fte += row.fte;
        if (row.korekce > korekce) korekce = row.korekce;
        if (row.os && (!osReal || row.os.length <= osReal.length)) osReal = row.os;
        workplace = workplace || row.workplace;
        classroom = classroom || row.classroom;
        if (row.note) notes.push(row.note);
        if (row.funkce) funkce = row.funkce;
        role = mapped[0];
        form = mapped[1];
      });
      if (g.rows.length > 1) form = "extra";
      var sched = extraSched[key] || pedSched[key] || provSched[key];
      if (extraSched[key]) form = "extra";
      else if (provSched[key] && !pedSched[key]) {
        form = "provoz";
        if (role !== "provoz" && role !== "administrativa") role = "provoz";
      }
      if (form === "provoz" && role !== "provoz" && role !== "administrativa") role = "provoz";
      if (form === "pedagog" && role !== "ucitelka" && role !== "asistentka") role = "ucitelka";
      var prev = byName[key] || (osReal ? byId["os-" + osReal] : null);
      if (!sched && prev && prev.templates) {
        sched = { templates: prev.templates, alternatesShifts: prev.alternatesShifts !== false, hasSchedule: prev.hasSchedule };
      }
      var templates = sched ? sched.templates : emptyTemplates(form);
      var has = sched ? !!sched.hasSchedule : false;
      var person = {
        id: personId(osReal, g.prijmeni, g.jmeno),
        name: (g.prijmeni + " " + g.jmeno).trim(),
        lastName: g.prijmeni,
        firstName: g.jmeno,
        role: role,
        formType: form,
        funkce: funkce,
        email: prev && prev.email || "",
        workplace: workplace,
        classroom: classroom,
        personalNumber: osReal,
        fte: Math.round((fte <= 1.05 ? Math.min(fte, 1) : fte) * 10000) / 10000,
        weeklyHours: 40,
        primaWeekly: form === "provoz" ? 0 : 31,
        korekcePrima: korekce,
        alternatesShifts: has ? !!sched.alternatesShifts : true,
        hasSchedule: has,
        note: notes.filter(function (x, i) { return notes.indexOf(x) === i; }).join("; "),
        contracts: contracts.length > 1 ? contracts : [],
        active: prev && prev.active === false ? false : true,
        templates: templates
      };
      return person;
    });
    people.sort(function (a, b) { return a.name.localeCompare(b.name, "cs"); });
    return people;
  }

  function parseBuffers(files, existingPeople) {
    var pedSched = {};
    var provSched = {};
    var extraSched = {};
    var grouped = null;
    files.forEach(function (file) {
      var wb = XLSX.read(file.data, { type: "array", cellDates: true });
      var kind = detectKind(file.name, wb);
      var names = (wb.SheetNames || []).map(norm);
      var hasDir = names.indexOf("zaměstnanci") >= 0 || names.indexOf("zamestnanci") >= 0;
      if (hasDir && kind !== "provoz") {
        if (!grouped || kind === "pedagog") grouped = readDirectory(wb);
      }
      var sched = loadSchedules(wb, kind === "provoz" ? "provoz" : "pedagog");
      if (kind === "provoz") Object.assign(provSched, sched);
      else if (kind === "extra") Object.assign(extraSched, sched);
      else Object.assign(pedSched, sched);
    });
    if (!grouped) throw new Error("Chybí Pedagog.xlsx se seznamem Zaměstnanci.");
    var people = buildPeople(grouped, pedSched, provSched, extraSched, existingPeople);
    return {
      stamp: "excel-" + new Date().toISOString(),
      orgName: "Mateřská škola Olomouc, Wolkerova 34",
      schoolYear: "2026/2027",
      cycleStart: "2026-09-01",
      people: people
    };
  }

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve({ name: file.name, data: new Uint8Array(reader.result) }); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  function parseFileList(fileList, existingPeople) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return Promise.reject(new Error("Vyberte Pedagog.xlsx (případně Provozní.xlsx a Extra.xlsx)."));
    return Promise.all(files.map(readFile)).then(function (bufs) {
      return parseBuffers(bufs, existingPeople);
    });
  }

  var api = { parseBuffers: parseBuffers, parseFileList: parseFileList, asTime: asTime };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.DochazkaExcelImport = api;
})(typeof window !== "undefined" ? window : global);
