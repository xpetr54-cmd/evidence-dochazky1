/* Evidence pracovní doby — čtrnáctidenní cyklus ranní / odpolední. */
(function (global) {
  var ROLES = [
    { id: "ucitelka", label: "Učitelka" },
    { id: "asistentka", label: "Asistentka" },
    { id: "provoz", label: "Provoz" },
    { id: "administrativa", label: "Administrativa" }
  ];

  var WEEKDAYS = [
    { n: 1, short: "Po", long: "Pondělí" },
    { n: 2, short: "Út", long: "Úterý" },
    { n: 3, short: "St", long: "Středa" },
    { n: 4, short: "Čt", long: "Čtvrtek" },
    { n: 5, short: "Pá", long: "Pátek" }
  ];

  var EXCEPTION_TYPES = [
    { id: "dovolena", code: "D", label: "Dovolená", countsAs: "absence" },
    { id: "nemoc", code: "PN", label: "Nemoc (PN)", countsAs: "absence" },
    { id: "ocr", code: "OČR", label: "OČR / ošetřovné", countsAs: "absence" },
    { id: "svatek", code: "SV", label: "Svátek", countsAs: "absence" },
    { id: "md", code: "MD", label: "Mateřská / rodičovská", countsAs: "absence" },
    { id: "prescas", code: "", label: "Úprava času / přesčas", countsAs: "worked" },
    { id: "jine", code: "", label: "Jiné", countsAs: "worked" }
  ];

  var STATUS = {
    draft: "Rozepsáno",
    submitted: "Odesláno",
    approved: "Schváleno",
    returned: "Vráceno k úpravě"
  };

  var ROLE_DEFAULTS = {
    ucitelka: {
      weeklyHours: 40,
      primaWeekly: 31,
      ranni: slotPedFri(
        ["07:00", "12:30", "12:30", "15:00"],
        ["07:00", "12:15", "12:15", "15:00"]
      ),
      odpoledni: slotPedFri(
        ["09:45", "16:45", "08:45", "09:45"],
        ["10:00", "16:45", "08:45", "10:00"]
      )
    },
    asistentka: {
      weeklyHours: 40,
      primaWeekly: 31,
      ranni: slotPedFri(
        ["07:00", "12:30", "12:30", "15:00"],
        ["07:00", "12:15", "12:15", "15:00"]
      ),
      odpoledni: slotPedFri(
        ["09:45", "16:45", "08:45", "09:45"],
        ["10:00", "16:45", "08:45", "10:00"]
      )
    },
    provoz: {
      weeklyHours: 40,
      primaWeekly: 0,
      ranni: slot("07:15", "15:15", 0),
      odpoledni: slot("07:30", "15:30", 0)
    },
    administrativa: {
      weeklyHours: 40,
      primaWeekly: 0,
      ranni: slot("07:00", "15:30", 30),
      odpoledni: slot("07:00", "15:30", 30)
    }
  };

  function isPedRole(role) {
    return role === "ucitelka" || role === "asistentka";
  }

  function isPedPerson(person) {
    if (!person) return false;
    if (person.formType === "provoz") return false;
    if (person.formType === "pedagog" || person.formType === "extra") return true;
    return isPedRole(person.role);
  }

  function emptyPedSlot() {
    return {
      n1Start: "", n1End: "",
      primaStart: "", primaEnd: "",
      n2Start: "", n2End: "",
      p2Start: "", p2End: "",
      n3Start: "", n3End: ""
    };
  }

  function emptyProvozSlot() {
    return { start: "", end: "", breakStart: "", breakEnd: "", breakMin: 30 };
  }

  function pedSlot(primaStart, primaEnd, neprimaStart, neprimaEnd) {
    var slot = emptyPedSlot();
    slot.primaStart = primaStart || "";
    slot.primaEnd = primaEnd || "";
    if (neprimaStart && primaStart && neprimaStart < primaStart) {
      slot.n1Start = neprimaStart || "";
      slot.n1End = neprimaEnd || "";
    } else {
      slot.n2Start = neprimaStart || "";
      slot.n2End = neprimaEnd || "";
    }
    return slot;
  }

  function slot(start, end, breakMin) {
    var days = {};
    WEEKDAYS.forEach(function (d) {
      days[d.n] = { start: start, end: end, breakMin: breakMin, breakStart: "", breakEnd: "" };
    });
    return days;
  }

  function slotPed(primaStart, primaEnd, neprimaStart, neprimaEnd) {
    var days = {};
    WEEKDAYS.forEach(function (d) {
      days[d.n] = pedSlot(primaStart, primaEnd, neprimaStart, neprimaEnd);
    });
    return days;
  }

  function slotPedFri(week, friday) {
    var days = slotPed(week[0], week[1], week[2], week[3]);
    days[5] = pedSlot(friday[0], friday[1], friday[2], friday[3]);
    return days;
  }

  function roundHours(n) {
    return Math.round(n * 100) / 100;
  }

  function earlier(a, b) {
    if (!a) return b;
    if (!b) return a;
    return a < b ? a : b;
  }

  function later(a, b) {
    if (!a) return b;
    if (!b) return a;
    return a > b ? a : b;
  }

  function pedIntervals(slot) {
    var n = slot || {};
    return [
      { kind: "neprima", start: n.n1Start || "", end: n.n1End || "" },
      { kind: "prima", start: n.primaStart || "", end: n.primaEnd || "" },
      { kind: "neprima", start: n.n2Start || "", end: n.n2End || "" },
      { kind: "prima", start: n.p2Start || "", end: n.p2End || "" },
      { kind: "neprima", start: n.n3Start || "", end: n.n3End || "" }
    ].filter(function (b) { return b.start && b.end; });
  }

  function normalizeSlot(slot, ped) {
    if (ped) {
      if (slot && (slot.n1Start || slot.n2Start || slot.primaStart || slot.p2Start || slot.n3Start)) {
        return Object.assign(emptyPedSlot(), {
          n1Start: slot.n1Start || "",
          n1End: slot.n1End || "",
          primaStart: slot.primaStart || "",
          primaEnd: slot.primaEnd || "",
          n2Start: slot.n2Start || "",
          n2End: slot.n2End || "",
          p2Start: slot.p2Start || "",
          p2End: slot.p2End || "",
          n3Start: slot.n3Start || "",
          n3End: slot.n3End || ""
        });
      }
      if (slot && (slot.primaStart || slot.neprimaStart || slot.primaEnd || slot.neprimaEnd)) {
        return pedSlot(slot.primaStart, slot.primaEnd, slot.neprimaStart, slot.neprimaEnd);
      }
      if (slot && slot.blocks && slot.blocks.length) {
        var converted = emptyPedSlot();
        var primas = slot.blocks.filter(function (b) { return b.kind === "prima" && b.start && b.end; });
        var neprimas = slot.blocks.filter(function (b) { return b.kind === "neprima" && b.start && b.end; });
        if (primas[0]) { converted.primaStart = primas[0].start; converted.primaEnd = primas[0].end; }
        if (primas[1]) { converted.p2Start = primas[1].start; converted.p2End = primas[1].end; }
        if (neprimas[0] && primas[0] && neprimas[0].start < primas[0].start) {
          converted.n1Start = neprimas[0].start; converted.n1End = neprimas[0].end;
          if (neprimas[1]) { converted.n2Start = neprimas[1].start; converted.n2End = neprimas[1].end; }
        } else if (neprimas[0]) {
          converted.n2Start = neprimas[0].start; converted.n2End = neprimas[0].end;
        }
        return converted;
      }
      if (slot && (slot.start || slot.end)) {
        return pedSlot(slot.start, slot.end, "", "");
      }
      return emptyPedSlot();
    }
    var breakStart = slot && slot.breakStart || "";
    var breakEnd = slot && slot.breakEnd || "";
    var breakMin = slot && slot.breakMin != null ? Number(slot.breakMin) : 0;
    if (breakStart && breakEnd) breakMin = Math.round(hoursBetween(breakStart, breakEnd, 0) * 60);
    return {
      start: slot && slot.start || "",
      end: slot && slot.end || "",
      breakStart: breakStart,
      breakEnd: breakEnd,
      breakMin: breakMin
    };
  }

  function analyzeSlot(slot, ped) {
    var n = normalizeSlot(slot, ped);
    if (ped) {
      var parts = pedIntervals(n);
      var hoursPrima = 0;
      var hoursNeprima = 0;
      var start = "";
      var end = "";
      parts.forEach(function (b) {
        var h = hoursBetween(b.start, b.end, 0);
        if (b.kind === "prima") hoursPrima += h;
        else hoursNeprima += h;
        start = earlier(start, b.start);
        end = later(end, b.end);
      });
      var firstP = parts.find(function (b) { return b.kind === "prima"; }) || { start: "", end: "" };
      var firstN = parts.find(function (b) { return b.kind === "neprima"; }) || { start: "", end: "" };
      return {
        ped: true,
        n1Start: n.n1Start, n1End: n.n1End,
        primaStart: firstP.start, primaEnd: firstP.end,
        n2Start: n.n2Start, n2End: n.n2End,
        p2Start: n.p2Start, p2End: n.p2End,
        n3Start: n.n3Start, n3End: n.n3End,
        neprimaStart: firstN.start, neprimaEnd: firstN.end,
        start: start,
        end: end,
        breakMin: 0,
        hoursPrima: roundHours(hoursPrima),
        hoursNeprima: roundHours(hoursNeprima),
        hours: roundHours(hoursPrima + hoursNeprima)
      };
    }
    var br = n.breakStart && n.breakEnd ? hoursBetween(n.breakStart, n.breakEnd, 0) : 0;
    var worked = hoursBetween(n.start, n.end, br ? Math.round(br * 60) : n.breakMin);
    return {
      ped: false,
      primaStart: "",
      primaEnd: "",
      neprimaStart: "",
      neprimaEnd: "",
      start: n.start,
      end: n.end,
      breakStart: n.breakStart,
      breakEnd: n.breakEnd,
      breakMin: n.breakMin,
      hoursPrima: 0,
      hoursNeprima: 0,
      hours: worked
    };
  }

  function formatBlocks(info) {
    if (!info || !info.ped) return "";
    return pedIntervals(info).map(function (b) {
      return b.start + "–" + b.end + " " + (b.kind === "prima" ? "P" : "N");
    }).join(" · ");
  }

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function uid() {
    if (global.crypto && crypto.randomUUID) return crypto.randomUUID();
    return "id-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
  }

  function pad(n) {
    return String(n).padStart(2, "0");
  }

  function parseISODate(s) {
    var p = String(s).split("-").map(Number);
    return new Date(p[0], p[1] - 1, p[2]);
  }

  function formatISODate(d) {
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }

  function addDays(d, n) {
    var x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
    return x;
  }

  function isoWeekParts(d) {
    var date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - day);
    var year = date.getUTCFullYear();
    var yearStart = new Date(Date.UTC(year, 0, 1));
    var week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
    return { year: year, week: week };
  }

  function isoWeekKey(d) {
    var p = isoWeekParts(d);
    return p.year + "-W" + pad(p.week);
  }

  function mondayOfISOWeek(weekKey) {
    var m = /^(\d{4})-W(\d{2})$/.exec(weekKey);
    var year = Number(m[1]);
    var week = Number(m[2]);
    var jan4 = new Date(year, 0, 4);
    var day = jan4.getDay() || 7;
    var mondayWeek1 = addDays(jan4, 1 - day);
    return addDays(mondayWeek1, (week - 1) * 7);
  }

  function weekLabel(weekKey) {
    var mon = mondayOfISOWeek(weekKey);
    var fri = addDays(mon, 4);
    return "Týden " + weekKey.replace("-W", " / ") + "  (" + formatCzechDate(mon) + " – " + formatCzechDate(fri) + ")";
  }

  function formatCzechDate(d) {
    return d.getDate() + ". " + (d.getMonth() + 1) + ".";
  }

  function formatCzechDateFull(d) {
    return d.getDate() + ". " + (d.getMonth() + 1) + ". " + d.getFullYear();
  }

  function easterSunday(year) {
    var a = year % 19;
    var b = Math.floor(year / 100);
    var c = year % 100;
    var d = Math.floor(b / 4);
    var e = b % 4;
    var f = Math.floor((b + 8) / 25);
    var g = Math.floor((b - f + 1) / 3);
    var h = (19 * a + b - d - g + 15) % 30;
    var i = Math.floor(c / 4);
    var k = c % 4;
    var l = (32 + 2 * e + 2 * i - h - k) % 7;
    var m = Math.floor((a + 11 * h + 22 * l) / 451);
    var month = Math.floor((h + l - 7 * m + 114) / 31);
    var day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, month - 1, day);
  }

  function czechHolidays(year) {
    var easter = easterSunday(year);
    var list = [
      [year + "-01-01", "Nový rok"],
      [formatISODate(addDays(easter, -2)), "Velký pátek"],
      [formatISODate(addDays(easter, 1)), "Velikonoční pondělí"],
      [year + "-05-01", "Svátek práce"],
      [year + "-05-08", "Den vítězství"],
      [year + "-07-05", "Cyril a Metoděj"],
      [year + "-07-06", "Den upálení mistra Jana Husa"],
      [year + "-09-28", "Den české státnosti"],
      [year + "-10-28", "Den vzniku samostatného československého státu"],
      [year + "-11-17", "Den boje za svobodu a demokracii"],
      [year + "-12-24", "Štědrý den"],
      [year + "-12-25", "1. svátek vánoční"],
      [year + "-12-26", "2. svátek vánoční"]
    ];
    var map = {};
    list.forEach(function (row) {
      map[row[0]] = row[1];
    });
    return map;
  }

  function hoursBetween(start, end, breakMin) {
    if (!start || !end) return 0;
    var s = start.split(":").map(Number);
    var e = end.split(":").map(Number);
    var mins = e[0] * 60 + e[1] - (s[0] * 60 + s[1]) - (Number(breakMin) || 0);
    return Math.max(0, Math.round((mins / 60) * 100) / 100);
  }

  function formatHours(n) {
    if (n == null || Number.isNaN(n)) return "0";
    var r = Math.round(n * 100) / 100;
    return String(r).replace(".", ",");
  }

  function monthName(month) {
    return ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec"][month - 1];
  }

  function daysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
  }

  function weeksInMonth(year, month) {
    var holidays = czechHolidays(year);
    var seen = [];
    var map = {};
    var last = daysInMonth(year, month);
    for (var day = 1; day <= last; day++) {
      var d = new Date(year, month - 1, day);
      var wd = d.getDay();
      if (wd === 0 || wd === 6) continue;
      var iso = formatISODate(d);
      if (holidays[iso]) continue;
      var key = isoWeekKey(d);
      if (!map[key]) {
        map[key] = { key: key, dates: [] };
        seen.push(map[key]);
      }
      map[key].dates.push(iso);
    }
    return seen;
  }

  function weekParityFromCycle(cycleStart, weekKey) {
    if (!cycleStart) return 0;
    var startMon = mondayOfISOWeek(isoWeekKey(parseISODate(cycleStart)));
    var weekMon = mondayOfISOWeek(weekKey);
    var diff = Math.round((weekMon - startMon) / 86400000 / 7);
    return ((diff % 2) + 2) % 2;
  }

  function suggestedShift(cycleStart, weekKey) {
    return weekParityFromCycle(cycleStart, weekKey) === 0 ? "ranni" : "odpoledni";
  }

  function emptyTemplates(role, formType) {
    var ped = formType ? formType !== "provoz" : isPedRole(role);
    var t = { ranni: {}, odpoledni: {} };
    WEEKDAYS.forEach(function (d) {
      t.ranni[d.n] = ped ? emptyPedSlot() : emptyProvozSlot();
      t.odpoledni[d.n] = ped ? emptyPedSlot() : emptyProvozSlot();
    });
    return t;
  }

  function templatesForRole(role) {
    var def = ROLE_DEFAULTS[role] || ROLE_DEFAULTS.ucitelka;
    return { ranni: clone(def.ranni), odpoledni: clone(def.odpoledni) };
  }

  function emptyState() {
    return {
      version: 1,
      settings: {
        orgName: "Mateřská škola",
        schoolYear: "2026/2027",
        cycleStart: "2026-09-01",
        weeklyHoursFull: 40,
        adminEmails: "",
        leadershipEmails: ""
      },
      people: [],
      timesheets: [],
      accounts: []
    };
  }

  function demoState() {
    var state = emptyState();
    state.settings.orgName = "Mateřská škola (ukázka)";
    var defs = [
      ["Jana Nováková", "ucitelka", "jana.novakova@skolka.cz"],
      ["Petra Svobodová", "ucitelka", "petra.svobodova@skolka.cz"],
      ["Eva Dvořáková", "asistentka", "eva.dvorakova@skolka.cz"],
      ["Martin Černý", "provoz", "martin.cerny@skolka.cz"],
      ["Alena Veselá", "administrativa", "alena.vesela@skolka.cz"]
    ];
    state.people = defs.map(function (row, i) {
      var role = row[1];
      var def = ROLE_DEFAULTS[role];
      return {
        id: "demo-" + (i + 1),
        name: row[0],
        role: role,
        email: row[2],
        workplace: "Wolkerova",
        personalNumber: "",
        fte: 1,
        weeklyHours: def.weeklyHours,
        primaWeekly: def.primaWeekly || 0,
        korekcePrima: 0,
        active: true,
        templates: templatesForRole(role)
      };
    });
    return state;
  }

  function getPerson(state, id) {
    return (state.people || []).find(function (p) { return p.id === id; }) || null;
  }

  function timesheetKey(personId, year, month) {
    return personId + ":" + year + "-" + pad(month);
  }

  function getTimesheet(state, personId, year, month) {
    return (state.timesheets || []).find(function (t) {
      return t.personId === personId && t.year === year && t.month === month;
    }) || null;
  }

  function ensureTimesheet(state, personId, year, month) {
    var existing = getTimesheet(state, personId, year, month);
    if (existing) return existing;
    var weeks = weeksInMonth(year, month);
    var weekShifts = {};
    var person = getPerson(state, personId);
    weeks.forEach(function (w) {
      if (person && person.alternatesShifts === false) {
        weekShifts[w.key] = "ranni";
      } else {
        weekShifts[w.key] = suggestedShift(state.settings.cycleStart, w.key);
      }
    });
    var sheet = {
      id: uid(),
      personId: personId,
      year: year,
      month: month,
      status: "draft",
      weekShifts: weekShifts,
      exceptions: {},
      submittedAt: null,
      approvedAt: null,
      approvedBy: null,
      returnComment: ""
    };
    state.timesheets.push(sheet);
    return sheet;
  }

  function plannedSlot(person, shift, weekday) {
    if (!person || !person.templates || !person.templates[shift]) return null;
    var days = person.templates[shift];
    return days[weekday] || days[String(weekday)] || null;
  }

  function clearDayTimes(row) {
    row.start = "";
    row.end = "";
    row.breakMin = 0;
    row.breakStart = "";
    row.breakEnd = "";
    row.primaStart = "";
    row.primaEnd = "";
    row.neprimaStart = "";
    row.neprimaEnd = "";
    row.schema = "";
    row.hoursWorked = 0;
    row.hoursPrima = 0;
    row.hoursNeprima = 0;
  }

  function personFond(person) {
    return roundHours((Number(person.weeklyHours) || 40) * (Number(person.fte) || 1));
  }

  function personPrimaBase(person) {
    var raw;
    if (person.primaWeekly != null && person.primaWeekly !== "") raw = Number(person.primaWeekly) || 0;
    else {
      var def = ROLE_DEFAULTS[person.role];
      raw = def && def.primaWeekly ? def.primaWeekly : 0;
    }
    var fte = Number(person.fte) || 1;
    if (raw > 0 && fte > 0 && Math.abs(fte - 1) > 0.001) {
      var scaled = roundHours(31 * fte);
      if (Math.abs(raw - scaled) <= 0.05) return 31;
    }
    return raw;
  }

  function personPrimaWeekly(person) {
    var fte = Number(person.fte) || 1;
    return roundHours(Math.max(0, personPrimaBase(person) * fte - personKorekce(person)));
  }

  function personKorekce(person) {
    return Number(person.korekcePrima) || 0;
  }

  function personNeprimaWeekly(person) {
    return roundHours(personFond(person) - personPrimaWeekly(person));
  }

  function exceptionMeta(type) {
    return EXCEPTION_TYPES.find(function (x) { return x.id === type; }) || null;
  }

  function buildMonth(state, person, year, month) {
    var sheet = getTimesheet(state, person.id, year, month) || ensureTimesheet(state, person.id, year, month);
    var holidays = Object.assign({}, czechHolidays(year));
    if (month === 1) Object.assign(holidays, czechHolidays(year - 1));
    if (month === 12) Object.assign(holidays, czechHolidays(year + 1));
    var weeks = weeksInMonth(year, month).map(function (w) {
      return {
        key: w.key,
        label: weekLabel(w.key),
        dates: w.dates,
        shift: sheet.weekShifts[w.key] || suggestedShift(state.settings.cycleStart, w.key)
      };
    });
    var days = [];
    var last = daysInMonth(year, month);
    for (var day = 1; day <= last; day++) {
      var d = new Date(year, month - 1, day);
      var wd = d.getDay();
      var iso = formatISODate(d);
      if (wd === 0 || wd === 6) continue;
      var holiday = holidays[iso] || null;
      var weekKey = isoWeekKey(d);
      var shift = sheet.weekShifts[weekKey] || suggestedShift(state.settings.cycleStart, weekKey);
      if (person.alternatesShifts === false) shift = "ranni";
      var ped = isPedPerson(person);
      var planned = holiday ? null : plannedSlot(person, shift, wd);
      var info = analyzeSlot(planned, ped);
      var plannedHours = planned ? info.hours : 0;
      var ex = sheet.exceptions[iso] || null;
      var row = {
        date: iso,
        dateObj: d,
        weekday: wd,
        weekdayShort: (WEEKDAYS.find(function (x) { return x.n === wd; }) || {}).short || "",
        weekKey: weekKey,
        holiday: holiday,
        shift: shift,
        planned: planned,
        plannedHours: plannedHours,
        plannedPrima: planned ? info.hoursPrima : 0,
        plannedNeprima: planned ? info.hoursNeprima : 0,
        exception: ex,
        start: info.start || "",
        end: info.end || "",
        breakMin: info.breakMin || 0,
        breakStart: info.breakStart || "",
        breakEnd: info.breakEnd || "",
        primaStart: info.primaStart || "",
        primaEnd: info.primaEnd || "",
        neprimaStart: info.neprimaStart || "",
        neprimaEnd: info.neprimaEnd || "",
        schema: formatBlocks(info),
        ped: ped,
        kind: holiday ? "svatek" : "prace",
        hoursWorked: 0,
        hoursPrima: 0,
        hoursNeprima: 0,
        hoursAbsence: 0,
        hoursOvertime: 0
      };
      if (holiday) {
        clearDayTimes(row);
        row.kind = "svatek";
      } else if (ex) {
        var meta = exceptionMeta(ex.type);
        if (meta && meta.countsAs === "absence") {
          row.kind = ex.type;
          row.hoursAbsence = plannedHours;
          clearDayTimes(row);
        } else {
          row.kind = ex.type || "jine";
          if (ped) {
            row.primaStart = ex.primaStart || row.primaStart;
            row.primaEnd = ex.primaEnd || row.primaEnd;
            row.neprimaStart = ex.neprimaStart || row.neprimaStart;
            row.neprimaEnd = ex.neprimaEnd || row.neprimaEnd;
            var adj = analyzeSlot({
              primaStart: row.primaStart,
              primaEnd: row.primaEnd,
              neprimaStart: row.neprimaStart,
              neprimaEnd: row.neprimaEnd
            }, true);
            row.start = adj.start;
            row.end = adj.end;
            row.hoursWorked = adj.hours;
            row.hoursPrima = adj.hoursPrima;
            row.hoursNeprima = adj.hoursNeprima;
            row.schema = formatBlocks(adj);
          } else {
            row.start = ex.start || row.start;
            row.end = ex.end || row.end;
            if (ex.breakStart != null) row.breakStart = ex.breakStart;
            if (ex.breakEnd != null) row.breakEnd = ex.breakEnd;
            row.breakMin = ex.breakMin != null ? Number(ex.breakMin) : row.breakMin;
            if (row.breakStart && row.breakEnd) {
              row.breakMin = Math.round(hoursBetween(row.breakStart, row.breakEnd, 0) * 60);
            }
            row.hoursWorked = hoursBetween(row.start, row.end, row.breakMin);
          }
          if (ex.type === "prescas") {
            row.hoursOvertime = Math.max(0, row.hoursWorked - plannedHours);
          }
        }
      } else {
        row.hoursWorked = plannedHours;
        row.hoursPrima = info.hoursPrima;
        row.hoursNeprima = info.hoursNeprima;
      }
      days.push(row);
    }

    var totals = {
      fond: 0,
      fondPrima: 0,
      fondNeprima: 0,
      worked: 0,
      prima: 0,
      neprima: 0,
      absence: 0,
      dovolena: 0,
      nemoc: 0,
      ocr: 0,
      md: 0,
      prescas: 0,
      svatek: 0
    };
    days.forEach(function (row) {
      totals.fond += row.plannedHours;
      totals.fondPrima += row.plannedPrima;
      totals.fondNeprima += row.plannedNeprima;
      totals.worked += row.hoursWorked;
      totals.prima += row.hoursPrima;
      totals.neprima += row.hoursNeprima;
      totals.absence += row.hoursAbsence;
      if (row.kind === "dovolena") totals.dovolena += row.hoursAbsence;
      if (row.kind === "nemoc") totals.nemoc += row.hoursAbsence;
      if (row.kind === "ocr") totals.ocr += row.hoursAbsence;
      if (row.kind === "md") totals.md += row.hoursAbsence;
      if (row.kind === "svatek") totals.svatek += 0;
      totals.prescas += row.hoursOvertime;
    });
    Object.keys(totals).forEach(function (k) {
      totals[k] = Math.round(totals[k] * 100) / 100;
    });
    totals.expected = Math.round((totals.fond) * 100) / 100;
    totals.difference = Math.round((totals.worked + totals.absence - totals.fond) * 100) / 100;

    return { sheet: sheet, weeks: weeks, days: days, totals: totals };
  }

  function roleLabel(id) {
    var r = ROLES.find(function (x) { return x.id === id; });
    return r ? r.label : id;
  }

  function shiftLabel(id) {
    return id === "odpoledni" ? "Odpolední" : "Ranní";
  }

  function shiftCode(id) {
    return id === "odpoledni" ? "O" : "R";
  }

  function kindLabel(row) {
    if (row.holiday) return "SV";
    var meta = exceptionMeta(row.kind);
    if (meta && meta.code) return meta.code;
    if (meta) return meta.label;
    return "";
  }

  function absenceLabel(row) {
    if (row.holiday) return "SV";
    var meta = exceptionMeta(row.kind);
    return meta && meta.code ? meta.code : "";
  }

  function formTypeLabel(id) {
    if (id === "provoz") return "Provozní";
    if (id === "extra") return "Extra";
    return "Pedagogický";
  }

  var api = {
    ROLES: ROLES,
    WEEKDAYS: WEEKDAYS,
    EXCEPTION_TYPES: EXCEPTION_TYPES,
    STATUS: STATUS,
    ROLE_DEFAULTS: ROLE_DEFAULTS,
    isPedRole: isPedRole,
    isPedPerson: isPedPerson,
    emptyPedSlot: emptyPedSlot,
    emptyProvozSlot: emptyProvozSlot,
    normalizeSlot: normalizeSlot,
    analyzeSlot: analyzeSlot,
    formatBlocks: formatBlocks,
    formTypeLabel: formTypeLabel,
    uid: uid,
    clone: clone,
    parseISODate: parseISODate,
    formatISODate: formatISODate,
    formatCzechDate: formatCzechDate,
    formatCzechDateFull: formatCzechDateFull,
    isoWeekKey: isoWeekKey,
    isoWeekParts: isoWeekParts,
    mondayOfISOWeek: mondayOfISOWeek,
    weekLabel: weekLabel,
    czechHolidays: czechHolidays,
    hoursBetween: hoursBetween,
    formatHours: formatHours,
    monthName: monthName,
    daysInMonth: daysInMonth,
    weeksInMonth: weeksInMonth,
    suggestedShift: suggestedShift,
    emptyTemplates: emptyTemplates,
    templatesForRole: templatesForRole,
    emptyState: emptyState,
    demoState: demoState,
    getPerson: getPerson,
    getTimesheet: getTimesheet,
    ensureTimesheet: ensureTimesheet,
    timesheetKey: timesheetKey,
    buildMonth: buildMonth,
    roleLabel: roleLabel,
    shiftLabel: shiftLabel,
    shiftCode: shiftCode,
    kindLabel: kindLabel,
    absenceLabel: absenceLabel,
    exceptionMeta: exceptionMeta,
    personFond: personFond,
    personPrimaBase: personPrimaBase,
    personPrimaWeekly: personPrimaWeekly,
    personKorekce: personKorekce,
    personNeprimaWeekly: personNeprimaWeekly
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  global.DochazkaModel = api;
})(typeof window !== "undefined" ? window : global);
