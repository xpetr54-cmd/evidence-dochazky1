(function () {
  var M = window.DochazkaModel;
  var S = window.DochazkaStorage;
  var A = window.DochazkaAuth;

  var ui = {
    view: "timesheet",
    year: new Date().getFullYear(),
    month: new Date().getMonth() + 1,
    personId: null,
    accountId: null,
    editPersonId: null,
    storage: "local",
    spWeb: "",
    spUser: null,
    busy: false,
    loginFails: 0,
    lockUntil: 0
  };

  var ACCESS_RESET = "zakladni-stav-2026-09-08";
  var state = S.loadLocal() || M.emptyState();
  if (!state.accounts) state.accounts = [];
  resetAccessToBasicIfNeeded();
  restoreSession();

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c];
    });
  }

  function $(id) {
    return document.getElementById(id);
  }

  function flash(msg, isError) {
    var el = $("flash");
    el.hidden = !msg;
    el.textContent = msg || "";
    el.classList.toggle("error", !!isError);
    if (msg) setTimeout(function () { if (el.textContent === msg) el.hidden = true; }, 5000);
  }

  function persist() {
    if (ui.storage === "sharepoint" && ui.spWeb) {
      return S.saveToSharePoint(ui.spWeb, state).then(function () {
        S.saveLocal(state);
      });
    }
    S.saveLocal(state);
    return Promise.resolve();
  }

  function currentAccount() {
    return A.findById(state, ui.accountId);
  }

  function currentPerson() {
    var acc = currentAccount();
    if (!canAdmin()) {
      return acc && acc.personId ? M.getPerson(state, acc.personId) : null;
    }
    return M.getPerson(state, ui.personId) || state.people.filter(function (p) { return p.active; })[0] || null;
  }

  function canAdmin() {
    var acc = currentAccount();
    return !!(acc && acc.active && acc.role === "admin");
  }

  function canLead() {
    return canAdmin();
  }

  function restoreSession() {
    var id = A.loadSession();
    var acc = id ? A.findById(state, id) : null;
    if (!acc || acc.active === false) {
      ui.accountId = null;
      A.clearSession();
      return;
    }
    ui.accountId = acc.id;
    if (acc.personId) ui.personId = acc.personId;
    else if (!ui.personId && state.people[0]) ui.personId = state.people[0].id;
  }

  function matchSharePointUser() {
    return;
  }

  function renderHeader() {
    $("org-label").textContent = state.settings.orgName || "Mateřská škola";
    $("year-label").textContent = "Čtrnáctidenní cyklus ranní a odpolední směny · školní rok " +
      (state.settings.schoolYear || "");
    var box = $("session-box");
    var acc = currentAccount();
    if (!acc) {
      box.innerHTML = "<p class=\"session-hint\" id=\"session-hint\">Přihlaste se svým jménem a heslem.</p>";
      $("nav").innerHTML = "";
      return;
    }
    var personSelect = "";
    if (canAdmin() && state.people.length) {
      var people = state.people.slice().sort(function (a, b) {
        return a.name.localeCompare(b.name, "cs");
      });
      personSelect = "<label>Výkaz zaměstnance<select id=\"person-select\">" +
        people.map(function (p) {
          return "<option value=\"" + esc(p.id) + "\"" + (p.id === ui.personId ? " selected" : "") + ">" +
            esc(p.name) + " · " + esc(M.roleLabel(p.role)) + "</option>";
        }).join("") + "</select></label>";
      if (!ui.personId && people[0]) ui.personId = people[0].id;
    }
    var roleLabel = canAdmin() ? "Správce" : "Zaměstnanec";
    box.innerHTML =
      "<p class=\"session-who\"><strong>" + esc(acc.name || acc.login) + "</strong>" +
      "<span class=\"badge " + (canAdmin() ? "approved" : "draft") + "\">" + esc(roleLabel) + "</span></p>" +
      "<p class=\"session-hint\">Přihlášení: " + esc(acc.login) +
      (ui.spUser ? " · SharePoint: " + esc(ui.spUser.Email || ui.spUser.Title) : "") + "</p>" +
      personSelect +
      "<div class=\"row\">" +
      "<button type=\"button\" class=\"secondary\" data-act=\"change-password\">Změnit heslo</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"logout\">Odhlásit</button>" +
      "</div>";

    var tabs = [];
    if (acc.personId || canAdmin()) tabs.push({ id: "timesheet", label: "Můj výkaz" });
    if (canAdmin()) {
      tabs.push({ id: "approve", label: "Schvalování" });
      tabs.push({ id: "admin", label: "Administrace" });
    }
    if (canAdmin()) tabs.push({ id: "license", label: "Licence M365" });
    if (!tabs.some(function (t) { return t.id === ui.view; })) ui.view = tabs[0] ? tabs[0].id : "timesheet";
    $("nav").innerHTML = tabs.map(function (t) {
      return "<button type=\"button\" data-view=\"" + t.id + "\" class=\"" + (ui.view === t.id ? "active" : "") + "\">" +
        esc(t.label) + "</button>";
    }).join("");
  }

  function render() {
    closeModal();
    renderHeader();
    var app = $("app");
    if (!currentAccount()) {
      app.innerHTML = A.hasActiveAdmin(state) ? renderLogin() : renderSetup();
      bind();
      return;
    }
    if (ui.view === "admin" && canAdmin()) app.innerHTML = renderAdmin();
    else if (ui.view === "approve" && canLead()) app.innerHTML = renderApprove();
    else if (ui.view === "license" && canAdmin()) app.innerHTML = renderLicense();
    else app.innerHTML = renderTimesheet();
    bind();
    if (ui.view === "admin" && ui.editPersonId) {
      var ed = document.getElementById("person-editor");
      if (ed && ed.scrollIntoView) ed.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  }

  function monthNav() {
    return "<div class=\"row no-print\">" +
      "<button type=\"button\" class=\"secondary\" data-act=\"prev-month\">←</button>" +
      "<strong style=\"min-width:12ch;text-align:center;text-transform:capitalize\">" +
      esc(M.monthName(ui.month)) + " " + ui.year + "</strong>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"next-month\">→</button>" +
      "<select id=\"month-jump\">" + monthsOptions() + "</select>" +
      "</div>";
  }

  function monthsOptions() {
    var html = "";
    var y = ui.year;
    for (var m = 1; m <= 12; m++) {
      html += "<option value=\"" + y + "-" + m + "\"" + (m === ui.month ? " selected" : "") + ">" +
        M.monthName(m) + " " + y + "</option>";
    }
    return html;
  }

  function renderTimesheet() {
    var person = currentPerson();
    if (!person) {
      if (!canAdmin()) {
        return "<div class=\"card\"><p>K tomuto účtu není přiřazen výkaz. Požádejte správce o spojení účtu se zaměstnancem.</p></div>";
      }
      return "<div class=\"card\"><p>Nejdřív v administraci přidejte zaměstnance, nebo načtěte data z Excelu.</p></div>";
    }
    var view = M.buildMonth(state, person, ui.year, ui.month);
    var locked = view.sheet.status === "submitted" || view.sheet.status === "approved";
    var weeks = view.weeks.map(function (w) {
      return "<div class=\"week-card\">" +
        "<div><strong>" + esc(w.label) + "</strong><div class=\"muted\">" +
        w.dates.map(function (d) { return M.formatCzechDate(M.parseISODate(d)); }).join(", ") +
        "</div></div>" +
        "<div class=\"shift-toggle\">" +
        "<button type=\"button\" " + (locked ? "disabled" : "") + " data-week=\"" + esc(w.key) + "\" data-shift=\"ranni\" class=\"" +
        (w.shift === "ranni" ? "on-ranni" : "") + "\">R</button>" +
        "<button type=\"button\" " + (locked ? "disabled" : "") + " data-week=\"" + esc(w.key) + "\" data-shift=\"odpoledni\" class=\"" +
        (w.shift === "odpoledni" ? "on-odpoledni" : "") + "\">O</button>" +
        "</div></div>";
    }).join("");

    var weeksHtml = person.alternatesShifts === false
      ? "<p class=\"hint\">Tato osoba nestřídá ranní a odpolední týden — platí jeden rozvrh.</p>"
      : weeks;

    var ped = M.isPedPerson(person);
    var rows = view.days.map(function (d) {
      var note = (d.exception && d.exception.note) || "";
      var absence = M.absenceLabel(d);
      if (ped) {
        return "<tr class=\"clickable\" data-day=\"" + esc(d.date) + "\">" +
          "<td>" + esc(M.formatCzechDate(d.dateObj)) + "</td>" +
          "<td>" + esc(d.weekdayShort) + "</td>" +
          "<td>" + esc(d.primaStart) + "</td>" +
          "<td>" + esc(d.primaEnd) + "</td>" +
          "<td>" + esc(d.neprimaStart) + "</td>" +
          "<td>" + esc(d.neprimaEnd) + "</td>" +
          "<td>" + esc(note) + "</td>" +
          "<td>" + esc(d.holiday ? "" : M.shiftCode(d.shift)) + "</td>" +
          "<td>" + esc(absence) + "</td>" +
          "</tr>";
      }
      return "<tr class=\"clickable\" data-day=\"" + esc(d.date) + "\">" +
        "<td>" + esc(M.formatCzechDate(d.dateObj)) + "</td>" +
        "<td>" + esc(d.weekdayShort) + "</td>" +
        "<td>" + esc(d.start) + "</td>" +
        "<td>" + esc(d.end) + "</td>" +
        "<td>" + (d.breakStart && d.breakEnd ? esc(d.breakStart + "–" + d.breakEnd) : (d.breakMin ? d.breakMin + " min" : "")) + "</td>" +
        "<td>" + esc(d.holiday ? "" : M.shiftCode(d.shift)) + "</td>" +
        "<td>" + esc(note) + "</td>" +
        "<td>" + esc(absence) + "</td>" +
        "</tr>";
    }).join("");

    var t = view.totals;
    var extraTiles = ped
      ? tile("Nepřímá práce", t.neprima) + tile("Přímá práce", t.prima)
      : "";
    var hourHead = ped
      ? "<tr><th rowspan=\"2\">Datum</th><th rowspan=\"2\">Den</th>" +
        "<th colspan=\"2\">Přímá</th><th colspan=\"2\">Nepřímá</th>" +
        "<th rowspan=\"2\">Poznámka</th><th rowspan=\"2\">Směna</th><th rowspan=\"2\">Důvod absence</th></tr>" +
        "<tr><th>Příchod</th><th>Odchod</th><th>Příchod</th><th>Odchod</th></tr>"
      : "<tr><th>Datum</th><th>Den</th><th>Příchod</th><th>Odchod</th><th>Pauza</th><th>Směna</th><th>Poznámka</th><th>Důvod absence</th></tr>";
    return "<div class=\"card\">" +
      "<div class=\"row\" style=\"justify-content:space-between\">" +
      "<div><h2 style=\"margin:0\">" + esc(person.name) + "</h2>" +
      "<p class=\"muted\">" + esc(M.roleLabel(person.role)) +
      (person.workplace ? " · " + esc(person.workplace) : "") +
      (person.personalNumber ? " · č. " + esc(person.personalNumber) : "") +
      " · úvazek " + esc(String(person.fte)) +
      " · celkem " + M.formatHours(M.personFond(person)) + " h/týden" +
      (ped ? " · přímá " + M.formatHours(M.personPrimaWeekly(person)) +
        " · nepřímá " + M.formatHours(M.personNeprimaWeekly(person)) : "") +
      "</p></div>" +
      "<span class=\"badge " + esc(view.sheet.status) + "\">" + esc(M.STATUS[view.sheet.status]) + "</span></div>" +
      monthNav() +
      (view.sheet.returnComment ? "<p class=\"flash error\" style=\"display:block\">Vráceno: " + esc(view.sheet.returnComment) + "</p>" : "") +
      "<p class=\"hint\">U každého týdne zvolte směnu R (ranní) nebo O (odpolední). Časy se doplní ze šablony, stejně jako v Excelu. Kliknutím na den zadáte absenci (D, PN, OČR, MD) nebo úpravu času.</p>" +
      weeksHtml +
      "</div>" +
      "<div class=\"totals\">" +
      tile("Fond / šablona", t.fond) +
      extraTiles +
      tile("Odpracováno", t.worked) +
      tile("Dovolená", t.dovolena) +
      tile("Nemoc / OČR", t.nemoc + t.ocr) +
      tile("Přesčas", t.prescas) +
      "</div>" +
      "<div class=\"card\">" +
      "<h3>Výkaz</h3>" +
      "<div class=\"table-wrap\"><table class=\"vykaz\">" +
      "<thead>" + hourHead + "</thead><tbody>" + rows + "</tbody></table></div>" +
      "<div class=\"row no-print\" style=\"margin-top:16px\">" +
      (locked ? "" : "<button type=\"button\" data-act=\"submit\">Odeslat vedení</button>") +
      (view.sheet.status === "submitted" && canLead() ? "<button type=\"button\" data-act=\"unlock\">Vrátit k úpravě</button>" : "") +
      "<button type=\"button\" class=\"secondary\" data-act=\"print\">Tisk / PDF</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"xlsx-one\">Excel výkazu</button>" +
      "</div></div>";
  }

  function tile(label, value) {
    return "<div class=\"total\"><b>" + M.formatHours(value) + " h</b><span>" + esc(label) + "</span></div>";
  }

  function renderApprove() {
    var people = state.people.filter(function (p) { return p.active; });
    var rows = people.map(function (p) {
      var view = M.buildMonth(state, p, ui.year, ui.month);
      var s = view.sheet;
      return "<tr>" +
        "<td>" + esc(p.name) + "<div class=\"muted\">" + esc(M.roleLabel(p.role)) + "</div></td>" +
        "<td><span class=\"badge " + esc(s.status) + "\">" + esc(M.STATUS[s.status]) + "</span></td>" +
        "<td>" + M.formatHours(view.totals.fond) + "</td>" +
        "<td>" + M.formatHours(view.totals.neprima) + "</td>" +
        "<td>" + M.formatHours(view.totals.prima) + "</td>" +
        "<td>" + M.formatHours(view.totals.worked) + "</td>" +
        "<td>" + M.formatHours(view.totals.dovolena) + "</td>" +
        "<td>" + M.formatHours(view.totals.nemoc + view.totals.ocr) + "</td>" +
        "<td>" + M.formatHours(view.totals.prescas) + "</td>" +
        "<td class=\"row\">" +
        "<button type=\"button\" class=\"secondary\" data-open=\"" + esc(p.id) + "\">Otevřít</button>" +
        (s.status === "submitted" ? "<button type=\"button\" data-approve=\"" + esc(p.id) + "\">Schválit</button>" +
          "<button type=\"button\" class=\"secondary\" data-return=\"" + esc(p.id) + "\">Vrátit</button>" : "") +
        "</td></tr>";
    }).join("");
    return "<div class=\"card\">" +
      "<h2>Schvalování a podklady pro mzdy</h2>" +
      monthNav() +
      "<p class=\"hint\">Vedení vidí všechny výkazy za měsíc, schvaluje je a exportuje souhrn pro mzdovou účetní (list Přehled + list Dny).</p>" +
      "<div class=\"row no-print\" style=\"margin-bottom:12px\">" +
      "<button type=\"button\" data-act=\"xlsx-all\">Excel pro mzdovou účetní</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"print-all\">Tisk všech schválených</button>" +
      "</div>" +
      "<div class=\"table-wrap\"><table><thead><tr>" +
      "<th>Zaměstnanec</th><th>Stav</th><th>Fond</th><th>Nepřímá</th><th>Přímá</th><th>Odpracováno</th><th>Dovolená</th><th>Nemoc/OČR</th><th>Přesčas</th><th></th>" +
      "</tr></thead><tbody>" + (rows || "<tr><td colspan=\"10\">Nikdo není v seznamu.</td></tr>") +
      "</tbody></table></div></div>";
  }

  function renderLogin() {
    return "<div class=\"card login-card\">" +
      "<h2>Přihlášení</h2>" +
      "<p class=\"hint\">Každý má vlastní login a heslo. Zaměstnanec vidí jen svůj výkaz (směny R/O a výjimky). Správce vidí celou aplikaci.</p>" +
      "<div class=\"grid-2\">" +
      "<label class=\"field\">Login<input id=\"login-name\" autocomplete=\"username\"></label>" +
      "<label class=\"field\">Heslo<input id=\"login-pass\" type=\"password\" autocomplete=\"current-password\"></label>" +
      "</div>" +
      "<div class=\"row\" style=\"margin-top:14px\"><button type=\"button\" data-act=\"login\">Přihlásit</button></div>" +
      "</div>";
  }

  function renderSetup() {
    return "<div class=\"card login-card\">" +
      "<h2>Zřídit prvního správce</h2>" +
      "<p class=\"hint\">Aplikace zatím nemá žádný účet. První správce může být zaměstnanec i někdo mimo školku (účetní, zřizovatel). Další hesla pak přidělíte v administraci.</p>" +
      "<div class=\"grid-2\">" +
      "<label class=\"field\">Jméno a příjmení<input id=\"setup-name\" autocomplete=\"name\"></label>" +
      "<label class=\"field\">Login<input id=\"setup-login\" autocomplete=\"username\" placeholder=\"např. novak\"></label>" +
      "<label class=\"field\">Heslo<input id=\"setup-pass\" type=\"password\" autocomplete=\"new-password\"></label>" +
      "<label class=\"field\">Heslo znovu<input id=\"setup-pass2\" type=\"password\" autocomplete=\"new-password\"></label>" +
      "</div>" +
      "<div class=\"row\" style=\"margin-top:14px\"><button type=\"button\" data-act=\"setup-admin\">Vytvořit správce a přihlásit</button></div>" +
      "</div>";
  }

  function renderAccessCard() {
    var missing = state.people.filter(function (p) {
      return p.active && !A.accountForPerson(state, p.id);
    }).length;
    var rows = A.accountsOf(state).slice().sort(function (a, b) {
      return String(a.name || a.login).localeCompare(String(b.name || b.login), "cs");
    }).map(function (a) {
      var person = a.personId ? M.getPerson(state, a.personId) : null;
      return "<tr>" +
        "<td>" + esc(a.name || a.login) + "<div class=\"muted\">" + esc(a.login) + "</div></td>" +
        "<td>" + (a.role === "admin" ? "Správce" : "Zaměstnanec") + "</td>" +
        "<td>" + esc(person ? person.name : (a.role === "admin" ? "mimo zaměstnance" : "—")) + "</td>" +
        "<td>" + (a.active === false ? "<span class=\"badge returned\">neaktivní</span>" : "<span class=\"badge approved\">aktivní</span>") + "</td>" +
        "<td><div class=\"row\">" +
        "<button type=\"button\" class=\"secondary\" data-act=\"reset-pass\" data-acc=\"" + esc(a.id) + "\">Nové heslo</button>" +
        (a.id === ui.accountId ? "" :
          "<button type=\"button\" class=\"secondary\" data-act=\"toggle-acc\" data-acc=\"" + esc(a.id) + "\">" +
          (a.active === false ? "Aktivovat" : "Deaktivovat") + "</button>") +
        "</div></td></tr>";
    }).join("");
    return "<div class=\"card\">" +
      "<h2>Přístupy a hesla</h2>" +
      "<p class=\"hint\">Zaměstnancům vygenerujte heslo ve tvaru slovo-číslice (např. javor-4821) a předejte ho osobně nebo na papírku. Login je z příjmení. Správce může být i člověk, který ve školce nepracuje — neuvede se k žádnému výkazu, ale uvidí administraci, schvalování i všechny výkazy.</p>" +
      "<div class=\"row\" style=\"margin-bottom:12px\">" +
      "<button type=\"button\" data-act=\"issue-missing\">" +
      (missing ? "Vygenerovat hesla (" + missing + " bez přístupu)" : "Vygenerovat chybějící hesla") +
      "</button>" +
      "</div>" +
      "<div class=\"table-wrap\"><table><thead><tr>" +
      "<th>Účet</th><th>Oprávnění</th><th>Výkaz</th><th>Stav</th><th></th>" +
      "</tr></thead><tbody>" + (rows || "<tr><td colspan=\"5\">Zatím jen tento správce po zřízení účtu.</td></tr>") +
      "</tbody></table></div>" +
      "<h3 style=\"margin-top:18px\">Přidat správce</h3>" +
      "<div class=\"grid-2\">" +
      "<label class=\"field\">Jméno<input id=\"adm-name\"></label>" +
      "<label class=\"field\">Login<input id=\"adm-login\" placeholder=\"ponechte prázdné = z příjmení\"></label>" +
      "</div>" +
      "<div class=\"row\" style=\"margin-top:12px\">" +
      "<button type=\"button\" data-act=\"add-admin\">Přidat správce a vygenerovat heslo</button>" +
      "</div>" +
      "</div>";
  }

  function renderAdmin() {
    var people = state.people.slice().sort(function (a, b) {
      return a.name.localeCompare(b.name, "cs");
    }).map(function (p) {
      var open = p.id === ui.editPersonId;
      var acc = A.accountForPerson(state, p.id);
      var row = "<div class=\"person-row" + (open ? " open" : "") + "\">" +
        "<div><strong>" + esc(p.name) + "</strong>" +
        (p.active ? "" : " <span class=\"badge returned\">neaktivní</span>") +
        (p.hasSchedule === false ? " <span class=\"badge returned\">bez rozvrhu</span>" : "") +
        (!acc ? " <span class=\"badge returned\">bez hesla</span>" : "") +
        "<div class=\"muted\">" + esc(M.formTypeLabel(p.formType)) + " · " + esc(M.roleLabel(p.role)) +
        (p.workplace ? " · " + esc(p.workplace) : "") +
        (p.classroom ? " · " + esc(p.classroom) : "") +
        " · úvazek " + esc(String(p.fte)) +
        (p.personalNumber ? " · č. " + esc(p.personalNumber) : "") +
        (acc ? " · login " + esc(acc.login) : "") +
        "</div></div>" +
        "<div class=\"row\">" +
        "<button type=\"button\" class=\"secondary\" data-edit=\"" + esc(p.id) + "\">Pracovní doba</button>" +
        (acc
          ? "<button type=\"button\" class=\"secondary\" data-act=\"reset-pass\" data-acc=\"" + esc(acc.id) + "\">Nové heslo</button>"
          : "<button type=\"button\" class=\"secondary\" data-act=\"issue-pass\" data-person=\"" + esc(p.id) + "\">Přidělit heslo</button>") +
        "<button type=\"button\" class=\"secondary\" data-toggle-active=\"" + esc(p.id) + "\">" +
        (p.active ? "Deaktivovat" : "Aktivovat") + "</button>" +
        "</div></div>";
      if (!open) return row;
      return "<div class=\"person-block\" id=\"person-editor\">" + row + renderPersonEditor(p) + "</div>";
    }).join("");

    return "<div class=\"card\">" +
      "<h2>Nastavení školního roku</h2>" +
      "<p class=\"hint\">Na začátku roku vyplňte organizaci, školní rok a pondělí (nebo libovolný den) prvního týdne cyklu. Sudé týdny od tohoto data se předvyplní jako ranní, liché jako odpolední — zaměstnanec to ve výkazu může změnit.</p>" +
      "<div class=\"grid-2\">" +
      field("org-name", "Název", state.settings.orgName) +
      field("school-year", "Školní rok", state.settings.schoolYear) +
      "<label class=\"field\">Začátek čtrnáctidenního cyklu<input type=\"date\" id=\"cycle-start\" value=\"" + esc(state.settings.cycleStart) + "\"></label>" +
      field("weekly-full", "Týdenní fond plného úvazku (h)", state.settings.weeklyHoursFull) +
      "</div>" +
      "<div class=\"row\" style=\"margin-top:14px\"><button type=\"button\" data-act=\"save-settings\">Uložit nastavení</button></div>" +
      "</div>" +
      renderAccessCard() +
      "<div class=\"card\">" +
      "<div class=\"row\" style=\"justify-content:space-between\">" +
      "<h2 style=\"margin:0\">Zaměstnanci a šablony směn</h2>" +
      "<button type=\"button\" data-act=\"add-person\">Přidat osobu</button></div>" +
      "<p class=\"hint\">Časy bereme z Excelu Personalistika (Pedagog / Provozní / Extra). Po úpravě v Excelu soubor uložte a tady ho znovu načtěte.</p>" +
      "<div class=\"people-list\">" + (people || "<p>Zatím nikdo.</p>") + "</div>" +
      "</div>" +
      renderSharePointCard() +
      "<div class=\"card\">" +
      "<h2>Data</h2>" +
      "<div class=\"row\">" +
      "<button type=\"button\" data-act=\"load-school\">Načíst zaměstnance z Excelu</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"demo\">Načíst ukázkové zaměstnance</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"export-json\">Záloha JSON</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"import-json\">Obnovit ze zálohy</button>" +
      "<input type=\"file\" id=\"import-file\" accept=\"application/json\" hidden>" +
      "<input type=\"file\" id=\"excel-files\" accept=\".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\" multiple hidden>" +
      "</div>" +
      "<p class=\"hint\">Vyberte uložený Pedagog.xlsx, případně i Provozní.xlsx a Extra.xlsx, ze složky Personalistika.</p>" +
      "</div></div>";
  }

  function field(id, label, value) {
    return "<label class=\"field\">" + esc(label) + "<input id=\"" + id + "\" value=\"" + esc(value) + "\"></label>";
  }

  function renderPersonEditor(person) {
    if (!person) return "";
    return "<div class=\"person-editor\">" +
      "<h3>" + esc(person.name) + "</h3>" +
      "<div class=\"grid-2\">" +
      "<label class=\"field\">Jméno<input id=\"p-name\" value=\"" + esc(person.name) + "\"></label>" +
      "<label class=\"field\">Role<select id=\"p-role\">" +
      M.ROLES.map(function (r) {
        return "<option value=\"" + r.id + "\"" + (person.role === r.id ? " selected" : "") + ">" + esc(r.label) + "</option>";
      }).join("") + "</select></label>" +
      "<label class=\"field\">Typ formuláře<select id=\"p-form\">" +
      [["pedagog", "Pedagogický"], ["provoz", "Provozní"], ["extra", "Extra"]].map(function (row) {
        return "<option value=\"" + row[0] + "\"" + ((person.formType || (M.isPedPerson(person) ? "pedagog" : "provoz")) === row[0] ? " selected" : "") + ">" + esc(row[1]) + "</option>";
      }).join("") + "</select></label>" +
      "<label class=\"field\">Třída<input id=\"p-class\" value=\"" + esc(person.classroom || "") + "\"></label>" +
      "<label class=\"field\">Pracoviště<input id=\"p-workplace\" value=\"" + esc(person.workplace || "") + "\"></label>" +
      "<label class=\"field\">Osobní číslo<input id=\"p-personal\" value=\"" + esc(person.personalNumber || "") + "\"></label>" +
      "<label class=\"field\">Úvazek (1 = plný)<input id=\"p-fte\" type=\"number\" min=\"0\" max=\"1\" step=\"0.1\" value=\"" + esc(person.fte) + "\"></label>" +
      "<label class=\"field\">Týdenní fond při úvazku 1,0 (h)<input id=\"p-hours\" type=\"number\" min=\"0\" step=\"0.5\" value=\"" + esc(person.weeklyHours) + "\"></label>" +
      (M.isPedPerson(person)
        ? "<label class=\"field\">Přímá práce / týden při úvazku 1,0 (h)<input id=\"p-prima\" type=\"number\" min=\"0\" step=\"0.5\" value=\"" + esc(M.personPrimaBase(person)) + "\"></label>" +
          "<label class=\"field\">Snížení přímé práce (SníženíPP, h)<input id=\"p-korekce\" type=\"number\" step=\"0.5\" value=\"" + esc(person.korekcePrima || 0) + "\"></label>" +
          "<label class=\"field\">Skutečná PP / týden (h)<input id=\"p-skutecna-pp\" type=\"number\" readonly tabindex=\"-1\" value=\"" + esc(M.personPrimaWeekly(person)) + "\"></label>" +
          "<label class=\"field\">Skutečná NP / týden (h)<input id=\"p-skutecna-np\" type=\"number\" readonly tabindex=\"-1\" value=\"" + esc(M.personNeprimaWeekly(person)) + "\"></label>" +
          "<p class=\"hint\" id=\"p-hours-hint\" style=\"grid-column:1/-1\">Skutečná PP = přímá práce × úvazek − SníženíPP. Skutečná NP = fond × úvazek − skutečná PP (SníženíPP se přičte k nepřímé). Fond " +
          M.formatHours(M.personFond(person)) + " h.</p>"
        : "") +
      "<label class=\"field\">E-mail (pro spárování se SharePoint účtem)<input id=\"p-email\" value=\"" + esc(person.email || "") + "\"></label>" +
      "<label class=\"check\" style=\"grid-column:1/-1;margin:4px 0 0\"><input type=\"checkbox\" id=\"p-alt\"" + (person.alternatesShifts === false ? "" : " checked") + ">" +
      "<span class=\"check-text\">Střídá ranní a odpolední týden" +
      "<span class=\"hint\">Zaškrtnuto: jeden týden platí ranní šablona, druhý odpolední. Odškrtnuto: každý týden stejný rozvrh (ranní týden).</span></span></label>" +
      "</div>" +
      "<p class=\"hint\">" + (M.isPedPerson(person)
        ? "Mřížka jako v Excelu: nepřímá – přímá – nepřímá. Prázdný úsek se nepočítá. U odpolední směny bývá nepřímá v prvním sloupci."
        : "Příchod, odchod a pauza od–do. Prázdný den = ten den nepracuje.") + "</p>" +
      templateTable("ranni", "Ranní týden", person.templates.ranni, M.isPedPerson(person)) +
      templateTable("odpoledni", "Odpolední týden", person.templates.odpoledni, M.isPedPerson(person)) +
      "<div class=\"row\" style=\"margin-top:12px\">" +
      "<button type=\"button\" data-act=\"save-person\">Uložit osobu a šablony</button>" +
      "<button type=\"button\" data-act=\"restore-excel\">Načíst časy z Excelu</button>" +
      "<button type=\"button\" class=\"secondary\" data-act=\"close-editor\">Zavřít</button>" +
      "</div>" +
      "<p class=\"hint\">Načíst časy z Excelu: uložte Pedagog.xlsx a vyberte ho. Aplikace vezme mřížku z listu této osoby.</p>" +
      "</div>";
  }

  function timeRange(kind, day, kStart, kEnd, start, end) {
    return "<div class=\"range\">" +
      "<input data-tpl=\"" + kind + "\" data-day=\"" + day + "\" data-k=\"" + kStart + "\" value=\"" + esc(start || "") + "\">" +
      "<span>–</span>" +
      "<input data-tpl=\"" + kind + "\" data-day=\"" + day + "\" data-k=\"" + kEnd + "\" value=\"" + esc(end || "") + "\">" +
      "</div>";
  }

  function templateTable(kind, title, days, ped) {
    if (ped) {
      var extra = M.WEEKDAYS.some(function (d) {
        var s = M.normalizeSlot(days && (days[d.n] || days[String(d.n)]), true);
        return s.p2Start || s.n3Start;
      });
      var body = M.WEEKDAYS.map(function (d) {
        var slot = M.normalizeSlot(days && (days[d.n] || days[String(d.n)]), true);
        return "<tr><th>" + esc(d.long) + "</th>" +
          "<td>" + timeRange(kind, d.n, "n1Start", "n1End", slot.n1Start, slot.n1End) + "</td>" +
          "<td>" + timeRange(kind, d.n, "primaStart", "primaEnd", slot.primaStart, slot.primaEnd) + "</td>" +
          "<td>" + timeRange(kind, d.n, "n2Start", "n2End", slot.n2Start, slot.n2End) + "</td>" +
          (extra
            ? "<td>" + timeRange(kind, d.n, "p2Start", "p2End", slot.p2Start, slot.p2End) + "</td>" +
              "<td>" + timeRange(kind, d.n, "n3Start", "n3End", slot.n3Start, slot.n3End) + "</td>"
            : "") +
          "</tr>";
      }).join("");
      return "<h3>" + esc(title) + "</h3><div class=\"table-wrap\"><table class=\"template-grid\"><thead><tr>" +
        "<th>Den</th><th class=\"th-schema\">Nepřímá<div class=\"muted\">od – do</div></th>" +
        "<th class=\"th-schema\">Přímá<div class=\"muted\">od – do</div></th>" +
        "<th class=\"th-schema\">Nepřímá<div class=\"muted\">od – do</div></th>" +
        (extra
          ? "<th class=\"th-schema\">Přímá<div class=\"muted\">od – do</div></th>" +
            "<th class=\"th-schema\">Nepřímá<div class=\"muted\">od – do</div></th>"
          : "") +
        "</tr></thead><tbody>" + body + "</tbody></table></div>";
    }
    var bodySimple = M.WEEKDAYS.map(function (d) {
      var slot = M.normalizeSlot(days && (days[d.n] || days[String(d.n)]), false);
      return "<tr><th>" + esc(d.long) + "</th>" +
        "<td><input data-tpl=\"" + kind + "\" data-day=\"" + d.n + "\" data-k=\"start\" value=\"" + esc(slot.start) + "\"></td>" +
        "<td><input data-tpl=\"" + kind + "\" data-day=\"" + d.n + "\" data-k=\"end\" value=\"" + esc(slot.end) + "\"></td>" +
        "<td>" + timeRange(kind, d.n, "breakStart", "breakEnd", slot.breakStart, slot.breakEnd) + "</td>" +
        "</tr>";
    }).join("");
    return "<h3>" + esc(title) + "</h3><div class=\"table-wrap\"><table class=\"template-grid\"><thead><tr>" +
      "<th>Den</th><th>Příchod</th><th>Odchod</th><th>Pauza od – do</th></tr></thead><tbody>" + bodySimple + "</tbody></table></div>";
  }

  function renderSharePointCard() {
    var on = ui.storage === "sharepoint";
    return "<div class=\"card\">" +
      "<h2>SharePoint</h2>" +
      "<p><span class=\"sp-dot " + (on ? "on" : "") + "\"></span>" +
      (on ? "Připojeno k " + esc(ui.spWeb) : "Běží lokálně v prohlížeči") + "</p>" +
      "<p class=\"hint\">Soubor <code>index.html</code> nahrajte do knihovny dokumentů webu školky (nejlépe Site Assets). Pak tu vytvořte seznamy — stanou se databází. Moderní SharePoint někdy stahuje HTML místo spuštění; v tom případě požádejte správce o povolení vlastních skriptů, nebo stránku otevřete jako odkaz z klasického zobrazení.</p>" +
      "<div class=\"row\">" +
      "<button type=\"button\" data-act=\"sp-detect\">Zjistit SharePoint</button>" +
      "<button type=\"button\" data-act=\"sp-provision\"" + (ui.spWeb ? "" : " disabled") + ">Vytvořit seznamy</button>" +
      "<button type=\"button\" data-act=\"sp-save\"" + (ui.spWeb ? "" : " disabled") + ">Uložit do seznamů</button>" +
      "<button type=\"button\" data-act=\"sp-load\"" + (ui.spWeb ? "" : " disabled") + ">Načíst ze seznamů</button>" +
      "</div></div>";
  }

  function renderLicense() {
    return "<div class=\"card license\">" +
      "<h2>Licence Microsoft 365 Education</h2>" +
      "<p>Tenant školky z tohoto počítače ověřit nelze. Níže je, co z oficiálního srovnání Microsoftu platí pro Power Apps / Power Automate / SharePoint, a jak to má zkontrolovat správce.</p>" +
      "<div class=\"ok\"><strong>Tento výkaz Power Apps nepotřebuje.</strong> Stačí SharePoint (nebo i jen prohlížeč s localStorage). Seznamy SharePointu jsou ve všech běžných edu plánech včetně A1. Power Apps je volitelná pozdější obálka se stejným datovým modelem.</div>" +
      "<h3>Co bývá v plánu zahrnuto</h3>" +
      "<ul>" +
      "<li><strong>SharePoint a Microsoft Lists</strong> — A1 / A3 / A5 (faculty). Databáze tohoto výkazu.</li>" +
      "<li><strong>Power Apps for Microsoft 365</strong> — plně u A3 a A5; u A1 jen omezená / seed licence. Stačí na aplikaci nad seznamy, bez premium konektorů a bez Dataverse.</li>" +
      "<li><strong>Power Automate for Microsoft 365</strong> — u A3/A5 cloud toky (schválení e-mailem). Ne desktop RPA, ne premium konektory.</li>" +
      "</ul>" +
      "<h3>Jak ověřit v tenantu</h3>" +
      "<ol>" +
      "<li>Otevřít <a href=\"https://admin.microsoft.com\" target=\"_blank\" rel=\"noreferrer\">admin.microsoft.com</a> → Fakturace → Licence. Hledat Microsoft 365 A1 / A3 / A5 (Faculty).</li>" +
      "<li>Otevřít <a href=\"https://make.powerapps.com\" target=\"_blank\" rel=\"noreferrer\">make.powerapps.com</a> pod účtem hospodářky. Když jde vytvořit Canvas app a připojit SharePoint seznam, licence na obálku stačí.</li>" +
      "<li>Na webu SharePointu vytvořit zkušební seznam. Když to jde, databáze pro tento výkaz je k dispozici.</li>" +
      "<li>Pokud HTML v knihovně nejde spustit, správce v admin centru povolí vlastní skripty na daném webu, nebo se použije Power App nad stejnými seznamy.</li>" +
      "</ol>" +
      "<div class=\"warn\">Premium Power Apps / Dataverse se nekupuje. SMS brána s tímto výkazem nesouvisí.</div>" +
      "<p class=\"hint\">Zdroj: Microsoft 365 Education license comparison a Power Platform licensing overview (standard konektor SharePoint).</p>" +
      "</div>";
  }

  function bind() {
    $("nav").onclick = function (e) {
      var btn = e.target.closest("button[data-view]");
      if (!btn) return;
      if (!currentAccount()) return;
      ui.view = btn.getAttribute("data-view");
      render();
    };
    var sel = $("person-select");
    if (sel) {
      sel.onchange = function () {
        if (!canAdmin()) return;
        ui.personId = this.value;
        render();
      };
    }
    $("session-box").onclick = onAppClick;
    $("app").onclick = onAppClick;
    $("app").onchange = onAppChange;
    $("app").oninput = onAppInput;
    $("app").onkeydown = onAppKey;
  }

  function onAppKey(e) {
    if (e.key !== "Enter") return;
    var id = e.target && e.target.id;
    if (id === "login-name" || id === "login-pass") {
      e.preventDefault();
      doLogin();
    }
    if (id === "setup-name" || id === "setup-login" || id === "setup-pass" || id === "setup-pass2") {
      e.preventDefault();
      setupAdmin();
    }
  }

  function editorPersonFromForm(person) {
    var draft = Object.assign({}, person);
    if ($("p-fte")) draft.fte = Number($("p-fte").value) || 0;
    if ($("p-hours")) draft.weeklyHours = Number($("p-hours").value) || 0;
    if ($("p-prima")) draft.primaWeekly = Number($("p-prima").value) || 0;
    if ($("p-korekce")) draft.korekcePrima = Number($("p-korekce").value) || 0;
    return draft;
  }

  function refreshComputedHours() {
    var pp = $("p-skutecna-pp");
    var np = $("p-skutecna-np");
    if (!pp || !np) return;
    var person = M.getPerson(state, ui.editPersonId);
    if (!person) return;
    var draft = editorPersonFromForm(person);
    pp.value = M.personPrimaWeekly(draft);
    np.value = M.personNeprimaWeekly(draft);
    var hint = $("p-hours-hint");
    if (hint) {
      hint.textContent = "Skutečná PP = přímá práce × úvazek − SníženíPP. Skutečná NP = fond × úvazek − skutečná PP (SníženíPP se přičte k nepřímé). Fond " +
        M.formatHours(M.personFond(draft)) + " h.";
    }
  }

  function onAppInput(e) {
    var id = e.target && e.target.id;
    if (id === "p-fte" || id === "p-hours" || id === "p-prima" || id === "p-korekce") {
      refreshComputedHours();
    }
  }

  function onAppChange(e) {
    var t = e.target;
    if (t.id === "month-jump") {
      var p = t.value.split("-");
      ui.year = Number(p[0]);
      ui.month = Number(p[1]);
      render();
      return;
    }
    if (t.id === "p-fte" || t.id === "p-hours" || t.id === "p-prima" || t.id === "p-korekce") {
      refreshComputedHours();
    }
  }

  function onAppClick(e) {
    var t = e.target.closest("[data-act],[data-week],[data-day],[data-edit],[data-toggle-active],[data-open],[data-approve],[data-return]");
    if (!t) return;
    var act = t.getAttribute("data-act");
    if (act === "login") { doLogin(); return; }
    if (act === "setup-admin") { setupAdmin(); return; }
    if (act === "logout") { doLogout(); return; }
    if (!currentAccount()) return;
    if (act === "change-password") { openChangePassword(); return; }

    var needAdmin = {
      "print-all": 1, "xlsx-all": 1, "save-settings": 1, "add-person": 1, "save-person": 1,
      "restore-excel": 1, "load-school": 1, "demo": 1, "export-json": 1, "import-json": 1,
      "sp-detect": 1, "sp-provision": 1, "sp-save": 1, "sp-load": 1, "issue-pass": 1,
      "issue-missing": 1, "reset-pass": 1, "add-admin": 1, "toggle-acc": 1, "unlock": 1
    };
    if (needAdmin[act] && !canAdmin()) return;
    if ((t.getAttribute("data-edit") || t.getAttribute("data-toggle-active") || t.getAttribute("data-open") ||
      t.getAttribute("data-approve") || t.getAttribute("data-return")) && !canAdmin()) return;

    if (t.getAttribute("data-week")) {
      setWeekShift(t.getAttribute("data-week"), t.getAttribute("data-shift"));
      return;
    }
    if (t.getAttribute("data-day")) {
      openException(t.getAttribute("data-day"));
      return;
    }
    if (t.getAttribute("data-edit")) {
      ui.editPersonId = t.getAttribute("data-edit");
      render();
      return;
    }
    if (t.getAttribute("data-toggle-active")) {
      var p = M.getPerson(state, t.getAttribute("data-toggle-active"));
      if (p) { p.active = !p.active; persist().then(render); }
      return;
    }
    if (t.getAttribute("data-open")) {
      ui.personId = t.getAttribute("data-open");
      ui.view = "timesheet";
      render();
      return;
    }
    if (t.getAttribute("data-approve")) {
      approvePerson(t.getAttribute("data-approve"));
      return;
    }
    if (t.getAttribute("data-return")) {
      returnPerson(t.getAttribute("data-return"));
      return;
    }
    if (act === "prev-month") { shiftMonth(-1); }
    if (act === "next-month") { shiftMonth(1); }
    if (act === "submit") submitSheet();
    if (act === "unlock") {
      var person = currentPerson();
      var sheet = M.getTimesheet(state, person.id, ui.year, ui.month);
      if (sheet) { sheet.status = "returned"; persist().then(function () { flash("Výkaz je znovu k úpravě."); render(); }); }
    }
    if (act === "print") printOne(currentPerson());
    if (act === "print-all") printAllApproved();
    if (act === "xlsx-one") exportExcel([currentPerson()], false);
    if (act === "xlsx-all") exportExcel(state.people.filter(function (p) { return p.active; }), true);
    if (act === "save-settings") saveSettings();
    if (act === "add-person") addPerson();
    if (act === "save-person") savePerson();
    if (act === "restore-excel") restoreExcelTimes();
    if (act === "close-editor") { ui.editPersonId = null; render(); }
    if (act === "load-school") pickExcelFiles(null);
    if (act === "demo") loadDemo();
    if (act === "export-json") exportJson();
    if (act === "import-json") $("import-file").click();
    if (act === "sp-detect") detectSharePoint(true);
    if (act === "sp-provision") runProvision();
    if (act === "sp-save") saveSharePoint();
    if (act === "sp-load") loadSharePoint();
    if (act === "issue-pass") issuePasswordForPerson(t.getAttribute("data-person"));
    if (act === "issue-missing") issueMissingPasswords();
    if (act === "reset-pass") resetAccountPassword(t.getAttribute("data-acc"));
    if (act === "add-admin") addAdminAccount();
    if (act === "toggle-acc") toggleAccount(t.getAttribute("data-acc"));
  }

  document.addEventListener("change", function (e) {
    if (e.target && e.target.id === "import-file" && e.target.files[0]) {
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var data = JSON.parse(reader.result);
          if (!data || data.version !== 1) throw new Error("Neplatná záloha");
          if (!canAdmin()) throw new Error("Zálohu může obnovit jen správce");
          state = data;
          if (!state.accounts) state.accounts = [];
          restoreSession();
          ui.personId = (state.people[0] || {}).id || null;
          persist().then(function () { flash("Záloha obnovena."); render(); });
        } catch (err) {
          flash(String(err.message || err), true);
        }
      };
      reader.readAsText(e.target.files[0]);
    }
    if (e.target && e.target.id === "excel-files" && e.target.files && e.target.files.length) {
      if (canAdmin()) importExcelFiles(e.target.files);
      e.target.value = "";
    }
  });

  function shiftMonth(delta) {
    var d = new Date(ui.year, ui.month - 1 + delta, 1);
    ui.year = d.getFullYear();
    ui.month = d.getMonth() + 1;
    render();
  }

  function setWeekShift(week, shift) {
    var person = currentPerson();
    if (!person) return;
    var sheet = M.ensureTimesheet(state, person.id, ui.year, ui.month);
    if (sheet.status === "submitted" || sheet.status === "approved") return;
    sheet.weekShifts[week] = shift;
    persist().then(render);
  }

  function openException(date) {
    var person = currentPerson();
    var sheet = M.ensureTimesheet(state, person.id, ui.year, ui.month);
    if (sheet.status === "submitted" || sheet.status === "approved") {
      flash("Schválený nebo odeslaný výkaz nelze měnit.", true);
      return;
    }
    var view = M.buildMonth(state, person, ui.year, ui.month);
    var day = view.days.find(function (d) { return d.date === date; });
    if (!day || day.holiday) return;
    var ex = sheet.exceptions[date] || {};
    var modal = $("modal");
    var ped = M.isPedPerson(person);
    var absence = isAbsenceType(ex.type);
    var timeFields = ped
      ? "<div class=\"grid-2\" id=\"ex-times\">" +
        "<label class=\"field\">Přímá od<input id=\"ex-prima-start\" value=\"" + esc(absence ? "" : (ex.primaStart || day.primaStart)) + "\"></label>" +
        "<label class=\"field\">Přímá do<input id=\"ex-prima-end\" value=\"" + esc(absence ? "" : (ex.primaEnd || day.primaEnd)) + "\"></label>" +
        "<label class=\"field\">Nepřímá od<input id=\"ex-neprima-start\" value=\"" + esc(absence ? "" : (ex.neprimaStart || day.neprimaStart)) + "\"></label>" +
        "<label class=\"field\">Nepřímá do<input id=\"ex-neprima-end\" value=\"" + esc(absence ? "" : (ex.neprimaEnd || day.neprimaEnd)) + "\"></label>" +
        "</div>"
      : "<div class=\"grid-2\" id=\"ex-times\">" +
        "<label class=\"field\">Příchod<input id=\"ex-start\" value=\"" + esc(absence ? "" : (ex.start || day.start)) + "\"></label>" +
        "<label class=\"field\">Odchod<input id=\"ex-end\" value=\"" + esc(absence ? "" : (ex.end || day.end)) + "\"></label>" +
        "<label class=\"field\">Pauza od<input id=\"ex-break-start\" value=\"" + esc(absence ? "" : (ex.breakStart || day.breakStart || "")) + "\"></label>" +
        "<label class=\"field\">Pauza do<input id=\"ex-break-end\" value=\"" + esc(absence ? "" : (ex.breakEnd || day.breakEnd || "")) + "\"></label>" +
        "</div>";
    $("modal-card").innerHTML =
      "<h3>Výjimka " + esc(M.formatCzechDateFull(day.dateObj)) + "</h3>" +
      "<label class=\"field\">Typ<select id=\"ex-type\">" +
      "<option value=\"\">Práce podle šablony</option>" +
      M.EXCEPTION_TYPES.map(function (x) {
        var label = x.code ? x.code + " — " + x.label : x.label;
        return "<option value=\"" + x.id + "\"" + (ex.type === x.id ? " selected" : "") + ">" + esc(label) + "</option>";
      }).join("") + "</select></label>" +
      timeFields +
      "<label class=\"field\">Poznámka<input id=\"ex-note\" value=\"" + esc(ex.note || "") + "\"></label>" +
      "<div class=\"row\" style=\"margin-top:12px\">" +
      "<button type=\"button\" id=\"ex-save\">Uložit</button>" +
      "<button type=\"button\" class=\"secondary\" id=\"ex-clear\">Zrušit výjimku</button>" +
      "<button type=\"button\" class=\"secondary\" id=\"ex-close\">Zavřít</button>" +
      "</div>";
    modal.hidden = false;
    $("ex-close").onclick = closeModal;
    modal.onclick = function (ev) { if (ev.target === modal) closeModal(); };
    $("ex-type").onchange = function () {
      fillExceptionTimeFields(ped, day, isAbsenceType($("ex-type").value));
    };
    $("ex-clear").onclick = function () {
      delete sheet.exceptions[date];
      persist().then(function () { closeModal(); render(); });
    };
    $("ex-save").onclick = function () {
      var type = $("ex-type").value;
      if (!type) {
        delete sheet.exceptions[date];
      } else if (isAbsenceType(type)) {
        sheet.exceptions[date] = emptyAbsenceException(type, $("ex-note").value, ped);
      } else if (ped) {
        sheet.exceptions[date] = {
          type: type,
          primaStart: $("ex-prima-start").value,
          primaEnd: $("ex-prima-end").value,
          neprimaStart: $("ex-neprima-start").value,
          neprimaEnd: $("ex-neprima-end").value,
          note: $("ex-note").value
        };
      } else {
        var breakStart = $("ex-break-start").value.trim();
        var breakEnd = $("ex-break-end").value.trim();
        var breakMin = 0;
        if (breakStart && breakEnd) {
          breakMin = Math.round(M.hoursBetween(breakStart, breakEnd, 0) * 60);
        }
        sheet.exceptions[date] = {
          type: type,
          start: $("ex-start").value,
          end: $("ex-end").value,
          breakStart: breakStart,
          breakEnd: breakEnd,
          breakMin: breakMin,
          note: $("ex-note").value
        };
      }
      persist().then(function () { closeModal(); render(); });
    };
  }

  function isAbsenceType(type) {
    var meta = M.exceptionMeta(type);
    return !!(meta && meta.countsAs === "absence");
  }

  function emptyAbsenceException(type, note, ped) {
    if (ped) {
      return {
        type: type,
        primaStart: "",
        primaEnd: "",
        neprimaStart: "",
        neprimaEnd: "",
        note: note
      };
    }
    return {
      type: type,
      start: "",
      end: "",
      breakStart: "",
      breakEnd: "",
      breakMin: 0,
      note: note
    };
  }

  function fillExceptionTimeFields(ped, day, absence) {
    var ids = ped
      ? ["ex-prima-start", "ex-prima-end", "ex-neprima-start", "ex-neprima-end"]
      : ["ex-start", "ex-end", "ex-break-start", "ex-break-end"];
    var fromDay = ped
      ? [day.primaStart, day.primaEnd, day.neprimaStart, day.neprimaEnd]
      : [day.start, day.end, day.breakStart || "", day.breakEnd || ""];
    ids.forEach(function (id, i) {
      var el = $(id);
      if (el) el.value = absence ? "" : (fromDay[i] || "");
    });
  }

  function closeModal() {
    var modal = $("modal");
    modal.hidden = true;
    $("modal-card").innerHTML = "";
  }

  function submitSheet() {
    var person = currentPerson();
    var sheet = M.ensureTimesheet(state, person.id, ui.year, ui.month);
    sheet.status = "submitted";
    sheet.submittedAt = new Date().toISOString();
    sheet.returnComment = "";
    persist().then(function () {
      flash("Výkaz odeslán vedení.");
      render();
    });
  }

  function approvePerson(id) {
    var sheet = M.ensureTimesheet(state, id, ui.year, ui.month);
    sheet.status = "approved";
    sheet.approvedAt = new Date().toISOString();
    var who = currentAccount();
    sheet.approvedBy = (who && who.name) || (ui.spUser && ui.spUser.Email) || "správce";
    persist().then(function () { flash("Schváleno."); render(); });
  }

  function returnPerson(id) {
    var comment = window.prompt("Důvod vrácení:");
    if (comment == null) return;
    var sheet = M.ensureTimesheet(state, id, ui.year, ui.month);
    sheet.status = "returned";
    sheet.returnComment = comment;
    persist().then(function () { flash("Vráceno k úpravě."); render(); });
  }

  function saveSettings() {
    state.settings.orgName = $("org-name").value.trim();
    state.settings.schoolYear = $("school-year").value.trim();
    state.settings.cycleStart = $("cycle-start").value;
    state.settings.weeklyHoursFull = Number($("weekly-full").value) || 40;
    persist().then(function () { flash("Nastavení uloženo."); render(); });
  }

  function enterAccount(account, shouldPersist) {
    ui.accountId = account.id;
    ui.loginFails = 0;
    A.saveSession(account.id);
    if (account.personId) ui.personId = account.personId;
    else if (!ui.personId && state.people[0]) ui.personId = state.people[0].id;
    ui.view = (account.role === "admin" && !account.personId) ? "admin" : "timesheet";
    var done = function () {
      flash("Přihlášeno jako " + (account.name || account.login) + ".");
      render();
    };
    if (shouldPersist) persist().then(done);
    else done();
  }

  function doLogin() {
    var now = Date.now();
    if (now < ui.lockUntil) {
      flash("Příliš mnoho pokusů. Zkuste to za chvíli.", true);
      return;
    }
    var login = $("login-name") && $("login-name").value;
    var pass = $("login-pass") && $("login-pass").value;
    var acc = A.findByLogin(state, login);
    if (!acc || acc.active === false) {
      ui.loginFails += 1;
      if (ui.loginFails >= 5) ui.lockUntil = now + 30000;
      flash("Neplatný login nebo heslo.", true);
      return;
    }
    A.verifyPassword(acc, pass).then(function (ok) {
      if (!ok) {
        ui.loginFails += 1;
        if (ui.loginFails >= 5) ui.lockUntil = Date.now() + 30000;
        flash("Neplatný login nebo heslo.", true);
        return;
      }
      enterAccount(acc);
    }).catch(function (err) {
      flash("Přihlášení selhalo: " + err.message, true);
    });
  }

  function setupAdmin() {
    var name = ($("setup-name") && $("setup-name").value || "").trim();
    var login = ($("setup-login") && $("setup-login").value || "").trim().toLowerCase();
    var pass = $("setup-pass") && $("setup-pass").value;
    var pass2 = $("setup-pass2") && $("setup-pass2").value;
    if (!name) { flash("Vyplňte jméno.", true); return; }
    if (!login) login = A.uniqueLogin(state, A.slugLogin(name));
    if (!pass || pass.length < 6) { flash("Heslo musí mít aspoň 6 znaků.", true); return; }
    if (pass !== pass2) { flash("Hesla se neshodují.", true); return; }
    if (A.findByLogin(state, login)) { flash("Tento login už existuje.", true); return; }
    var acc = A.newAccount({ login: login, name: name, role: "admin" });
    A.setPassword(acc, pass).then(function () {
      A.accountsOf(state).push(acc);
      enterAccount(acc, true);
    }).catch(function (err) {
      flash("Účet se nepodařilo vytvořit: " + err.message, true);
    });
  }

  function doLogout() {
    A.clearSession();
    ui.accountId = null;
    ui.view = "timesheet";
    flash("Odhlášeno.");
    render();
  }

  function openChangePassword() {
    $("modal-card").innerHTML =
      "<h3>Změna hesla</h3>" +
      "<label class=\"field\">Současné heslo<input id=\"pw-old\" type=\"password\"></label>" +
      "<label class=\"field\">Nové heslo<input id=\"pw-new\" type=\"password\"></label>" +
      "<label class=\"field\">Nové heslo znovu<input id=\"pw-new2\" type=\"password\"></label>" +
      "<div class=\"row\" style=\"margin-top:12px\">" +
      "<button type=\"button\" id=\"pw-save\">Uložit heslo</button>" +
      "<button type=\"button\" class=\"secondary\" id=\"pw-close\">Zavřít</button></div>";
    $("modal").hidden = false;
    $("pw-close").onclick = closeModal;
    $("pw-save").onclick = saveOwnPassword;
    $("modal").onclick = function (ev) { if (ev.target === $("modal")) closeModal(); };
  }

  function saveOwnPassword() {
    var acc = currentAccount();
    var oldPw = $("pw-old") && $("pw-old").value;
    var nw = $("pw-new") && $("pw-new").value;
    var nw2 = $("pw-new2") && $("pw-new2").value;
    if (!nw || nw.length < 6) { flash("Nové heslo musí mít aspoň 6 znaků.", true); return; }
    if (nw !== nw2) { flash("Nová hesla se neshodují.", true); return; }
    A.verifyPassword(acc, oldPw).then(function (ok) {
      if (!ok) { flash("Současné heslo nesedí.", true); return; }
      return A.setPassword(acc, nw);
    }).then(function (updated) {
      if (!updated) return;
      persist().then(function () {
        closeModal();
        flash("Heslo je změněné.");
      });
    });
  }

  function showIssuedPasswords(rows, title) {
    var body = rows.map(function (r) {
      return "<tr><td>" + esc(r.name) + "</td><td><code>" + esc(r.login) + "</code></td><td><strong>" + esc(r.password) + "</strong></td></tr>";
    }).join("");
    $("modal-card").innerHTML =
      "<h3>" + esc(title || "Nová hesla") + "</h3>" +
      "<p class=\"hint\">Heslo ukažte jen jednou. Zapište ho nebo vytiskněte — později ho z dat nepřečtete, jen vygenerujete nové.</p>" +
      "<div class=\"table-wrap\"><table><thead><tr><th>Jméno</th><th>Login</th><th>Heslo</th></tr></thead><tbody>" +
      body + "</tbody></table></div>" +
      "<div class=\"row\" style=\"margin-top:12px\">" +
      "<button type=\"button\" id=\"pw-print\">Tisk lístků s hesly</button>" +
      "<button type=\"button\" class=\"secondary\" id=\"pw-issued-close\">Zavřít</button></div>";
    $("modal").hidden = false;
    ui.issuedPasswords = rows;
    $("pw-issued-close").onclick = closeModal;
    $("pw-print").onclick = printIssuedPasswords;
    $("modal").onclick = function (ev) { if (ev.target === $("modal")) closeModal(); };
  }

  function printIssuedPasswords() {
    var rows = ui.issuedPasswords || [];
    var html = "<section class=\"card\"><h2>Přihlášení do výkazu práce</h2><p>" +
      esc(state.settings.orgName) + "</p><table><thead><tr><th>Jméno</th><th>Login</th><th>Heslo</th></tr></thead><tbody>" +
      rows.map(function (r) {
        return "<tr><td>" + esc(r.name) + "</td><td>" + esc(r.login) + "</td><td>" + esc(r.password) + "</td></tr>";
      }).join("") + "</tbody></table></section>";
    var root = $("print-root");
    root.hidden = false;
    root.innerHTML = html;
    window.print();
  }

  function issuePasswordForPerson(personId) {
    var person = M.getPerson(state, personId);
    if (!person) return;
    if (A.accountForPerson(state, person.id)) {
      flash("Tato osoba už účet má. Použijte Nové heslo.", true);
      return;
    }
    var login = A.uniqueLogin(state, A.slugLogin(person.name));
    var password = A.generatePassword();
    var acc = A.newAccount({ login: login, name: person.name, role: "user", personId: person.id });
    A.setPassword(acc, password).then(function () {
      A.accountsOf(state).push(acc);
      return persist();
    }).then(function () {
      render();
      showIssuedPasswords([{ name: person.name, login: login, password: password }], "Heslo pro " + person.name);
    });
  }

  function issueMissingPasswords() {
    var missing = state.people.filter(function (p) {
      return p.active && !A.accountForPerson(state, p.id);
    });
    if (!missing.length) {
      flash("Všichni aktivní zaměstnanci už mají heslo.");
      return;
    }
    var issued = [];
    var chain = Promise.resolve();
    missing.forEach(function (person) {
      chain = chain.then(function () {
        var login = A.uniqueLogin(state, A.slugLogin(person.name));
        var password = A.generatePassword();
        var acc = A.newAccount({ login: login, name: person.name, role: "user", personId: person.id });
        return A.setPassword(acc, password).then(function () {
          A.accountsOf(state).push(acc);
          issued.push({ name: person.name, login: login, password: password });
        });
      });
    });
    chain.then(function () { return persist(); }).then(function () {
      render();
      showIssuedPasswords(issued, "Hesla pro " + issued.length + " zaměstnanců");
    });
  }

  function resetAccountPassword(accId) {
    var acc = A.findById(state, accId);
    if (!acc) return;
    var password = A.generatePassword();
    A.setPassword(acc, password).then(function () { return persist(); }).then(function () {
      showIssuedPasswords([{ name: acc.name, login: acc.login, password: password }], "Nové heslo");
    });
  }

  function addAdminAccount() {
    var name = ($("adm-name") && $("adm-name").value || "").trim();
    if (!name) { flash("Vyplňte jméno správce.", true); return; }
    var loginRaw = ($("adm-login") && $("adm-login").value || "").trim();
    var login = A.uniqueLogin(state, loginRaw ? loginRaw.toLowerCase() : A.slugLogin(name));
    var password = A.generatePassword();
    var acc = A.newAccount({ login: login, name: name, role: "admin", personId: null });
    A.setPassword(acc, password).then(function () {
      A.accountsOf(state).push(acc);
      return persist();
    }).then(function () {
      render();
      showIssuedPasswords([{ name: name, login: login, password: password }], "Účet správce");
    });
  }

  function toggleAccount(accId) {
    var acc = A.findById(state, accId);
    if (!acc || acc.id === ui.accountId) return;
    var turningOff = acc.active !== false;
    if (turningOff && acc.role === "admin" && !A.otherActiveAdmin(state, acc.id)) {
      flash("Posledního správce nelze deaktivovat.", true);
      return;
    }
    acc.active = acc.active === false;
    persist().then(function () {
      flash(acc.active === false ? "Účet je neaktivní." : "Účet je znovu aktivní.");
      render();
    });
  }

  function resetAccessToBasicIfNeeded() {
    var lockedOut = A.accountsOf(state).length && !A.hasActiveAdmin(state);
    if (state.accessResetStamp === ACCESS_RESET && !lockedOut) return;
    A.clearSession();
    ui.accountId = null;
    ui.loginFails = 0;
    ui.lockUntil = 0;
    state.accounts = [];
    state.accessResetStamp = ACCESS_RESET;
    try { S.saveLocal(state); } catch (e) { /* private mode */ }
  }

  function addPerson() {
    var name = window.prompt("Jméno a příjmení:");
    if (!name) return;
    var person = {
      id: M.uid(),
      name: name.trim(),
      role: "ucitelka",
      formType: "pedagog",
      email: "",
      fte: 1,
      weeklyHours: 40,
      primaWeekly: 31,
      korekcePrima: 0,
      workplace: "",
      personalNumber: "",
      active: true,
      templates: M.emptyTemplates("ucitelka", "pedagog")
    };
    state.people.push(person);
    ui.editPersonId = person.id;
    if (!ui.personId) ui.personId = person.id;
    persist().then(render);
  }

  function readTemplatesFromForm(ped) {
    var templates = { ranni: {}, odpoledni: {} };
    document.querySelectorAll("[data-tpl]").forEach(function (input) {
      var kind = input.getAttribute("data-tpl");
      var day = input.getAttribute("data-day");
      var k = input.getAttribute("data-k");
      if (!templates[kind][day]) {
        templates[kind][day] = ped ? M.emptyPedSlot() : M.emptyProvozSlot();
      }
      templates[kind][day][k] = k === "breakMin" ? Number(input.value) || 0 : input.value.trim();
    });
    return templates;
  }

  function savePerson() {
    var person = M.getPerson(state, ui.editPersonId);
    if (!person) return;
    person.name = $("p-name").value.trim();
    person.role = $("p-role").value;
    if ($("p-form")) person.formType = $("p-form").value;
    if ($("p-class")) person.classroom = $("p-class").value.trim();
    person.email = $("p-email") ? $("p-email").value.trim() : (person.email || "");
    person.alternatesShifts = $("p-alt") ? $("p-alt").checked : person.alternatesShifts !== false;
    person.workplace = $("p-workplace") ? $("p-workplace").value.trim() : (person.workplace || "");
    person.personalNumber = $("p-personal") ? $("p-personal").value.trim() : (person.personalNumber || "");
    person.fte = Number($("p-fte").value) || 1;
    person.weeklyHours = Number($("p-hours").value) || 40;
    if ($("p-prima")) person.primaWeekly = Number($("p-prima").value) || 0;
    if ($("p-korekce")) person.korekcePrima = Number($("p-korekce").value) || 0;
    person.templates = readTemplatesFromForm(M.isPedPerson(person));
    persist().then(function () { flash("Uloženo."); render(); });
  }

  function pickExcelFiles(personId) {
    ui.excelRestoreId = personId || null;
    var input = $("excel-files");
    if (!input) {
      flash("Chybí výběr souboru Excel.", true);
      return;
    }
    input.value = "";
    input.click();
  }

  function importExcelFiles(fileList) {
    var Imp = window.DochazkaExcelImport;
    if (!Imp) {
      flash("Chybí excel-import.js.", true);
      return;
    }
    Imp.parseFileList(fileList, state.people).then(function (data) {
      applyExcelData(data, ui.excelRestoreId);
      ui.excelRestoreId = null;
    }).catch(function (err) {
      ui.excelRestoreId = null;
      flash(String(err.message || err), true);
    });
  }

  function applyExcelData(data, onlyPersonId) {
    if (!data || !data.people || !data.people.length) {
      flash("V Excelu se nenašli zaměstnanci.", true);
      return;
    }
    window.DOCHAZKA_IMPORT = data;
    if (onlyPersonId) {
      var person = M.getPerson(state, onlyPersonId);
      var src = data.people.filter(function (p) {
        return p.id === onlyPersonId || (person && p.name === person.name);
      })[0];
      if (!person || !src) {
        flash("Tato osoba v nahraném Excelu není.", true);
        return;
      }
      person.templates = M.clone(src.templates);
      person.alternatesShifts = src.alternatesShifts;
      person.formType = src.formType;
      person.hasSchedule = src.hasSchedule;
      persist().then(function () {
        flash("Časy načteny z Excelu (" + person.name + ").");
        render();
      });
      return;
    }
    state.people = data.people;
    state.importStamp = data.stamp;
    if (data.orgName) state.settings.orgName = data.orgName;
    if (data.schoolYear) state.settings.schoolYear = data.schoolYear;
    if (data.cycleStart) state.settings.cycleStart = data.cycleStart;
    state.timesheets = (state.timesheets || []).filter(function (t) {
      return state.people.some(function (p) { return p.id === t.personId; });
    });
    A.pruneAccounts(state);
    if (!state.people.some(function (p) { return p.id === ui.personId; })) {
      ui.personId = (state.people[0] || {}).id || null;
    }
    persist().then(function () {
      flash("Načteno " + state.people.length + " zaměstnanců z Excelu.");
      render();
    });
  }

  function restoreExcelTimes() {
    pickExcelFiles(ui.editPersonId);
  }

  function loadSchoolImport(force) {
    var data = window.DOCHAZKA_IMPORT;
    if (!data || !data.people || !data.people.length) {
      flash("Chybí zamestnanci-data.js. Ve složce aplikace spusťte: python scripts/import-zamestnanci.py", true);
      return false;
    }
    if (!force && state.importStamp === data.stamp && state.people.length && String(state.people[0].id || "").indexOf("demo-") !== 0) {
      return false;
    }
    if (force && state.people.length && !window.confirm("Nahradit seznam zaměstnanců daty z Excelu (Pedagog / Provozní / Extra)?")) {
      return false;
    }
    state.people = data.people;
    state.importStamp = data.stamp;
    if (data.orgName) state.settings.orgName = data.orgName;
    if (data.schoolYear) state.settings.schoolYear = data.schoolYear;
    if (data.cycleStart) state.settings.cycleStart = data.cycleStart;
    state.timesheets = (state.timesheets || []).filter(function (t) {
      return state.people.some(function (p) { return p.id === t.personId; });
    });
    A.pruneAccounts(state);
    ui.personId = (state.people[0] || {}).id || null;
    persist().then(function () {
      flash("Načteno " + state.people.length + " zaměstnanců z Excelu.");
      render();
    });
    return true;
  }

  function loadDemo() {
    if (state.people.length && !window.confirm("Nahradit stávající zaměstnance ukázkou?")) return;
    var keepSettings = state.settings;
    var keepAccounts = (state.accounts || []).filter(function (a) { return a.role === "admin"; });
    state = M.demoState();
    state.settings.orgName = keepSettings.orgName || state.settings.orgName;
    state.settings.schoolYear = keepSettings.schoolYear;
    state.settings.cycleStart = keepSettings.cycleStart;
    state.accounts = keepAccounts;
    ui.personId = state.people[0].id;
    persist().then(function () { flash("Ukázková data načtena."); render(); });
  }

  function exportJson() {
    var blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    download(blob, "dochazka-" + state.settings.schoolYear.replace("/", "-") + ".json");
  }

  function download(blob, name) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  function sheetRows(person) {
    var view = M.buildMonth(state, person, ui.year, ui.month);
    return view.days.map(function (d) {
      return {
        Zaměstnanec: person.name,
        Role: M.roleLabel(person.role),
        Pracoviště: person.workplace || "",
        Osobní_číslo: person.personalNumber || "",
        Datum: d.date,
        Den: d.weekdayShort,
        Přímá_příchod: d.ped ? d.primaStart : d.start,
        Přímá_odchod: d.ped ? d.primaEnd : d.end,
        Nepřímá_příchod: d.ped ? d.neprimaStart : "",
        Nepřímá_odchod: d.ped ? d.neprimaEnd : "",
        Poznámka: (d.exception && d.exception.note) || "",
        Směna: d.holiday ? "" : M.shiftCode(d.shift),
        Důvod_absence: M.absenceLabel(d),
        Stav_výkazu: M.STATUS[view.sheet.status]
      };
    });
  }

  function summaryRow(person) {
    var view = M.buildMonth(state, person, ui.year, ui.month);
    return {
      Zaměstnanec: person.name,
      Role: M.roleLabel(person.role),
      Úvazek: person.fte,
      Měsíc: M.monthName(ui.month) + " " + ui.year,
      Stav: M.STATUS[view.sheet.status],
      Fond_h: view.totals.fond,
      Nepřímá_h: view.totals.neprima,
      Přímá_h: view.totals.prima,
      Odpracováno_h: view.totals.worked,
      Dovolená_h: view.totals.dovolena,
      Nemoc_h: view.totals.nemoc,
      OCR_h: view.totals.ocr,
      MD_h: view.totals.md,
      Přesčas_h: view.totals.prescas,
      Odesláno: view.sheet.submittedAt || "",
      Schválil: view.sheet.approvedBy || ""
    };
  }

  function exportExcel(people, payroll) {
    if (typeof XLSX === "undefined") {
      exportCsv(people, payroll);
      return;
    }
    var wb = XLSX.utils.book_new();
    if (payroll) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(people.map(summaryRow)), "Prehled");
    }
    var days = [];
    people.forEach(function (p) { days = days.concat(sheetRows(p)); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(days), payroll ? "Dny" : "Vykaz");
    if (!payroll) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(people.map(summaryRow)), "Souhrn");
    }
    var name = payroll
      ? "mzdy-" + ui.year + "-" + String(ui.month).padStart(2, "0") + ".xlsx"
      : "vykaz-" + (people[0] && people[0].name || "zamestnanec") + "-" + ui.year + "-" + String(ui.month).padStart(2, "0") + ".xlsx";
    XLSX.writeFile(wb, name);
  }

  function exportCsv(people, payroll) {
    var rows = payroll ? people.map(summaryRow) : sheetRows(people[0]);
    var keys = Object.keys(rows[0] || {});
    var csv = keys.join(";") + "\n" + rows.map(function (r) {
      return keys.map(function (k) { return String(r[k]).replace(/;/g, ","); }).join(";");
    }).join("\n");
    download(new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" }), "vykaz.csv");
  }

  function printHtml(person) {
    var view = M.buildMonth(state, person, ui.year, ui.month);
    var ped = M.isPedPerson(person);
    var body = view.days.map(function (d) {
      if (ped) {
        return "<tr><td>" + esc(M.formatCzechDate(d.dateObj)) + "</td><td>" + esc(d.weekdayShort) + "</td><td>" +
          esc(d.primaStart) + "</td><td>" + esc(d.primaEnd) + "</td><td>" +
          esc(d.neprimaStart) + "</td><td>" + esc(d.neprimaEnd) + "</td><td>" +
          esc((d.exception && d.exception.note) || "") + "</td><td>" +
          esc(d.holiday ? "" : M.shiftCode(d.shift)) + "</td><td>" +
          esc(M.absenceLabel(d)) + "</td></tr>";
      }
      return "<tr><td>" + esc(M.formatCzechDate(d.dateObj)) + "</td><td>" + esc(d.weekdayShort) + "</td><td>" +
        esc(d.start) + "</td><td>" + esc(d.end) + "</td><td>" +
        esc(d.holiday ? "" : M.shiftCode(d.shift)) + "</td><td>" +
        esc(M.absenceLabel(d)) + "</td></tr>";
    }).join("");
    var head = ped
      ? "<tr><th>Datum</th><th>Den</th><th>Přímá od</th><th>Přímá do</th><th>Nepřímá od</th><th>Nepřímá do</th><th>Poznámka</th><th>Směna</th><th>Důvod absence</th></tr>"
      : "<tr><th>Datum</th><th>Den</th><th>Příchod</th><th>Odchod</th><th>Směna</th><th>Důvod absence</th></tr>";
    return "<section class=\"card\" style=\"page-break-after:always\">" +
      "<h2>Výkaz práce — " + esc(person.name) + "</h2>" +
      "<p>" + esc(state.settings.orgName) + " · " + esc(M.roleLabel(person.role)) +
      (person.workplace ? " · " + esc(person.workplace) : "") +
      " · " + esc(M.monthName(ui.month)) + " " + ui.year + " · " + esc(M.STATUS[view.sheet.status]) + "</p>" +
      "<p>Úvazek " + esc(String(person.fte)) +
      " · celkem " + M.formatHours(M.personFond(person)) + " h/týden" +
      (ped ? " · přímá " + M.formatHours(M.personPrimaWeekly(person)) +
        " · nepřímá " + M.formatHours(M.personNeprimaWeekly(person)) : "") +
      " · fond měsíce " + M.formatHours(view.totals.fond) +
      " h · odpracováno " + M.formatHours(view.totals.worked) +
      " h · dovolená " + M.formatHours(view.totals.dovolena) + " h</p>" +
      "<table><thead>" + head + "</thead><tbody>" + body + "</tbody></table></section>";
  }

  function printOne(person) {
    if (!person) return;
    var root = $("print-root");
    root.hidden = false;
    root.innerHTML = printHtml(person);
    window.print();
  }

  function printAllApproved() {
    var html = state.people.filter(function (p) { return p.active; }).map(function (p) {
      var s = M.ensureTimesheet(state, p.id, ui.year, ui.month);
      return s.status === "approved" || s.status === "submitted" ? printHtml(p) : "";
    }).join("");
    var root = $("print-root");
    root.hidden = false;
    root.innerHTML = html || "<p>Žádný odeslaný výkaz.</p>";
    window.print();
  }

  async function detectSharePoint(notify) {
    ui.spWeb = S.spWebUrl();
    if (!ui.spWeb && S.looksLikeSharePoint()) ui.spWeb = location.origin;
    if (!ui.spWeb) {
      if (notify) flash("Nejste na webu SharePoint. Aplikaci nahrajte do knihovny dokumentů školky.", true);
      render();
      return;
    }
    try {
      ui.spUser = await S.currentUser(ui.spWeb);
      ui.storage = "sharepoint";
      matchSharePointUser();
      if (notify) flash("SharePoint web: " + ui.spWeb);
    } catch (e) {
      ui.spWeb = "";
      if (notify) flash("SharePoint API neodpovědělo: " + e.message, true);
    }
    render();
  }

  async function runProvision() {
    try {
      flash("Vytvářím seznamy…");
      await S.provision(ui.spWeb);
      flash("Seznamy DochazkaNastaveni, DochazkaZamestnanci a DochazkaVykazy jsou připravené.");
    } catch (e) {
      flash(e.message, true);
    }
  }

  async function saveSharePoint() {
    try {
      await S.saveToSharePoint(ui.spWeb, state);
      ui.storage = "sharepoint";
      flash("Uloženo do seznamů SharePointu.");
      render();
    } catch (e) {
      flash(e.message, true);
    }
  }

  async function loadSharePoint() {
    try {
      state = await S.loadFromSharePoint(ui.spWeb);
      if (!state.accounts) state.accounts = [];
      ui.storage = "sharepoint";
      restoreSession();
      if (!ui.personId) ui.personId = (state.people[0] || {}).id || null;
      S.saveLocal(state);
      flash("Načteno ze SharePointu.");
      render();
    } catch (e) {
      flash(e.message, true);
    }
  }

  var demoOnly = !state.people.length || state.people.every(function (p) {
    return String(p.id).indexOf("demo-") === 0;
  });
  if (window.DOCHAZKA_IMPORT && (demoOnly || state.importStamp !== window.DOCHAZKA_IMPORT.stamp)) {
    loadSchoolImport(false);
  }

  if (!state.people.length) {
    /* prázdný start — správce načte zaměstnance nebo ukázku */
  } else if (!ui.personId) {
    ui.personId = state.people[0].id;
  }

  render();
  detectSharePoint(false);

  window.addEventListener("afterprint", function () {
    var root = $("print-root");
    root.hidden = true;
    root.innerHTML = "";
  });
})();
