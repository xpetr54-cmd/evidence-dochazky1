/* Úložiště: prohlížeč (localStorage) nebo SharePoint seznamy na stejném webu. */
(function (global) {
  var Model = global.DochazkaModel;
  var KEY = "ms-dochazka-v1";
  var LIST = {
    settings: "DochazkaNastaveni",
    people: "DochazkaZamestnanci",
    timesheets: "DochazkaVykazy"
  };

  function spWebUrl() {
    if (global._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) {
      return _spPageContextInfo.webAbsoluteUrl;
    }
    var path = location.pathname;
    var markers = ["/Shared Documents/", "/Documents/", "/SiteAssets/", "/sites/"];
    var lower = path.toLowerCase();
    var cut = -1;
    ["/shared documents/", "/documents/", "/siteassets/"].forEach(function (m) {
      var i = lower.indexOf(m);
      if (i > 0 && (cut < 0 || i < cut)) cut = i;
    });
    if (cut > 0) return location.origin + path.slice(0, cut);
    var m = path.match(/^(\/sites\/[^/]+)/i);
    if (m) return location.origin + m[1];
    return "";
  }

  function looksLikeSharePoint() {
    if (global._spPageContextInfo) return true;
    var host = location.hostname || "";
    return /sharepoint\.com$/i.test(host) || host.indexOf(".sharepoint.") !== -1;
  }

  function jsonHeaders(digest) {
    var h = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose"
    };
    if (digest) h["X-RequestDigest"] = digest;
    return h;
  }

  async function spFetch(web, path, options) {
    options = options || {};
    var res = await fetch(web + path, options);
    var text = await res.text();
    var data = text ? JSON.parse(text) : {};
    if (!res.ok) {
      var msg = (data && data.error && data.error.message && data.error.message.value) || res.statusText;
      throw new Error(msg || ("SharePoint " + res.status));
    }
    return data;
  }

  async function getDigest(web) {
    var data = await spFetch(web, "/_api/contextinfo", {
      method: "POST",
      headers: jsonHeaders()
    });
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  async function currentUser(web) {
    var data = await spFetch(web, "/_api/web/currentuser", { headers: jsonHeaders() });
    return data.d;
  }

  async function listExists(web, title) {
    try {
      await spFetch(web, "/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')", {
        headers: jsonHeaders()
      });
      return true;
    } catch (e) {
      return false;
    }
  }

  async function createList(web, digest, title, description) {
    await spFetch(web, "/_api/web/lists", {
      method: "POST",
      headers: jsonHeaders(digest),
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        BaseTemplate: 100,
        Title: title,
        Description: description
      })
    });
  }

  async function addField(web, digest, list, payload) {
    await spFetch(web, "/_api/web/lists/getbytitle('" + list + "')/fields", {
      method: "POST",
      headers: jsonHeaders(digest),
      body: JSON.stringify(payload)
    });
  }

  async function provision(web) {
    var digest = await getDigest(web);
    if (!(await listExists(web, LIST.settings))) {
      await createList(web, digest, LIST.settings, "Nastavení školního roku a čtrnáctidenního cyklu");
      digest = await getDigest(web);
      await addField(web, digest, LIST.settings, fieldText("OrgName"));
      await addField(web, digest, LIST.settings, fieldText("SkolniRok"));
      await addField(web, digest, LIST.settings, fieldText("ZacatekCyklu"));
      await addField(web, digest, LIST.settings, fieldNumber("TydenniFond"));
      await addField(web, digest, LIST.settings, fieldNote("AdminEmails"));
      await addField(web, digest, LIST.settings, fieldNote("LeadershipEmails"));
      await addField(web, digest, LIST.settings, fieldNote("AccountsJson"));
    } else {
      try {
        digest = await getDigest(web);
        await addField(web, digest, LIST.settings, fieldNote("AccountsJson"));
      } catch (e) { /* pole už může existovat */ }
    }
    digest = await getDigest(web);
    if (!(await listExists(web, LIST.people))) {
      await createList(web, digest, LIST.people, "Zaměstnanci a šablony směn");
      digest = await getDigest(web);
      await addField(web, digest, LIST.people, fieldChoice("Role", ["ucitelka", "asistentka", "provoz", "administrativa"]));
      await addField(web, digest, LIST.people, fieldText("Email"));
      await addField(web, digest, LIST.people, fieldNumber("Uvazek"));
      await addField(web, digest, LIST.people, fieldNumber("TydenniHodin"));
      await addField(web, digest, LIST.people, fieldBool("Aktivni"));
      await addField(web, digest, LIST.people, fieldNote("SablonyJson"));
      await addField(web, digest, LIST.people, fieldText("PersonId"));
    }
    digest = await getDigest(web);
    if (!(await listExists(web, LIST.timesheets))) {
      await createList(web, digest, LIST.timesheets, "Měsíční výkazy práce");
      digest = await getDigest(web);
      await addField(web, digest, LIST.timesheets, fieldText("PersonId"));
      await addField(web, digest, LIST.timesheets, fieldNumber("Rok"));
      await addField(web, digest, LIST.timesheets, fieldNumber("Mesic"));
      await addField(web, digest, LIST.timesheets, fieldChoice("Stav", ["draft", "submitted", "approved", "returned"]));
      await addField(web, digest, LIST.timesheets, fieldNote("TydnyJson"));
      await addField(web, digest, LIST.timesheets, fieldNote("VyjimkyJson"));
      await addField(web, digest, LIST.timesheets, fieldText("Odeslano"));
      await addField(web, digest, LIST.timesheets, fieldText("Schvaleno"));
      await addField(web, digest, LIST.timesheets, fieldText("Schvalil"));
      await addField(web, digest, LIST.timesheets, fieldNote("Komentar"));
      await addField(web, digest, LIST.timesheets, fieldText("SheetId"));
    }
    return { ok: true, web: web };
  }

  function fieldText(name) {
    return { __metadata: { type: "SP.FieldText" }, FieldTypeKind: 2, Title: name, MaxLength: 255 };
  }
  function fieldNote(name) {
    return { __metadata: { type: "SP.FieldMultiLineText" }, FieldTypeKind: 3, Title: name, NumberOfLines: 12, RichText: false };
  }
  function fieldNumber(name) {
    return { __metadata: { type: "SP.FieldNumber" }, FieldTypeKind: 9, Title: name };
  }
  function fieldBool(name) {
    return { __metadata: { type: "SP.Field" }, FieldTypeKind: 8, Title: name };
  }
  function fieldChoice(name, choices) {
    return {
      __metadata: { type: "SP.FieldChoice" },
      FieldTypeKind: 6,
      Title: name,
      Choices: { results: choices }
    };
  }

  async function getItems(web, list) {
    var data = await spFetch(
      web,
      "/_api/web/lists/getbytitle('" + list + "')/items?$top=5000",
      { headers: jsonHeaders() }
    );
    return (data.d && data.d.results) || [];
  }

  async function postItem(web, digest, list, payload) {
    return spFetch(web, "/_api/web/lists/getbytitle('" + list + "')/items", {
      method: "POST",
      headers: jsonHeaders(digest),
      body: JSON.stringify(payload)
    });
  }

  async function mergeItem(web, digest, list, id, payload) {
    var headers = jsonHeaders(digest);
    headers["IF-MATCH"] = "*";
    headers["X-HTTP-Method"] = "MERGE";
    return spFetch(web, "/_api/web/lists/getbytitle('" + list + "')/items(" + id + ")", {
      method: "POST",
      headers: headers,
      body: JSON.stringify(payload)
    });
  }

  async function deleteItem(web, digest, list, id) {
    var headers = jsonHeaders(digest);
    headers["IF-MATCH"] = "*";
    headers["X-HTTP-Method"] = "DELETE";
    return spFetch(web, "/_api/web/lists/getbytitle('" + list + "')/items(" + id + ")", {
      method: "POST",
      headers: headers
    });
  }

  function parseJson(s, fallback) {
    if (!s) return fallback;
    try { return JSON.parse(s); } catch (e) { return fallback; }
  }

  async function loadFromSharePoint(web) {
    var settingsItems = await getItems(web, LIST.settings);
    var peopleItems = await getItems(web, LIST.people);
    var sheetItems = await getItems(web, LIST.timesheets);
    var state = Model.emptyState();
    var s = settingsItems[0];
    if (s) {
      state.settings.orgName = s.OrgName || state.settings.orgName;
      state.settings.schoolYear = s.SkolniRok || state.settings.schoolYear;
      state.settings.cycleStart = s.ZacatekCyklu || state.settings.cycleStart;
      state.settings.weeklyHoursFull = s.TydenniFond || state.settings.weeklyHoursFull;
      state.settings.adminEmails = s.AdminEmails || "";
      state.settings.leadershipEmails = s.LeadershipEmails || "";
      state.accounts = parseJson(s.AccountsJson, []);
      state._spSettingsId = s.Id;
    }
    state.people = peopleItems.map(function (p) {
      return {
        id: p.PersonId || ("sp-" + p.Id),
        name: p.Title || "",
        role: p.Role || "ucitelka",
        email: p.Email || "",
        fte: p.Uvazek != null ? Number(p.Uvazek) : 1,
        weeklyHours: p.TydenniHodin != null ? Number(p.TydenniHodin) : 40,
        active: p.Aktivni !== false,
        templates: parseJson(p.SablonyJson, Model.templatesForRole(p.Role || "ucitelka")),
        _spId: p.Id
      };
    });
    state.timesheets = sheetItems.map(function (t) {
      return {
        id: t.SheetId || ("sp-ts-" + t.Id),
        personId: t.PersonId,
        year: Number(t.Rok),
        month: Number(t.Mesic),
        status: t.Stav || "draft",
        weekShifts: parseJson(t.TydnyJson, {}),
        exceptions: parseJson(t.VyjimkyJson, {}),
        submittedAt: t.Odeslano || null,
        approvedAt: t.Schvaleno || null,
        approvedBy: t.Schvalil || null,
        returnComment: t.Komentar || "",
        _spId: t.Id
      };
    });
    return state;
  }

  async function saveToSharePoint(web, state) {
    var digest = await getDigest(web);
    var settingsPayload = {
      __metadata: { type: "SP.Data.DochazkaNastaveniListItem" },
      Title: "Aktualni",
      OrgName: state.settings.orgName,
      SkolniRok: state.settings.schoolYear,
      ZacatekCyklu: state.settings.cycleStart,
      TydenniFond: Number(state.settings.weeklyHoursFull) || 40,
      AdminEmails: state.settings.adminEmails || "",
      LeadershipEmails: state.settings.leadershipEmails || "",
      AccountsJson: JSON.stringify(state.accounts || [])
    };
    if (state._spSettingsId) {
      await mergeItem(web, digest, LIST.settings, state._spSettingsId, settingsPayload);
    } else {
      var created = await postItem(web, digest, LIST.settings, settingsPayload);
      state._spSettingsId = created.d.Id;
    }

    var existingPeople = await getItems(web, LIST.people);
    var keepPeople = {};
    for (var i = 0; i < state.people.length; i++) {
      var p = state.people[i];
      var payload = {
        __metadata: { type: "SP.Data.DochazkaZamestnanciListItem" },
        Title: p.name,
        Role: p.role,
        Email: p.email || "",
        Uvazek: Number(p.fte) || 1,
        TydenniHodin: Number(p.weeklyHours) || 40,
        Aktivni: !!p.active,
        SablonyJson: JSON.stringify(p.templates || {}),
        PersonId: p.id
      };
      digest = await getDigest(web);
      if (p._spId) {
        await mergeItem(web, digest, LIST.people, p._spId, payload);
        keepPeople[p._spId] = true;
      } else {
        var saved = await postItem(web, digest, LIST.people, payload);
        p._spId = saved.d.Id;
        keepPeople[p._spId] = true;
      }
    }
    for (var j = 0; j < existingPeople.length; j++) {
      if (!keepPeople[existingPeople[j].Id]) {
        digest = await getDigest(web);
        await deleteItem(web, digest, LIST.people, existingPeople[j].Id);
      }
    }

    var existingSheets = await getItems(web, LIST.timesheets);
    var keepSheets = {};
    for (var k = 0; k < state.timesheets.length; k++) {
      var t = state.timesheets[k];
      var person = Model.getPerson(state, t.personId);
      var row = {
        __metadata: { type: "SP.Data.DochazkaVykazyListItem" },
        Title: (person ? person.name : t.personId) + " " + t.year + "-" + String(t.month).padStart(2, "0"),
        PersonId: t.personId,
        Rok: t.year,
        Mesic: t.month,
        Stav: t.status,
        TydnyJson: JSON.stringify(t.weekShifts || {}),
        VyjimkyJson: JSON.stringify(t.exceptions || {}),
        Odeslano: t.submittedAt || "",
        Schvaleno: t.approvedAt || "",
        Schvalil: t.approvedBy || "",
        Komentar: t.returnComment || "",
        SheetId: t.id
      };
      digest = await getDigest(web);
      if (t._spId) {
        await mergeItem(web, digest, LIST.timesheets, t._spId, row);
        keepSheets[t._spId] = true;
      } else {
        var savedSheet = await postItem(web, digest, LIST.timesheets, row);
        t._spId = savedSheet.d.Id;
        keepSheets[t._spId] = true;
      }
    }
    for (var n = 0; n < existingSheets.length; n++) {
      if (!keepSheets[existingSheets[n].Id]) {
        digest = await getDigest(web);
        await deleteItem(web, digest, LIST.timesheets, existingSheets[n].Id);
      }
    }
  }

  function loadLocal() {
    try {
      var raw = localStorage.getItem(KEY);
      if (!raw) return null;
      var state = JSON.parse(raw);
      if (!state || state.version !== 1) return null;
      if (!state.accounts) state.accounts = [];
      return state;
    } catch (e) {
      return null;
    }
  }

  function saveLocal(state) {
    var copy = Model.clone(state);
    delete copy._spSettingsId;
    localStorage.setItem(KEY, JSON.stringify(copy));
  }

  var api = {
    KEY: KEY,
    LIST: LIST,
    spWebUrl: spWebUrl,
    looksLikeSharePoint: looksLikeSharePoint,
    currentUser: currentUser,
    provision: provision,
    loadFromSharePoint: loadFromSharePoint,
    saveToSharePoint: saveToSharePoint,
    listExists: listExists,
    loadLocal: loadLocal,
    saveLocal: saveLocal
  };

  global.DochazkaStorage = api;
})(typeof window !== "undefined" ? window : global);
