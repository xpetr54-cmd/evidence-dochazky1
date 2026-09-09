/* Přihlášení: hesla jako SHA-256 se solí, účty v datech aplikace. */
(function (global) {
  var SESSION_KEY = "ms-dochazka-session";
  var WORDS = [
    "javor", "lipa", "dub", "buk", "smrk", "vrba", "trnka", "malina",
    "visen", "sliva", "hruska", "jablko", "med", "maslo", "chleb", "syr",
    "mleko", "kava", "caj", "kakao", "klic", "okna", "dvere", "stul",
    "zidle", "lampa", "kniha", "sesit", "tuzka", "mapa", "most", "reka",
    "hora", "louka", "trava", "rosa", "vitr", "snih", "jaro", "leto",
    "podzim", "zima", "jasan", "borovice", "sedmikraska", "kopretina"
  ];

  function cryptoObj() {
    return global.crypto || (typeof require === "function" ? require("crypto").webcrypto : null);
  }

  function toHex(buf) {
    return Array.from(new Uint8Array(buf)).map(function (b) {
      return b.toString(16).padStart(2, "0");
    }).join("");
  }

  function randomBytes(n) {
    var c = cryptoObj();
    var buf = new Uint8Array(n);
    c.getRandomValues(buf);
    return buf;
  }

  function randomSalt() {
    return toHex(randomBytes(16));
  }

  function randomInt(max) {
    var c = cryptoObj();
    var buf = new Uint32Array(1);
    c.getRandomValues(buf);
    return buf[0] % max;
  }

  function generatePassword() {
    var word = WORDS[randomInt(WORDS.length)];
    var num = 1000 + randomInt(9000);
    return word + "-" + num;
  }

  function slugLogin(name) {
    var map = {
      á: "a", č: "c", ď: "d", é: "e", ě: "e", í: "i", ň: "n", ó: "o",
      ř: "r", š: "s", ť: "t", ú: "u", ů: "u", ý: "y", ž: "z",
      Á: "a", Č: "c", Ď: "d", É: "e", Ě: "e", Í: "i", Ň: "n", Ó: "o",
      Ř: "r", Š: "s", Ť: "t", Ú: "u", Ů: "u", Ý: "y", Ž: "z"
    };
    var s = String(name || "").trim();
    var parts = s.split(/\s+/);
    var last = parts[0] || s;
    return last.split("").map(function (ch) { return map[ch] || ch; }).join("")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "")
      .slice(0, 24) || "uzivatel";
  }

  function accountsOf(state) {
    if (!state.accounts) state.accounts = [];
    return state.accounts;
  }

  function uniqueLogin(state, base, exceptId) {
    var login = base || "uzivatel";
    var n = 2;
    function taken(candidate) {
      return accountsOf(state).some(function (a) {
        return a.login === candidate && a.id !== exceptId;
      });
    }
    while (taken(login)) {
      login = base + n;
      n += 1;
    }
    return login;
  }

  async function hashPassword(password, salt) {
    var c = cryptoObj();
    var raw = new TextEncoder().encode(salt + "\n" + password);
    var digest = await c.subtle.digest("SHA-256", raw);
    return toHex(digest);
  }

  async function setPassword(account, password) {
    account.salt = randomSalt();
    account.passwordHash = await hashPassword(password, account.salt);
    account.passwordSetAt = new Date().toISOString();
    return account;
  }

  async function verifyPassword(account, password) {
    if (!account || !account.passwordHash || !account.salt) return false;
    var hash = await hashPassword(password, account.salt);
    return hash === account.passwordHash;
  }

  function newAccount(fields) {
    var Model = global.DochazkaModel;
    return {
      id: Model.uid(),
      login: String(fields.login || "").trim().toLowerCase(),
      name: String(fields.name || "").trim(),
      role: fields.role === "admin" ? "admin" : "user",
      personId: fields.personId || null,
      salt: "",
      passwordHash: "",
      passwordSetAt: null,
      active: true
    };
  }

  function findByLogin(state, login) {
    var key = String(login || "").trim().toLowerCase();
    return accountsOf(state).find(function (a) { return a.login === key; }) || null;
  }

  function findById(state, id) {
    return accountsOf(state).find(function (a) { return a.id === id; }) || null;
  }

  function accountForPerson(state, personId) {
    return accountsOf(state).find(function (a) { return a.personId === personId; }) || null;
  }

  function pruneAccounts(state) {
    var people = state.people || [];
    state.accounts = accountsOf(state).filter(function (a) {
      if (a.role === "admin" && !a.personId) return true;
      if (!a.personId) return false;
      return people.some(function (p) { return p.id === a.personId; });
    });
  }

  function hasActiveAdmin(state) {
    return accountsOf(state).some(function (a) {
      return a.role === "admin" && a.active !== false;
    });
  }

  function otherActiveAdmin(state, exceptId) {
    return accountsOf(state).some(function (a) {
      return a.id !== exceptId && a.role === "admin" && a.active !== false;
    });
  }

  function saveSession(accountId) {
    try {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({
        accountId: accountId,
        at: Date.now()
      }));
    } catch (e) { /* private mode */ }
  }

  function loadSession() {
    try {
      var raw = sessionStorage.getItem(SESSION_KEY);
      if (!raw) return null;
      var data = JSON.parse(raw);
      return data && data.accountId ? data.accountId : null;
    } catch (e) {
      return null;
    }
  }

  function clearSession() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ }
  }

  var api = {
    SESSION_KEY: SESSION_KEY,
    generatePassword: generatePassword,
    slugLogin: slugLogin,
    uniqueLogin: uniqueLogin,
    accountsOf: accountsOf,
    hashPassword: hashPassword,
    setPassword: setPassword,
    verifyPassword: verifyPassword,
    newAccount: newAccount,
    findByLogin: findByLogin,
    findById: findById,
    accountForPerson: accountForPerson,
    pruneAccounts: pruneAccounts,
    hasActiveAdmin: hasActiveAdmin,
    otherActiveAdmin: otherActiveAdmin,
    saveSession: saveSession,
    loadSession: loadSession,
    clearSession: clearSession
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  global.DochazkaAuth = api;
})(typeof window !== "undefined" ? window : global);
