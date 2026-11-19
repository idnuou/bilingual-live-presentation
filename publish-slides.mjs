// Додає slides.js до бібліотеки презентацій у Firebase і одразу показує її
// (альтернатива кнопці «☁ Додати в бібліотеку й показати» на пульті).
// Використання:  node publish-slides.mjs <PIN ведучого> [кімната] [--no-show]
// Приклад:       node publish-slides.mjs 12345678 panel
//   --no-show  — лише додати до бібліотеки, не перемикаючи поточний показ
import { readFileSync } from "node:fs";
import vm from "node:vm";

const API_KEY = "AIzaSyA2W8eW7jm9MzJdh-4Zpu72Lrn1c8wUuWs";
const DB_URL  = "https://bilingual-live-presentation-default-rtdb.europe-west1.firebasedatabase.app";
const PRESENTER_EMAIL = "presenter@panel.local";

const args = process.argv.slice(2).filter(a => !a.startsWith("--"));
const NO_SHOW = process.argv.includes("--no-show");
const pin  = args[0];
const room = (args[1] || "panel").replace(/[^a-zA-Z0-9_-]/g, "");
// версія застосунку з index.html: старі кешовані сторінки учасників самі перезавантажаться
let APP_VERSION = 0;
try { APP_VERSION = +(readFileSync(new URL("./index.html", import.meta.url), "utf8").match(/const APP_VERSION = (\d+)/) || [])[1] || 0; } catch {}
if (!pin) { console.error("Вкажіть PIN ведучого: node publish-slides.mjs <PIN> [кімната]"); process.exit(1); }

// 1. Прочитати slides.js (UTF-8)
const sandbox = { window: {} };
vm.runInNewContext(readFileSync(new URL("./slides.js", import.meta.url), "utf8"), sandbox);
const presentation = sandbox.window.PRESENTATION;
if (!presentation || !Array.isArray(presentation.slides)) { console.error("У slides.js не знайдено window.PRESENTATION.slides"); process.exit(1); }

// 2. Увійти як ведучий
const auth = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${API_KEY}`, {
  method: "POST", headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ email: PRESENTER_EMAIL, password: pin, returnSecureToken: true })
}).then(r => r.json());
if (!auth.idToken) { console.error("Помилка входу:", auth.error?.message || auth); process.exit(1); }

// 3. Записати в бібліотеку (і показати)
const id = String(presentation.id || ("d" + Date.now())).replace(/[.#$\[\]\/]/g, "_");
const pres = { ...presentation, id, appMin: APP_VERSION };
const patch = { [`library/${id}`]: { ...pres, addedAt: { ".sv": "timestamp" } } };
if (!NO_SHOW) {
  patch[`rooms/${room}/presentation`] = pres;
  patch[`rooms/${room}/state`] = { slide: 0, started: false, at: Date.now() };
}
const res = await fetch(`${DB_URL}/.json?auth=${auth.idToken}`, {
  method: "PATCH", headers: { "Content-Type": "application/json; charset=utf-8" },
  body: JSON.stringify(patch)
});
if (!res.ok) { console.error("Помилка запису:", res.status, await res.text()); process.exit(1); }
console.log(`OK: "${presentation.deck?.uk?.title || id}" (${presentation.slides.length} slides) -> library` + (NO_SHOW ? "" : `, shown in room "${room}"`));
