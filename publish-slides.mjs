// Публікація слайдів із slides.js у Firebase (альтернатива кнопці «☁ Опублікувати слайди»).
// Використання:  node publish-slides.mjs <PIN ведучого> [кімната]
// Приклад:       node publish-slides.mjs 12345678 panel
import { readFileSync } from "node:fs";
import vm from "node:vm";

const API_KEY = "AIzaSyA2W8eW7jm9MzJdh-4Zpu72Lrn1c8wUuWs";
const DB_URL  = "https://bilingual-live-presentation-default-rtdb.europe-west1.firebasedatabase.app";
const PRESENTER_EMAIL = "presenter@panel.local";

const pin  = process.argv[2];
const room = (process.argv[3] || "panel").replace(/[^a-zA-Z0-9_-]/g, "");
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

// 3. Записати презентацію
const res = await fetch(`${DB_URL}/rooms/${room}/presentation.json?auth=${auth.idToken}`, {
  method: "PUT", headers: { "Content-Type": "application/json; charset=utf-8" },
  body: JSON.stringify(presentation)
});
if (!res.ok) { console.error("Помилка запису:", res.status, await res.text()); process.exit(1); }
console.log(`OK: ${presentation.slides.length} slides -> room "${room}"`);
