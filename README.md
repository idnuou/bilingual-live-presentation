# Двомовна презентація наживо · Bilingual live presentation

Учасники відкривають сторінку, вводять PIN, обирають мову (UA/EN) і бачать слайди, які гортає ведучий.
Ведучий відкриває ту саму сторінку і вводить PIN ведучого — відкривається пульт керування.

| | |
|---|---|
| Сайт (учасники й ведучий) | https://idnuou.github.io/bilingual-live-presentation/ |
| Firebase-проєкт | `bilingual-live-presentation` (акаунт id.nuou@gmail.com), Realtime Database у europe-west1 |
| Службові записи Authentication | `participants@panel.local` (пароль = PIN учасника), `presenter@panel.local` (пароль = PIN ведучого) |

PIN-коди в репозиторії **не зберігаються**.

## Файли

- `index.html` — застосунок (публікується на GitHub Pages)
- `slides.js` — тексти слайдів UA/EN, **не публікується** (у `.gitignore`)
- `publish-slides.mjs` — публікація `slides.js` у Firebase з командного рядка
- `database.rules.json` — правила доступу Realtime Database
- `ІНСТРУКЦІЯ.md` — покрокова інструкція

## Оновити слайди

1. Відредагуйте `slides.js`.
2. Опублікуйте їх одним зі способів:
   - `node publish-slides.mjs <PIN ведучого>` у цій папці;
   - або на сайті: PIN ведучого → «📂 Обрати slides.js» → «☁ Опублікувати слайди».

Учасники отримають зміни миттєво. Сайт при цьому перевикладати не потрібно.
