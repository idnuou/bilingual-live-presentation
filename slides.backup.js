/* =====================================================================
   СЛАЙДИ ПРЕЗЕНТАЦІЇ / PRESENTATION SLIDES  (пробний набір, 4 слайди)
   ---------------------------------------------------------------------
   ⚠ Цей файл НЕ завантажуйте на GitHub / хостинг (він у .gitignore).
   Ведучий: увійти PIN-кодом ведучого → «📂 Обрати slides.js» →
   «☁ Опублікувати слайди».

   type: "title"   — kicker, title, lead
   type: "bullets" — kicker, title, bullets[]
   ===================================================================== */
window.PRESENTATION = {
  deck: {
    uk: { title: "Пробна презентація" },
    en: { title: "Test presentation" }
  },
  slides: [
    { type:"title",
      uk:{ kicker:"Тест", title:"Двомовна презентація наживо", lead:"Перевірка синхронного показу слайдів" },
      en:{ kicker:"Test", title:"Bilingual live presentation", lead:"Checking synchronized slide display" } },

    { type:"bullets",
      uk:{ kicker:"Як це працює", title:"Три кроки для учасника",
           bullets:["Скануйте QR-код","Введіть PIN-код","Оберіть мову — і слідкуйте за доповіддю"] },
      en:{ kicker:"How it works", title:"Three steps for participants",
           bullets:["Scan the QR code","Enter the PIN","Choose your language and follow the talk"] } },

    { type:"bullets",
      uk:{ kicker:"Можливості", title:"Що вміє застосунок",
           bullets:["Слайди гортаються на всіх пристроях одночасно","Кожен обирає українську або англійську","Можна переглянути вже показані слайди"] },
      en:{ kicker:"Features", title:"What the app can do",
           bullets:["Slides change on all devices at once","Everyone chooses Ukrainian or English","Previously shown slides can be reviewed"] } },

    { type:"title",
      uk:{ kicker:"Дякуємо", title:"Тест завершено", lead:"Запитання та обговорення" },
      en:{ kicker:"Thank you", title:"Test complete", lead:"Questions and discussion" } }
  ]
};
