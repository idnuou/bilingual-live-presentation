#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер презентацій PPT / PPTX / PDF  →  slides.js  (+ зашифровані картинки слайдів)
=====================================================================================

Працює на вашому ПК. Firebase отримує лише невеликий slides.js (тексти або ключ і
кількість слайдів). Картинки слайдів шифруються AES-256-GCM і лежать на GitHub Pages:
без PIN-коду (а отже, без ключа з Firebase) їх неможливо переглянути.

Встановлення (один раз):
    pip install -r requirements.txt

Приклади:
    # картинки, дві мови, вивантажити на GitHub і опублікувати в Firebase
    python convert.py --uk доповідь_UA.pptx --en report_EN.pptx --push --publish 29109351

    # лише текст (заголовки та пункти) у slides.js — для подальшого редагування
    python convert.py --uk доповідь_UA.pptx --en report_EN.pdf --mode text

    # одна мова: обидві версії показуватимуть ті самі слайди
    python convert.py --uk доповідь.pdf

    # видалити набір картинок (ID видно в бібліотеці на пульті / у папці decks)
    python convert.py --remove 20261119-034805 --push

Кожна конвертація додає презентацію до БІБЛІОТЕКИ: на пульті ведучого її можна
вибрати зі списку «📚 Презентація» і показати учасникам.

Режими:
    images (типово) — кожен слайд як зображення, дизайн 1:1;
                      назви слайдів (для підказки «Наступний: …» на пульті) береться з тексту
    text            — заголовок + пункти кожного слайда у slides.js, дизайн застосунку
"""
import argparse
import base64
import json
import os
import secrets
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DECKS_DIR = ROOT / "decks"
SLIDES_JS = ROOT / "slides.js"
DEFAULT_SITE = "https://idnuou.github.io/bilingual-live-presentation/"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def die(msg):
    print(f"✖ {msg}", file=sys.stderr)
    sys.exit(1)


def need(module, pip_name):
    try:
        return __import__(module)
    except ImportError:
        die(f"Не встановлено модуль «{module}». Виконайте:  pip install {pip_name}")


# ---------------------------------------------------------------- PowerPoint (Windows)
def powershell(script: str):
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        die("PowerPoint не зміг обробити файл:\n" + (r.stderr or r.stdout).strip())
    return r.stdout


def ps_quote(p) -> str:
    return "'" + str(p).replace("'", "''") + "'"


PS_OPEN = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$wasOpen = $app.Presentations.Count
$p = $app.Presentations.Open({src}, -1, 0, 0)   # ReadOnly, not Untitled, no window
try {{
{body}
}} finally {{
  $p.Close()
  if ($wasOpen -eq 0) {{ $app.Quit() }}           # не закриваємо PowerPoint, якщо ви з ним працюєте
}}
"""


def ppt_to_pptx(src: Path, tmp: Path) -> Path:
    """Старий .ppt → .pptx (для видобування тексту)."""
    out = tmp / (src.stem + ".pptx")
    powershell(PS_OPEN.format(src=ps_quote(src), body=f"  $p.SaveAs({ps_quote(out)}, 24)"))
    return out


def render_pptx(src: Path, outdir: Path, width: int):
    """Кожен (не прихований) слайд → PNG через встановлений PowerPoint."""
    body = f"""
  $w = {width}
  $h = [int][math]::Round($w * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth)
  $n = 0
  foreach ($s in $p.Slides) {{
    if ($s.SlideShowTransition.Hidden -eq -1) {{ continue }}
    $n++
    $s.Export((Join-Path {ps_quote(outdir)} ('{{0:D3}}.png' -f $n)), 'PNG', $w, $h)
  }}"""
    powershell(PS_OPEN.format(src=ps_quote(src), body=body))
    return sorted(outdir.glob("*.png"))


# ---------------------------------------------------------------- PDF
def render_pdf(src: Path, outdir: Path, width: int):
    pymupdf = need("pymupdf", "pymupdf")
    files = []
    with pymupdf.open(src) as doc:
        for i, page in enumerate(doc, 1):
            zoom = width / page.rect.width
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            f = outdir / f"{i:03d}.png"
            pix.save(f)
            files.append(f)
    return files


# ---------------------------------------------------------------- текст
def clean(t: str) -> str:
    return " ".join(str(t).replace("\x0b", " ").split())


def text_from_pptx(src: Path, include_hidden=False):
    need("pptx", "python-pptx")
    from pptx import Presentation
    slides = []
    for sl in Presentation(str(src)).slides:
        if not include_hidden and sl._element.get("show") == "0":
            continue
        title_shape = sl.shapes.title
        title = clean(title_shape.text_frame.text) if title_shape is not None and title_shape.has_text_frame else ""
        items = []
        for sh in sl.shapes:
            if title_shape is not None and sh.shape_id == title_shape.shape_id:
                continue
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    t = clean("".join(r.text for r in para.runs))
                    if t:
                        items.append(t)
            elif getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    t = clean(" | ".join(c.text for c in row.cells))
                    if t:
                        items.append(t)
        if not title and items:
            title = items.pop(0)
        slides.append((title, items))
    return slides


def text_from_pdf(src: Path):
    pymupdf = need("pymupdf", "pymupdf")
    slides = []
    with pymupdf.open(src) as doc:
        for page in doc:
            blocks = sorted(page.get_text("blocks"), key=lambda b: (round(b[1]), b[0]))
            lines = [clean(b[4]) for b in blocks if b[6] == 0 and clean(b[4])]
            title = lines.pop(0) if lines else ""
            slides.append((title, lines))
    return slides


def extract_text(src: Path, tmp: Path):
    ext = src.suffix.lower()
    if ext == ".pptx":
        return text_from_pptx(src)
    if ext == ".ppt":
        return text_from_pptx(ppt_to_pptx(src, tmp))
    if ext == ".pdf":
        return text_from_pdf(src)
    die(f"Непідтримуваний формат: {src.name} (потрібно .ppt, .pptx або .pdf)")


def to_text_slide(title, items, idx):
    title = title or f"Слайд {idx}"
    if items:
        return {"type": "bullets", "title": title, "bullets": items[:12]}
    return {"type": "title", "title": title}


# ---------------------------------------------------------------- картинки
def render(src: Path, outdir: Path, width: int):
    outdir.mkdir(parents=True, exist_ok=True)
    ext = src.suffix.lower()
    if ext in (".ppt", ".pptx"):
        if os.name != "nt":
            die("PPT/PPTX у режимі картинок конвертуються через Microsoft PowerPoint (Windows). Збережіть презентацію як PDF.")
        return render_pptx(src, outdir, width)
    if ext == ".pdf":
        return render_pdf(src, outdir, width)
    die(f"Непідтримуваний формат: {src.name}")


def encrypt_images(pngs, outdir: Path, key: bytes, quality: int):
    need("PIL", "pillow")
    need("cryptography", "cryptography")
    from PIL import Image
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    import io
    aes = AESGCM(key)
    outdir.mkdir(parents=True, exist_ok=True)
    size = 0
    w = h = None
    for i, png in enumerate(pngs, 1):
        with Image.open(png) as im:
            im = im.convert("RGB")
            w, h = im.size
            buf = io.BytesIO()
            im.save(buf, "WEBP", quality=quality, method=6)
        iv = secrets.token_bytes(12)
        blob = iv + aes.encrypt(iv, buf.getvalue(), None)
        (outdir / f"{i:03d}.bin").write_bytes(blob)
        size += len(blob)
    return size, (w, h)


# ---------------------------------------------------------------- git / публікація
def git(*args):
    r = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        die(f"git {' '.join(args)}:\n{(r.stderr or r.stdout).strip()}")
    return r.stdout


def wait_online(url: str, timeout=240):
    print(f"… чекаю, поки GitHub Pages опублікує картинки (до {timeout // 60} хв)")
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            with urllib.request.urlopen(urllib.request.Request(url, method="HEAD"), timeout=10) as r:
                if r.status == 200:
                    print(f"✓ картинки доступні ({int(time.time() - t0)} с)")
                    return True
        except Exception:
            pass
        time.sleep(8)
    print("⚠ GitHub Pages ще не оновився; учасники побачать картинки за хвилину-дві.")
    return False


def write_slides_js(pres: dict):
    head = ("/* Згенеровано convert.py " + time.strftime("%Y-%m-%d %H:%M") + ".\n"
            "   ⚠ НЕ завантажуйте цей файл на хостинг (він у .gitignore): у ньому ключ розшифрування картинок.\n"
            "   Публікація: node publish-slides.mjs <PIN ведучого>  або кнопка на пульті. */\n")
    SLIDES_JS.write_text(head + "window.PRESENTATION = " + json.dumps(pres, ensure_ascii=False, indent=2) + ";\n",
                         encoding="utf-8")


# ---------------------------------------------------------------- головне
def main():
    ap = argparse.ArgumentParser(description="PPT/PPTX/PDF → slides.js (+ зашифровані картинки)")
    ap.add_argument("--uk", type=Path, help="українська версія (.ppt/.pptx/.pdf)")
    ap.add_argument("--en", type=Path, help="англійська версія (якщо немає — показується українська)")
    ap.add_argument("--mode", choices=["images", "text"], default="images")
    ap.add_argument("--title-uk", help="назва презентації українською (типово — назва файлу)")
    ap.add_argument("--title-en", help="назва презентації англійською")
    ap.add_argument("--width", type=int, default=1920, help="ширина картинки, px (типово 1920)")
    ap.add_argument("--quality", type=int, default=82, help="якість WebP 1–100 (типово 82)")
    ap.add_argument("--prune", action="store_true", help="видалити з decks/ усі попередні набори картинок")
    ap.add_argument("--remove", metavar="ID", help="лише видалити набір decks/ID (з --push — і з GitHub) та вийти")
    ap.add_argument("--no-show", action="store_true", help="з --publish: лише додати в бібліотеку, не показувати одразу")
    ap.add_argument("--push", action="store_true", help="git commit + push картинок на GitHub")
    ap.add_argument("--publish", metavar="PIN", help="одразу опублікувати slides.js у Firebase (PIN ведучого)")
    ap.add_argument("--room", default="panel")
    ap.add_argument("--site", default=DEFAULT_SITE, help="адреса сайту на GitHub Pages")
    a = ap.parse_args()

    if a.remove:
        d = DECKS_DIR / a.remove
        if not d.is_dir():
            die(f"Немає набору decks/{a.remove}")
        shutil.rmtree(d)
        print(f"✓ видалено decks/{a.remove}")
        if a.push:
            git("add", "-A", "decks"); git("commit", "-q", "-m", f"Remove slides {a.remove}"); git("push", "-q")
            print("✓ видалено й на GitHub. Не забудьте прибрати її з бібліотеки на пульті (🗑).")
        return
    if not a.uk:
        die("Вкажіть презентацію: --uk файл.pptx|.pdf")

    for f in (a.uk, a.en):
        if f and not f.is_file():
            die(f"Файл не знайдено: {f}")
    srcs = {"uk": a.uk.resolve(), "en": a.en.resolve() if a.en else None}
    deck = {"uk": {"title": a.title_uk or a.uk.stem},
            "en": {"title": a.title_en or (a.en.stem if a.en else a.title_uk or a.uk.stem)}}

    with tempfile.TemporaryDirectory() as tmpd:
        tmp = Path(tmpd)

        # ---- текст (потрібен в обох режимах: у images — для назв слайдів на пульті)
        texts = {}
        for lang, src in srcs.items():
            if src:
                print(f"• текст [{lang}] ← {src.name}")
                try:
                    texts[lang] = extract_text(src, tmp)
                except SystemExit:
                    raise
                except Exception as e:
                    if a.mode == "text":
                        die(f"Не вдалося видобути текст із {src.name}: {e}")
                    print(f"  ⚠ текст не видобуто ({e}); назви слайдів будуть «Слайд N»")
                    texts[lang] = []

        if a.mode == "text":
            uk, en = texts["uk"], texts.get("en") or texts["uk"]
            if len(uk) != len(en):
                print(f"⚠ різна кількість слайдів: UA {len(uk)}, EN {len(en)} — зайві буде доповнено копією")
            n = max(len(uk), len(en))
            slides = []
            for i in range(n):
                u = to_text_slide(*(uk[i] if i < len(uk) else en[i]), i + 1)
                e = to_text_slide(*(en[i] if i < len(en) else uk[i]), i + 1)
                typ = "bullets" if "bullets" in u or "bullets" in e else "title"
                for s in (u, e):
                    s.pop("type", None)
                    if typ == "bullets":
                        s.setdefault("bullets", [])
                slides.append({"type": typ, "uk": u, "en": e})
            pres = {"id": "text-" + time.strftime("%Y%m%d-%H%M%S"), "deck": deck, "slides": slides}
            write_slides_js(pres)
            print(f"✓ slides.js: {n} текстових слайдів. Перевірте й за потреби відредагуйте тексти.")

        else:
            # ---- картинки
            deck_id = time.strftime("%Y%m%d-%H%M%S")
            target = DECKS_DIR / deck_id
            key = secrets.token_bytes(32)
            counts, sizes, dims = {}, 0, None
            for lang, src in srcs.items():
                if not src:
                    continue
                print(f"• рендер [{lang}] ← {src.name}")
                pngs = render(src, tmp / lang, a.width)
                if not pngs:
                    die(f"У {src.name} не знайдено слайдів")
                sz, dims = encrypt_images(pngs, target / lang, key, a.quality)
                counts[lang] = len(pngs)
                sizes += sz
                print(f"  {len(pngs)} слайдів, {sz / 1024 / 1024:.1f} МБ (зашифровано)")
            if a.en and counts["uk"] != counts["en"]:
                shutil.rmtree(target, ignore_errors=True)
                die(f"Різна кількість слайдів: UA {counts['uk']}, EN {counts['en']}. "
                    "Для синхронного показу вони мають збігатися (приховані слайди не враховуються).")
            n = counts["uk"]
            langs = {"uk": "uk", "en": "en" if a.en else "uk"}

            def title_of(lang, i):
                t = texts.get(lang) or texts.get("uk") or []
                return (t[i][0] if i < len(t) and t[i][0] else f"Слайд {i + 1}")[:120]

            slides = [{"type": "image", "uk": {"title": title_of("uk", i)}, "en": {"title": title_of("en", i)}}
                      for i in range(n)]
            pres = {"id": deck_id, "deck": deck, "type": "images", "base": f"decks/{deck_id}/", "langs": langs,
                    "aspect": round(dims[1] / dims[0], 5), "key": base64.b64encode(key).decode(), "slides": slides}
            write_slides_js(pres)

            if a.prune and DECKS_DIR.exists():
                for old in DECKS_DIR.iterdir():
                    if old.is_dir() and old.name != deck_id:
                        shutil.rmtree(old, ignore_errors=True)
                        print(f"  видалено старий набір decks/{old.name}")
            print(f"✓ decks/{deck_id}/: {n} слайдів × {len(counts)} мов(и), {sizes / 1024 / 1024:.1f} МБ")
            print("✓ slides.js оновлено (ключ розшифрування всередині)")

            if a.push:
                git("add", "-A", "decks")
                if git("status", "--porcelain", "decks").strip():
                    git("commit", "-q", "-m", f"Slides {deck_id}: {deck['uk']['title']}")
                    git("push", "-q")
                    print("✓ картинки вивантажено на GitHub")
                if a.publish:
                    wait_online(a.site.rstrip("/") + f"/decks/{deck_id}/uk/001.bin")
            elif a.publish:
                print("⚠ без --push картинки не потраплять на сайт; учасники їх не побачать")

    if a.publish:
        r = subprocess.run(["node", str(ROOT / "publish-slides.mjs"), a.publish, a.room] + (["--no-show"] if a.no_show else []), cwd=ROOT)
        if r.returncode != 0:
            die("Публікацію у Firebase не виконано")
    else:
        print("Далі: node publish-slides.mjs <PIN ведучого>   (або «📂 Обрати slides.js» на пульті)")


if __name__ == "__main__":
    main()
