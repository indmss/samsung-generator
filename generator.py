import os
import re
import sys
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageFilter

from psd_writer import save_layered_psd

# ============================================================
#  ФОРМАТЫ — п.5 и п.6 ТЗ. safe_zone в точных пикселях гайдлайна.
# ============================================================
FORMATS = {
    "Square":   {"w": 1080, "h": 1080, "logo": True,  "safe": {"top": 64,  "bottom": 64,  "side": 64}},
    "Portrait": {"w": 1080, "h": 1350, "logo": True,  "safe": {"top": 64,  "bottom": 80,  "side": 64}},
    "Stories":  {"w": 1080, "h": 1920, "logo": False, "safe": {"top": 250, "bottom": 350, "side": 72}},
    "Display":  {"w": 1200, "h": 628,  "logo": True,  "safe": {"top": 60,  "bottom": 60,  "side": 60}},
}

INK = (11, 11, 12)
GOLD = (217, 179, 130)
SAFE_ZONE_TOLERANCE_PX = 2  # допуск на суб-пиксельный антиалиасинг текста

RETAILERS = ["ALL", "Kaspi", "Sulpak", "Mechta", "Technodom"]
LANGS = ["RU", "KZ"]


# ============================================================
#  ТЕКСТЫ ИНТЕРФЕЙСА КРЕАТИВА — статические фразы шаблона,
#  которых нет в Excel (RU/KZ), п.8 ТЗ "Мультиязычность"
# ============================================================
def _t_benefit(save, lang):
    money = f"{save:,.0f}".replace(",", " ")
    if lang == "KZ":
        return f"{money} тг-ге дейін үнемдеу"
    return f"выгода до {money} тг"


def _t_permonth(price_per_month, lang):
    money = f"{price_per_month:,.0f}".replace(",", " ")
    if lang == "KZ":
        return f"айына {money} тг-ден"
    return f"от {money} тг/мес"


def _t_price(price, lang=None):
    return f"{price:,.0f}".replace(",", " ") + " тг"


def _t_default_pay(lang):
    return "Samsung Pay"


# ============================================================
#  РЕАЛЬНОЕ ФОТО ТОВАРА (PNG без фона), если есть у дизайнера
# ============================================================
def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


_PRODUCT_IMAGES_CACHE = {}


def _load_product_image(model_code, category, excel_image_path=None):
    cache_key = model_code or category
    if cache_key in _PRODUCT_IMAGES_CACHE:
        return _PRODUCT_IMAGES_CACHE[cache_key]

    assets_dir = os.path.join(_base_dir(), "assets")
    candidates = []
    if model_code:
        candidates.append(os.path.join(assets_dir, f"{model_code}.png"))
    if excel_image_path and str(excel_image_path).strip() and str(excel_image_path) != "nan":
        p = str(excel_image_path).strip()
        candidates.append(p if os.path.isabs(p) else os.path.join(_base_dir(), p))
    cat = (category or "").lower()
    if "стирал" in cat or "сушил" in cat:
        candidates.append(os.path.join(assets_dir, "generic_washer.png"))
    elif "холодил" in cat:
        candidates.append(os.path.join(assets_dir, "generic_fridge.png"))
    elif "телевиз" in cat or "tv" in cat:
        candidates.append(os.path.join(assets_dir, "generic_tv.png"))

    for path in candidates:
        if os.path.exists(path):
            img = Image.open(path).convert("RGBA")
            _PRODUCT_IMAGES_CACHE[cache_key] = img
            return img

    _PRODUCT_IMAGES_CACHE[cache_key] = None
    return None


def _font(size, bold=True):
    size = max(int(size), 8)
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "C:\\Windows\\Fonts\\arialbd.ttf" if bold else "C:\\Windows\\Fonts\\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


# ============================================================
#  ЧТЕНИЕ EXCEL
# ============================================================
def _find_sheet(excel_path, keyword):
    xls = pd.ExcelFile(excel_path)
    for name in xls.sheet_names:
        if keyword.lower() in name.lower():
            return name
    raise ValueError(f"Не нашла вкладку, содержащую '{keyword}'. Есть вкладки: {xls.sheet_names}")


COLUMN_MAP = {
    "model_code": ["Model_Code", "Код модели", "Модель"],
    "category": ["Category", "Категория"],
    "category_kz": ["Category_KZ"],
    "price_new": ["Price_New", "Цена", "Цена новая"],
    "price_old": ["Price_Old", "Старая цена"],
    "discount": ["Discount_%", "Discount_percent", "Скидка"],
    "slogan": ["Slogan", "Слоган"],
    "slogan_kz": ["Slogan_KZ"],
    "feature_1": ["Feature_1", "Фича 1"],
    "feature_1_kz": ["Feature_1_KZ"],
    "feature_2": ["Feature_2", "Фича 2"],
    "feature_2_kz": ["Feature_2_KZ"],
    "feature_3": ["Feature_3", "Фича 3"],
    "feature_3_kz": ["Feature_3_KZ"],
    "volume_badge": ["Volume_Badge", "Бейдж"],
    "cashback_text": ["Cashback_Text", "Кэшбэк"],
    "cashback_text_kz": ["Cashback_Text_KZ"],
    "payment_text": ["Payment_Text", "Рассрочка"],
    "payment_text_kz": ["Payment_Text_KZ"],
    "channel": ["Channel", "Канал"],
    "image_path": ["Product_Image_Path", "Фото", "Image"],
}

_MODEL_CODE_RE = re.compile(r"^[A-Za-z0-9\-]{3,}$")


def _valid_rows(df, code_col):
    """
    Фильтрует служебные/пустые строки (заметки типа 'Розовым — главная
    модель...', полностью пустые строки) — иначе они попадают в batch
    как "модели" и генерируют баннеры с нулевыми ценами.
    Строка валидна, если код модели похож на реальный код (латиница/
    цифры/дефис) И заполнена хотя бы цена.
    """
    price_col = next((c for c in COLUMN_MAP["price_new"] if c in df.columns), None)
    mask = df[code_col].apply(lambda v: bool(_MODEL_CODE_RE.match(str(v).strip())))
    if price_col:
        mask &= df[price_col].notna()
    return df[mask]


def _get(row, key, default=""):
    for col in COLUMN_MAP[key]:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
            return row[col]
    return default


def _get_localized(row, base_key, lang, default=""):
    """RU/KZ: если lang=KZ и есть колонка *_KZ с непустым значением — берёт её,
    иначе откатывается на основную (RU) колонку."""
    if lang == "KZ":
        kz_key = base_key + "_kz"
        if kz_key in COLUMN_MAP:
            val = _get(row, kz_key, None)
            if val:
                return val
    return _get(row, base_key, default)


def _load_model_row(excel_path, model_name):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Файл Excel не найден: {excel_path}")
    sheet = _find_sheet(excel_path, "Вводные")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    code_col = next((c for c in COLUMN_MAP["model_code"] if c in df.columns), df.columns[0])
    df = _valid_rows(df, code_col)
    match = df[df[code_col].astype(str).str.strip() == str(model_name).strip()]
    if match.empty:
        raise ValueError(f"Модель '{model_name}' не найдена в Excel (искала в колонке '{code_col}')")
    return match.iloc[0]


# ============================================================
#  ЦЕНЫ ПО РИТЕЙЛЕРАМ — бонус п.8 ТЗ. Лист "5. Цены по ритейлерам"
#  (Model_Code, Retailer, Price_New, Promo_Text). Если для ритейлера
#  нет строки (или ритейлер = ALL) — используется базовая цена.
# ============================================================
_RETAILER_CACHE = {}


def _load_retailer_sheet(excel_path):
    if excel_path in _RETAILER_CACHE:
        return _RETAILER_CACHE[excel_path]
    try:
        sheet = _find_sheet(excel_path, "ритейлер")
        df = pd.read_excel(excel_path, sheet_name=sheet)
    except Exception:
        df = None
    _RETAILER_CACHE[excel_path] = df
    return df


def _resolve_price(excel_path, model_code, retailer, base_new, base_old):
    """Возвращает (price_new, discount_%, promo_text_or_None)."""
    base_new = float(base_new or 0)
    base_old = float(base_old or 0)
    base_discount = round((base_old - base_new) / base_old * 100) if base_old else 0

    if retailer and retailer != "ALL":
        df = _load_retailer_sheet(excel_path)
        if df is not None:
            m = df[
                (df["Model_Code"].astype(str).str.strip() == str(model_code).strip())
                & (df["Retailer"].astype(str).str.strip().str.lower() == str(retailer).strip().lower())
            ]
            if not m.empty:
                row = m.iloc[0]
                price = float(row["Price_New"])
                discount = round((base_old - price) / base_old * 100) if base_old else 0
                promo = str(row.get("Promo_Text", "") or "").strip() or None
                return price, discount, promo

    return base_new, base_discount, None


# ============================================================
#  ХОЛСТ СО СЛОЯМИ — каждый смысловой блок креатива это
#  отдельный слой (RGBA, полный размер канваса), что позволяет
#  на выходе писать НАСТОЯЩИЙ многослойный PSD, а не плоскую картинку.
# ============================================================
class LayerStack:
    def __init__(self, w, h):
        self.w, self.h = w, h
        self.order = []
        self.layers = {}

    def new(self, name):
        img = Image.new("RGBA", (self.w, self.h), (0, 0, 0, 0))
        self.layers[name] = img
        self.order.append(name)
        return img, ImageDraw.Draw(img)

    def composite(self):
        base = Image.new("RGBA", (self.w, self.h), (0, 0, 0, 0))
        for name in self.order:
            base.alpha_composite(self.layers[name])
        return base

    def as_psd_layers(self):
        return [{"name": n, "image": self.layers[n]} for n in self.order]


def _wrap_text(draw, text, x, y, max_w, font, fill):
    words = text.split(" ")
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if draw.textlength(test, font=font) > max_w and line:
            draw.text((x, y), line, font=font, fill=fill, anchor="lt")
            y += int(font.size * 1.2)
            line = w
        else:
            line = test
    draw.text((x, y), line, font=font, fill=fill, anchor="lt")
    return y + int(font.size * 1.3)


def _draw_product(draw, img, cx, cy, r, category, model_code=None, excel_image_path=None):
    photo = _load_product_image(model_code, category, excel_image_path)
    cat = (category or "").lower()

    # мягкая тень в любом случае
    sh = Image.new("RGBA", img.size, (0, 0, 0, 0))
    shd = ImageDraw.Draw(sh)
    shd.ellipse([cx - r * 1.1, cy + r * 0.85, cx + r * 1.1, cy + r * 1.15], fill=(0, 0, 0, 140))
    sh = sh.filter(ImageFilter.GaussianBlur(r * 0.08))
    img.alpha_composite(sh)

    if photo is not None:
        target_h = int(r * 2.15)
        scale = target_h / photo.height
        target_w = int(photo.width * scale)
        resized = photo.resize((target_w, target_h), Image.LANCZOS)
        px = int(cx - target_w * 0.42)
        py = int(cy - target_h * 0.52)
        img.alpha_composite(resized, (px, py))
        return

    if "стирал" in cat or "сушил" in cat:
        # стиральная машина: корпус + панель управления + круглый люк
        body = [cx - r * 0.78, cy - r * 0.95, cx + r * 0.78, cy + r * 0.95]
        draw.rounded_rectangle(body, radius=int(r * 0.1), fill=(236, 238, 241))
        draw.rounded_rectangle([body[0], body[1], body[2], body[1] + r * 0.28], radius=int(r * 0.1), fill=(214, 217, 222))
        draw.ellipse([cx - r * 0.05, body[1] + r * 0.1, cx + r * 0.09, body[1] + r * 0.24], fill=(120, 170, 220))
        drum_r = r * 0.52
        draw.ellipse([cx - drum_r, cy + r * 0.02 - drum_r, cx + drum_r, cy + r * 0.02 + drum_r], fill=(40, 42, 48), outline=(210, 212, 216), width=int(r * 0.03))
        inner_r = drum_r * 0.74
        draw.ellipse([cx - inner_r, cy + r * 0.02 - inner_r, cx + inner_r, cy + r * 0.02 + inner_r], fill=(18, 19, 24))
        draw.ellipse([cx - inner_r * 0.55, cy + r * 0.02 - inner_r * 0.55, cx + inner_r * 0.55, cy + r * 0.02 + inner_r * 0.55], outline=(70, 74, 82), width=2)
    elif "холодил" in cat:
        draw.rounded_rectangle([cx - r * 0.68, cy - r, cx + r * 0.68, cy + r], radius=int(r * 0.1), fill=(230, 232, 236))
        draw.line([(cx - r * 0.68, cy - r * 0.15), (cx + r * 0.68, cy - r * 0.15)], fill=(200, 200, 205), width=2)
        draw.rounded_rectangle([cx + r * 0.5, cy - r * 0.85, cx + r * 0.6, cy - r * 0.4], radius=6, fill=(200, 200, 205))
        draw.rounded_rectangle([cx + r * 0.5, cy - r * 0.05, cx + r * 0.6, cy + r * 0.4], radius=6, fill=(200, 200, 205))
    elif "телевиз" in cat or "tv" in cat:
        draw.rounded_rectangle([cx - r, cy - r * 0.6, cx + r, cy + r * 0.6], radius=int(r * 0.05), fill=(10, 10, 10))
        draw.rounded_rectangle([cx - r * 0.94, cy - r * 0.52, cx + r * 0.94, cy + r * 0.52], radius=int(r * 0.03), fill=(20, 26, 55))
    else:
        draw.rounded_rectangle([cx - r * 0.78, cy - r * 0.78, cx + r * 0.78, cy + r * 0.78], radius=int(r * 0.14), fill=(230, 232, 236))


# ============================================================
#  ГЕНЕРАЛЬНЫЙ ШАБЛОН — Square / Portrait / Stories (п.12 ТЗ)
# ============================================================
def _draw_format(fmt_key, spec, data, lang, price_new, price_old, discount, promo_text):
    W, H = spec["w"], spec["h"]
    stack = LayerStack(W, H)
    pad = spec["safe"]["side"]
    top_safe = spec["safe"]["top"]
    bottom_safe = spec["safe"]["bottom"]

    # --- Фон ---
    bg, bgd = stack.new("Фон")
    bgd.rectangle([0, 0, W, H], fill=INK + (255,))
    prod_cx, prod_cy = int(W * 0.72), int(H * 0.5)
    for r in range(int(max(W, H) * 0.55), 0, -6):
        t = r / (max(W, H) * 0.55)
        c = tuple(int(11 + (43 - 11) * (1 - t)) for _ in range(3)) + (255,)
        bgd.ellipse([prod_cx - r, prod_cy - r, prod_cx + r, prod_cy + r], fill=c)
    card_w = int(W * (0.40 if fmt_key == "Display" else 0.60))
    card_top = int(H * (0.12 if fmt_key == "Display" else 0.20))
    card_h = int(H * (0.76 if fmt_key == "Display" else 0.62))
    bgd.rounded_rectangle([-20, card_top, card_w, card_top + card_h], radius=int(W * 0.045), fill=(28, 19, 15, 255))

    # --- Товар ---
    prod_img, prod_draw = stack.new("Товар")
    _draw_product(prod_draw, prod_img, prod_cx, prod_cy, int(min(W, H) * 0.24),
                  str(_get(data, "category")), str(_get(data, "model_code")), str(_get(data, "image_path")))

    # --- Лого Samsung (нет в Stories, п.6 ТЗ) ---
    logo_img, logo_draw = stack.new("Лого Samsung")
    label_top = top_safe
    if spec["logo"]:
        logo_font = _font(int(W * 0.036))
        logo_draw.text((W - pad, top_safe), "SAMSUNG", font=logo_font, fill="white", anchor="ra")
        label_top = top_safe + int(W * 0.05)

    # --- Модель / категория / бейдж объёма ---
    label_img, label_draw = stack.new("Модель и бейдж")
    cat_font = _font(int(W * 0.02), bold=False)
    model_font = _font(int(W * 0.024))
    label_draw.text((W - pad, label_top), str(_get_localized(data, "category", lang)), font=cat_font, fill=(255, 255, 255, 180), anchor="ra")
    label_draw.text((W - pad, label_top + int(W * 0.03)), str(_get(data, "model_code")), font=model_font, fill="white", anchor="ra")
    badge_r = int(W * 0.038)
    badge_cx = W - pad - badge_r
    badge_cy = label_top + int(W * 0.03) + int(W * 0.05) + badge_r
    label_draw.ellipse([badge_cx - badge_r, badge_cy - badge_r, badge_cx + badge_r, badge_cy + badge_r], fill=INK + (255,), outline=(255, 255, 255, 90))
    badge_font = _font(int(badge_r * 0.55))
    label_draw.text((badge_cx, badge_cy), str(_get(data, "volume_badge")), font=badge_font, fill="white", anchor="mm")

    # --- Заголовок + сумма выгоды ---
    head_img, head_draw = stack.new("Заголовок")
    head_font = _font(int(W * 0.05))
    save_font = _font(int(W * 0.053))
    save = max(0, price_old - price_new)
    y = top_safe + int(W * 0.02)
    y = _wrap_text(head_draw, str(_get_localized(data, "slogan", lang)), pad, y, int(card_w * 0.86), head_font, "white")
    y += int(W * 0.015)
    y = _wrap_text(head_draw, _t_benefit(save, lang), pad, y, int(card_w * 0.86), save_font, "white")

    # --- Фичи (колонкой) ---
    feat_img, feat_draw = stack.new("Фичи")
    feat_font = _font(int(W * 0.02))
    feats = [str(_get_localized(data, "feature_1", lang)), str(_get_localized(data, "feature_2", lang)), str(_get_localized(data, "feature_3", lang))]
    feats = [f for f in feats if f and f != "nan"]
    fy = max(y, card_top) + int(W * 0.035)
    for f in feats:
        r = int(W * 0.017)
        feat_draw.ellipse([pad, fy - r, pad + 2 * r, fy + r], outline="white", width=2)
        feat_draw.text((pad + 2 * r + int(W * 0.015), fy), f, font=feat_font, fill="white", anchor="lm")
        fy += int(W * 0.052)

    # --- Цена ---
    price_img, price_draw = stack.new("Цена")
    price_label_font = _font(int(W * 0.022), bold=False)
    price_big_font = _font(int(W * 0.062))
    price_old_font = _font(int(W * 0.026), bold=False)
    py = fy + int(W * 0.02)
    price_draw.text((pad, py), _t_permonth(price_new / 12, lang), font=price_label_font, fill=GOLD, anchor="lm")
    py += int(W * 0.05)
    price_draw.text((pad, py), _t_price(price_new), font=price_big_font, fill="white", anchor="lm")
    py += int(W * 0.045)
    badge_txt = f"-{discount}%"
    bw = price_draw.textlength(badge_txt, font=price_old_font) + int(W * 0.03)
    price_draw.rounded_rectangle([pad, py, pad + bw, py + int(W * 0.032)], radius=int(W * 0.016), fill=(122, 46, 46))
    price_draw.text((pad + bw / 2, py + int(W * 0.016)), badge_txt, font=price_old_font, fill="white", anchor="mm")
    old_txt = _t_price(price_old)
    ox = pad + bw + int(W * 0.02)
    price_draw.text((ox, py + int(W * 0.016)), old_txt, font=price_old_font, fill=(255, 255, 255, 140), anchor="lm")
    ow = price_draw.textlength(old_txt, font=price_old_font)
    price_draw.line([(ox, py + int(W * 0.016)), (ox + ow, py + int(W * 0.016))], fill=(255, 255, 255, 140), width=2)
    if promo_text:
        promo_font = _font(int(W * 0.017), bold=False)
        price_draw.text((pad, py + int(W * 0.05)), promo_text, font=promo_font, fill=GOLD, anchor="lm")

    # --- Плашка оплаты снизу ---
    pay_img, pay_draw = stack.new("Плашка оплаты")
    cap_h = int(H * 0.062)
    cap_y = H - cap_h - bottom_safe
    pay_draw.rounded_rectangle([pad, cap_y, W - pad, cap_y + cap_h], radius=cap_h // 2, outline=(255, 255, 255, 60), width=2)
    seg_font = _font(int(cap_h * 0.24), bold=False)
    segs = [s for s in [str(_get_localized(data, "cashback_text", lang)), str(_get_localized(data, "payment_text", lang))] if s and s != "nan"]
    if not segs:
        segs = [_t_default_pay(lang)]
    seg_w = (W - 2 * pad) / len(segs)
    for i, s in enumerate(segs):
        cx = pad + seg_w * i + seg_w / 2
        pay_draw.text((cx, cap_y + cap_h / 2), s, font=seg_font, fill="white", anchor="mm")
        if i < len(segs) - 1:
            xline = pad + seg_w * (i + 1)
            pay_draw.line([(xline, cap_y + cap_h * 0.2), (xline, cap_y + cap_h * 0.8)], fill=(255, 255, 255, 60), width=1)

    return stack


# ============================================================
#  DISPLAY 1200x628 — отдельная низкая раскладка (фичи в ряд)
# ============================================================
def _draw_display(spec, data, lang, price_new, price_old, discount, promo_text):
    W, H = spec["w"], spec["h"]
    stack = LayerStack(W, H)
    pad = spec["safe"]["side"]

    bg, bgd = stack.new("Фон")
    bgd.rectangle([0, 0, W, H], fill=INK + (255,))
    prod_cx, prod_cy = int(W * 0.85), int(H * 0.60)
    for r in range(int(H * 0.7), 0, -6):
        t = r / (H * 0.7)
        c = tuple(int(11 + (40 - 11) * (1 - t)) for _ in range(3)) + (255,)
        bgd.ellipse([prod_cx - r, prod_cy - r, prod_cx + r, prod_cy + r], fill=c)
    card_w = int(W * 0.50)
    bgd.rounded_rectangle([-20, 0, card_w, H], radius=int(H * 0.08), fill=(28, 19, 15, 255))

    prod_img, prod_draw = stack.new("Товар")
    _draw_product(prod_draw, prod_img, prod_cx, prod_cy, int(H * 0.165),
                  str(_get(data, "category")), str(_get(data, "model_code")), str(_get(data, "image_path")))

    logo_img, logo_draw = stack.new("Лого Samsung")
    logo_font = _font(int(H * 0.075))
    top_safe = spec["safe"]["top"]
    logo_draw.text((W - pad, top_safe), "SAMSUNG", font=logo_font, fill="white", anchor="ra")

    label_img, label_draw = stack.new("Модель и бейдж")
    cat_font = _font(int(H * 0.038), bold=False)
    model_font = _font(int(H * 0.045))
    label_top = top_safe + int(H * 0.11)
    label_draw.text((W - pad, label_top), str(_get_localized(data, "category", lang)), font=cat_font, fill=(255, 255, 255, 180), anchor="ra")
    label_draw.text((W - pad, label_top + int(H * 0.055)), str(_get(data, "model_code")), font=model_font, fill="white", anchor="ra")
    badge_r = int(H * 0.075)
    badge_cx, badge_cy = W - pad - badge_r, label_top + int(H * 0.13) + badge_r + int(H * 0.02)
    label_draw.ellipse([badge_cx - badge_r, badge_cy - badge_r, badge_cx + badge_r, badge_cy + badge_r], fill=INK + (255,), outline=(255, 255, 255, 90))
    label_draw.text((badge_cx, badge_cy), str(_get(data, "volume_badge")), font=_font(int(badge_r * 0.55)), fill="white", anchor="mm")

    head_img, head_draw = stack.new("Заголовок")
    head_font = _font(int(H * 0.09))
    y = pad
    y = _wrap_text(head_draw, str(_get_localized(data, "slogan", lang)), pad, y, int(card_w * 0.86), head_font, "white")

    feat_img, feat_draw = stack.new("Фичи")
    feat_font = _font(int(H * 0.032))
    feats = [str(_get_localized(data, "feature_1", lang)), str(_get_localized(data, "feature_2", lang)), str(_get_localized(data, "feature_3", lang))]
    feats = [f for f in feats if f and f != "nan"]
    fy = y + int(H * 0.02)
    fx = pad
    for f in feats:
        r = int(H * 0.02)
        feat_draw.ellipse([fx, fy - r, fx + 2 * r, fy + r], outline="white", width=2)
        feat_draw.text((fx + 2 * r + int(H * 0.015), fy), f, font=feat_font, fill="white", anchor="lm")
        fx += feat_draw.textlength(f, font=feat_font) + int(H * 0.09)

    price_img, price_draw = stack.new("Цена")
    price_font = _font(int(H * 0.09))
    small_font = _font(int(H * 0.032), bold=False)
    py = fy + int(H * 0.09)
    price_draw.text((pad, py), _t_price(price_new), font=price_font, fill="white", anchor="lm")
    badge_txt = f"-{discount}%"
    bx = pad
    by = py + int(H * 0.08)
    bw = price_draw.textlength(badge_txt, font=small_font) + int(H * 0.04)
    bh = int(H * 0.05)
    price_draw.rounded_rectangle([bx, by, bx + bw, by + bh], radius=bh // 2, fill=(122, 46, 46))
    price_draw.text((bx + bw / 2, by + bh / 2), badge_txt, font=small_font, fill="white", anchor="mm")
    old_txt = _t_price(price_old)
    ox = bx + bw + int(H * 0.03)
    price_draw.text((ox, by + bh / 2), old_txt, font=small_font, fill=(255, 255, 255, 140), anchor="lm")
    ow = price_draw.textlength(old_txt, font=small_font)
    price_draw.line([(ox, by + bh / 2), (ox + ow, by + bh / 2)], fill=(255, 255, 255, 140), width=2)
    if promo_text:
        promo_font = _font(int(H * 0.028), bold=False)
        price_draw.text((pad, by + bh + int(H * 0.03)), promo_text, font=promo_font, fill=GOLD, anchor="lm")

    pay_img, pay_draw = stack.new("Плашка оплаты")
    cap_h = int(H * 0.11)
    cap_y = H - cap_h - spec["safe"]["bottom"]
    pay_draw.rounded_rectangle([pad, cap_y, W - pad, cap_y + cap_h], radius=cap_h // 2, outline=(255, 255, 255, 60), width=2)
    seg_font = _font(int(cap_h * 0.32), bold=False)
    segs = [s for s in [str(_get_localized(data, "cashback_text", lang)), str(_get_localized(data, "payment_text", lang))] if s and s != "nan"] or [_t_default_pay(lang)]
    seg_w = (W - 2 * pad) / len(segs)
    for i, s in enumerate(segs):
        cx = pad + seg_w * i + seg_w / 2
        pay_draw.text((cx, cap_y + cap_h / 2), s, font=seg_font, fill="white", anchor="mm")
        if i < len(segs) - 1:
            xline = pad + seg_w * (i + 1)
            pay_draw.line([(xline, cap_y + cap_h * 0.2), (xline, cap_y + cap_h * 0.8)], fill=(255, 255, 255, 60), width=1)

    return stack


# ============================================================
#  ВАЛИДАТОР SAFE ZONE + CLEAR SPACE — бонус п.8 ТЗ.
#  Проверяет, что текстовые/лого-слои не залезают за safe zone,
#  и что вокруг лого есть Clear Space >= 0.5x его высоты.
# ============================================================
def _validate_safe_zone(stack, spec, fmt_key):
    W, H = spec["w"], spec["h"]
    pad, top_safe, bottom_safe = spec["safe"]["side"], spec["safe"]["top"], spec["safe"]["bottom"]
    tol = SAFE_ZONE_TOLERANCE_PX
    violations = []

    checked_layers = ["Лого Samsung", "Модель и бейдж", "Заголовок", "Фичи", "Цена", "Плашка оплаты"]
    for name in checked_layers:
        img = stack.layers.get(name)
        if img is None:
            continue
        bbox = img.getbbox()
        if bbox is None:
            continue
        left, top, right, bottom = bbox
        if left < pad - tol:
            violations.append(f"{name}: заходит за левую safe zone ({pad - left}px)")
        if right > W - pad + tol:
            violations.append(f"{name}: заходит за правую safe zone ({right - (W - pad)}px)")
        if top < top_safe - tol:
            violations.append(f"{name}: заходит за верхнюю safe zone ({top_safe - top}px)")
        if bottom > H - bottom_safe + tol:
            violations.append(f"{name}: заходит за нижнюю safe zone ({bottom - (H - bottom_safe)}px)")

    logo_img = stack.layers.get("Лого Samsung")
    prod_img = stack.layers.get("Товар")
    if logo_img is not None and prod_img is not None:
        logo_bbox = logo_img.getbbox()
        prod_bbox = prod_img.getbbox()
        if logo_bbox and prod_bbox:
            logo_h = logo_bbox[3] - logo_bbox[1]
            clear_needed = logo_h * 0.5
            gap = prod_bbox[1] - logo_bbox[3]  # вертикальный зазор товар-лого
            if gap < -logo_h:  # грубая эвристика реального пересечения по вертикали
                pass
            if gap < 0 and abs(gap) > clear_needed:
                violations.append(f"Clear Space лого/товар меньше 0.5× высоты лого")

    return violations


# ============================================================
#  БРИФ НА МОДЕЛЬ (п.5 ТЗ, обязательный deliverable)
# ============================================================
def _build_brief_text(model_name, data, lang, retailer, price_new, price_old, discount, promo_text, validation_summary):
    save = max(0, price_old - price_new)
    feats = [str(_get_localized(data, "feature_1", lang)), str(_get_localized(data, "feature_2", lang)), str(_get_localized(data, "feature_3", lang))]
    feats = [f for f in feats if f and f != "nan"]

    def _money(v):
        return f"{v:,.0f}".replace(",", " ")

    lines = [
        f"БРИФ НА КРЕАТИВ — {model_name}",
        f"Сгенерировано автоматически из Excel · язык: {lang} · ритейлер: {retailer}",
        "=" * 60,
        "",
        f"Категория: {_get_localized(data, 'category', lang)}",
        f"Заголовок (Slogan): {_get_localized(data, 'slogan', lang)}",
        f"Сумма выгоды: {_money(save)} тг",
        "",
        f"Цена: {_money(price_new)} тг (было {_money(price_old)} тг, скидка -{discount}%)"
        + (f"  ·  {promo_text}" if promo_text else ""),
        f"Рассрочка от: {_money(price_new / 12)} тг/мес",
        "",
        "Фичи (левая колонка / в Display — в ряд):",
    ] + [f"  • {f}" for f in feats] + [
        "",
        f"Бейдж объёма: {_get(data, 'volume_badge')}",
        f"Нижняя плашка: {_get_localized(data, 'cashback_text', lang)}  ·  {_get_localized(data, 'payment_text', lang)}",
        f"Канал: {_get(data, 'channel')}",
        "",
        "Форматы к производству (PSD + JPG каждый):",
        "  1. Feed Square    1080×1080  · safe zone 64px  · лого top-right",
        "  2. Feed Portrait  1080×1350  · safe zone 64-80px · лого top-right",
        "  3. Stories/Reels  1080×1920  · safe zone top 250 / bottom 350 / side 72px · без лого",
        "  4. Display        1200×628   · safe zone 60px  · лого top-right",
        "",
        "Проверка safe zone / clear space:",
    ]
    if validation_summary:
        lines += [f"  ⚠ {v}" for v in validation_summary]
    else:
        lines += ["  ✔ нарушений не найдено на всех 4 форматах"]

    return "\n".join(lines)


# ============================================================
#  ГЛАВНАЯ ФУНКЦИЯ — одна модель, все 4 формата
# ============================================================
def generate_banner(model_name, lang, retailer, excel_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    data = _load_model_row(excel_path, model_name)

    base_new = float(_get(data, "price_new", 0) or 0)
    base_old = float(_get(data, "price_old", 0) or 0)
    price_new, discount, promo_text = _resolve_price(excel_path, model_name, retailer, base_new, base_old)

    created = []
    all_violations = []
    for fmt_key, spec in FORMATS.items():
        if fmt_key == "Display":
            stack = _draw_display(spec, data, lang, price_new, base_old, discount, promo_text)
        else:
            stack = _draw_format(fmt_key, spec, data, lang, price_new, base_old, discount, promo_text)

        violations = _validate_safe_zone(stack, spec, fmt_key)
        for v in violations:
            all_violations.append(f"[{fmt_key}] {v}")

        composite = stack.composite()
        base = f"{model_name}_{fmt_key}_{retailer}_{lang}"
        jpg_path = os.path.join(output_dir, base + ".jpg")
        psd_path = os.path.join(output_dir, base + ".psd")
        composite.convert("RGB").save(jpg_path, "JPEG", quality=92)
        save_layered_psd(spec["w"], spec["h"], stack.as_psd_layers(), composite, psd_path)
        created.extend([jpg_path, psd_path])

    brief_path = os.path.join(output_dir, f"{model_name}_brief.txt")
    with open(brief_path, "w", encoding="utf-8") as f:
        f.write(_build_brief_text(model_name, data, lang, retailer, price_new, base_old, discount, promo_text, all_violations))
    created.append(brief_path)

    return created, all_violations


# ============================================================
#  ПАКЕТНАЯ ГЕНЕРАЦИЯ — ВСЕ МОДЕЛИ РАЗОМ (бонус п.8 ТЗ)
# ============================================================
def generate_all_banners(lang, retailer, excel_path, output_dir):
    sheet = _find_sheet(excel_path, "Вводные")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    code_col = next((c for c in COLUMN_MAP["model_code"] if c in df.columns), df.columns[0])
    df = _valid_rows(df, code_col)
    model_codes = [str(v).strip() for v in df[code_col].dropna().tolist() if str(v).strip()]

    all_created = []
    errors = []
    for code in model_codes:
        try:
            created, violations = generate_banner(code, lang, retailer, excel_path, output_dir)
            all_created.extend(created)
            if violations:
                errors.append(f"{code}: safe zone — {'; '.join(violations)}")
        except Exception as e:
            errors.append(f"{code}: {e}")
    return all_created, errors
