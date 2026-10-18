import os
import time
import json
import datetime as dt
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError
from PIL import Image, ImageDraw, ImageFont

PRINTER_NAME = "EPSON TM-T88VII Receipt"
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "print_bridge_config.json")
POLL_SECONDS = 3

def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise RuntimeError("print_bridge_config.json not found")
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    portal = str(cfg.get("portal_url", "")).strip().rstrip("/")
    key = str(cfg.get("bridge_key", "")).strip()
    if not portal.startswith("https://") or not key:
        raise RuntimeError("Set portal_url (https://...) and bridge_key in print_bridge_config.json")
    return portal, key

def api_json(method, url, key, payload=None):
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urlrequest.Request(url, data=data, method=method)
    req.add_header("X-Print-Bridge-Key", key)
    req.add_header("Content-Type", "application/json")
    with urlrequest.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))

def font(size, bold=False):
    choices = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
    ]
    for path in choices:
        if os.path.exists(path): return ImageFont.truetype(path, size)
    return ImageFont.load_default()

def fit_font(draw, text, max_width, start=76, minimum=40):
    for size in range(start, minimum - 1, -1):
        f = font(size, True)
        if draw.textbbox((0, 0), text, font=f)[2] <= max_width: return f
    return font(minimum, True)

def format_price(value):
    raw = str(value or "").strip().replace("KZ", "").replace("AOA", "").strip()
    try:
        normalized = raw.replace(" ", "")
        if "," in normalized and "." in normalized: normalized = normalized.replace(",", "")
        elif normalized.count(",") == 1 and len(normalized.split(",")[-1]) == 2: normalized = normalized.replace(",", ".")
        number = float(normalized)
        return f"{int(number):,}" if number.is_integer() else f"{number:,.2f}"
    except Exception:
        return raw

def make_ticket(job):
    w, h = 576, 320
    img = Image.new("1", (w, h), 1)
    d = ImageDraw.Draw(img)
    d.rectangle((1, 1, w - 2, h - 2), outline=0, width=2)
    name = str(job.get("item_name", "ARTICLE")).strip().upper()
    words, lines, current = name.split(), [], ""
    for word in words:
        candidate = (current + " " + word).strip()
        if len(candidate) <= 18 and len(lines) < 2: current = candidate
        else:
            if current: lines.append(current)
            current = word
    if current and len(lines) < 2: lines.append(current)
    for i, line in enumerate((lines or ["ARTICLE"])[:2]): d.text((14, 35 + i * 38), line, font=font(29, True), fill=0)
    d.text((365, 28), "FRESMART", font=font(27, True), fill=0)
    d.text((500, 68), str(job.get("currency", "KZ")).upper(), font=font(26, True), fill=0)
    price = format_price(job.get("price", ""))
    pf = fit_font(d, price, 300)
    pw = d.textbbox((0, 0), price, font=pf)[2]
    d.text((w - 18 - pw, 102), price, font=pf, fill=0)
    d.text((310, 197), f"PREÇO POR {str(job.get('unit', 'UN')).upper()}", font=font(24, True), fill=0)
    d.text((367, 237), "IVA INCLUIDO", font=font(22, True), fill=0)
    d.text((18, 270), str(job.get("article_code", "")), font=font(25, True), fill=0)
    d.text((213, 274), dt.datetime.now().strftime("%d/%m/%y"), font=font(21, True), fill=0)
    return img

def raster_bytes(img):
    img = img.convert("1")
    w, h = img.size
    width_bytes = (w + 7) // 8
    out = bytearray()
    pix = img.load()
    for y in range(h):
        for xb in range(width_bytes):
            b = 0
            for bit in range(8):
                x = xb * 8 + bit
                if x < w and pix[x, y] == 0: b |= 1 << (7 - bit)
            out.append(b)
    return b"\x1b@\x1dv0\x00" + bytes([width_bytes & 255, (width_bytes >> 8) & 255, h & 255, (h >> 8) & 255]) + bytes(out) + b"\n\n\n\x1dV\x00"

def raw_print(data, copies=1):
    import win32print
    handle = win32print.OpenPrinter(PRINTER_NAME)
    try:
        for _ in range(max(1, min(int(copies or 1), 10))):
            win32print.StartDocPrinter(handle, 1, ("Fresmart Price Ticket", None, "RAW"))
            try:
                win32print.StartPagePrinter(handle)
                win32print.WritePrinter(handle, data)
                win32print.EndPagePrinter(handle)
            finally:
                win32print.EndDocPrinter(handle)
    finally:
        win32print.ClosePrinter(handle)

def main():
    portal, key = load_config()
    print("Fresmart Render Print Bridge")
    print("Portal:", portal)
    print("Printer:", PRINTER_NAME)
    print("Waiting for print jobs. Do not close this window.")
    while True:
        try:
            result = api_json("GET", portal + "/api/print-jobs/next", key)
            job = result.get("job")
            if job:
                print("Printing job", job["id"], job["article_code"])
                try:
                    raw_print(raster_bytes(make_ticket(job)), job.get("copies", 1))
                    api_json("POST", portal + f"/api/print-jobs/{job['id']}/complete", key, {"success": True})
                    print("Printed job", job["id"])
                except Exception as exc:
                    api_json("POST", portal + f"/api/print-jobs/{job['id']}/complete", key, {"success": False, "error": str(exc)})
                    print("Print error:", exc)
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            print("Portal connection retry:", exc)
        except Exception as exc:
            print("Bridge error:", exc)
        time.sleep(POLL_SECONDS)

if __name__ == "__main__":
    main()
