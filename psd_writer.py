"""
Собственный писатель многослойного PSD — без сторонних библиотек
(psd-tools/pytoshop недоступны офлайн). Реализован по бинарной
спецификации Adobe Photoshop File Format:
https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/

Каждый слой сжимается PackBits (RLE) построчно, поканально (R,G,B,A).
Финальное сведённое изображение (Image Data) пишется без сжатия для
надёжности.
"""
import struct


def _packbits_encode_row(data: bytes) -> bytes:
    out = bytearray()
    n = len(data)
    i = 0
    while i < n:
        run_len = 1
        while i + run_len < n and data[i + run_len] == data[i] and run_len < 128:
            run_len += 1
        if run_len >= 2:
            out.append((257 - run_len) & 0xFF)
            out.append(data[i])
            i += run_len
        else:
            start = i
            i += 1
            while i < n:
                if i + 1 < n and data[i] == data[i + 1]:
                    break
                if i - start >= 128:
                    break
                i += 1
            lit = data[start:i]
            out.append(len(lit) - 1)
            out.extend(lit)
    return bytes(out)


def _encode_channel_rle(channel_bytes: bytes, width: int, height: int) -> bytes:
    out = bytearray()
    out += struct.pack(">H", 1)  # compression = 1 (RLE / PackBits)
    rows = []
    for r in range(height):
        row = channel_bytes[r * width:(r + 1) * width]
        rows.append(_packbits_encode_row(row))
    for enc in rows:
        out += struct.pack(">H", len(enc))
    for enc in rows:
        out += enc
    return bytes(out)


def _pascal_name(name: str) -> bytes:
    nb = name.encode("cp1251", errors="replace")[:255]
    s = bytes([len(nb)]) + nb
    pad = (-len(s)) % 4
    return s + b"\x00" * pad


def save_layered_psd(width, height, layers, composite_rgb, path):
    """
    layers: список слоёв снизу вверх, каждый:
        {"name": str, "image": PIL.Image в режиме RGBA, размер (width,height)}
    composite_rgb: PIL.Image RGB — сведённая финальная картинка
                   (то, что видно при открытии без чтения слоёв / thumbnail)
    """
    header = struct.pack(">4sH6sHIIHH", b"8BPS", 1, b"\x00" * 6, 3, height, width, 8, 3)
    color_mode_data = struct.pack(">I", 0)
    image_resources = struct.pack(">I", 0)

    layer_records = bytearray()
    channel_data_blocks = bytearray()
    kept = 0

    for layer in layers:
        img = layer["image"]
        bbox = img.getbbox()
        if bbox is None:
            continue  # пустой слой (например, лого в Stories) — не пишем
        left, top, right, bottom = bbox
        cropped = img.crop(bbox)
        w, h = right - left, bottom - top
        r, g, b, a = cropped.split()
        channels = [(0, r.tobytes()), (1, g.tobytes()), (2, b.tobytes()), (-1, a.tobytes())]

        rect = struct.pack(">4i", top, left, bottom, right)
        num_ch = struct.pack(">H", len(channels))

        channel_info = b""
        payloads = []
        for cid, raw in channels:
            enc = _encode_channel_rle(raw, w, h)
            channel_info += struct.pack(">hI", cid, len(enc))
            payloads.append(enc)

        name_field = _pascal_name(layer["name"])
        extra_data = struct.pack(">I", 0) + struct.pack(">I", 0) + name_field
        rec = (
            rect + num_ch + channel_info
            + b"8BIM" + b"norm"
            + struct.pack(">BBBB", 255, 0, 0, 0)
            + struct.pack(">I", len(extra_data)) + extra_data
        )
        layer_records += rec
        for p in payloads:
            channel_data_blocks += p
        kept += 1

    layer_info_content = struct.pack(">h", kept) + bytes(layer_records) + bytes(channel_data_blocks)
    if len(layer_info_content) % 2 != 0:
        layer_info_content += b"\x00"
    layer_info_section = struct.pack(">I", len(layer_info_content)) + layer_info_content
    global_layer_mask_info = struct.pack(">I", 0)
    layer_mask_info_content = layer_info_section + global_layer_mask_info
    layer_mask_info_section = struct.pack(">I", len(layer_mask_info_content)) + layer_mask_info_content

    comp = composite_rgb.convert("RGB")
    cr, cg, cb = comp.split()
    image_data = struct.pack(">H", 0) + cr.tobytes() + cg.tobytes() + cb.tobytes()

    with open(path, "wb") as f:
        f.write(header)
        f.write(color_mode_data)
        f.write(image_resources)
        f.write(layer_mask_info_section)
        f.write(image_data)

    return kept
