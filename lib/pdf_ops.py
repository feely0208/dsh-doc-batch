#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 工具箱 · Python 侧实现（PyMuPDF / fitz）

由 bin/cli.mjs 通过 spawnSync 调用：从 stdin 读 JSON 任务，向 stdout 写一行 JSON 结果。
职责边界：
  - 只做 PDF 读写，不做依赖探测（探测在 Node 侧，便于给出「缺什么、怎么装」）
  - 只写 outDir 下的文件，绝不改动输入文件
  - 所有输出文件都由 unique_path() 生成，同名自动加序号，不覆盖
"""

import json
import os
import random
import sys
import traceback

try:
    import fitz  # PyMuPDF
except Exception as e:  # pragma: no cover - Node 侧已探测，这里兜底
    print(json.dumps({
        "ok": False,
        "error": "import-fit-z-failed",
        "detail": f"Python 无法 import fitz：{e}",
        "hint": "安装 PyMuPDF：python3 -m pip install PyMuPDF",
    }, ensure_ascii=False))
    sys.exit(3)


# ---------------------------------------------------------------- 通用工具

def unique_path(out_dir, stem, ext):
    """返回 outDir/stem.ext；若已存在则依次尝试 stem-1.ext、stem-2.ext……"""
    os.makedirs(out_dir, exist_ok=True)
    cand = os.path.join(out_dir, f"{stem}{ext}")
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(out_dir, f"{stem}-{i}{ext}")
        i += 1
    return cand


def save_doc(doc, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    doc.save(path, garbage=4, deflate=True, deflate_images=True, clean=True, use_objstms=1)
    return path


def pick_watermark_font(text):
    """中文用内置 CJK 字体，纯 ASCII 用内置 Helvetica。"""
    return "china-s" if any(ord(c) > 127 for c in text) else "helv"


def make_png(width, height, mode="image"):
    """用 fitz 现造一张 PNG 素材（演示/自检用，不依赖 Pillow）。"""
    if mode == "noise":
        # 高熵噪声：几乎不可压缩，用来演示「降采样 + JPEG 重编码」的压缩效果
        data = bytes(random.getrandbits(8) for _ in range(width * height * 3))
    else:
        # 简单渐变
        row = bytearray()
        for x in range(width):
            r = int(255 * x / max(width - 1, 1))
            g = int(160 * (1 - x / max(width - 1, 1)))
            b = 90
            row += bytes((r, g, b))
        data = bytes(row) * height
    pix = fitz.Pixmap(fitz.csRGB, width, height, data, 0)
    try:
        return pix.tobytes("png")
    finally:
        pix = None


# ---------------------------------------------------------------- 各动作

def act_merge(job):
    files = job["files"]
    out_dir = job["outDir"]
    base = job.get("base") or "merged"
    out = unique_path(out_dir, f"{base}-merged", ".pdf")
    dst = fitz.open()
    per_file = []
    for f in files:
        src = fitz.open(f)
        pages = src.page_count
        dst.insert_pdf(src)
        per_file.append({"input": f, "pages": pages})
        src.close()
    save_doc(dst, out)
    total = dst.page_count
    dst.close()
    return {
        "ok": True,
        "outputs": [out],
        "stats": {"files": len(files), "pages": total, "perFile": per_file},
    }


def act_split(job):
    src_path = job["file"]
    out_dir = job["outDir"]
    base = job.get("base") or os.path.splitext(os.path.basename(src_path))[0]
    mode = job.get("mode", "page")
    n = int(job.get("n") or 1)
    src = fitz.open(src_path)
    total = src.page_count
    outputs = []
    if mode == "every":
        if n < 1:
            raise ValueError("--n 必须是 ≥1 的整数")
        for start in range(0, total, n):
            end = min(start + n, total)
            name = f"{base}-p{start + 1:03d}-{end:03d}"
            out = unique_path(out_dir, name, ".pdf")
            dst = fitz.open()
            dst.insert_pdf(src, from_page=start, to_page=end - 1)
            save_doc(dst, out)
            dst.close()
            outputs.append(out)
    else:
        for pno in range(total):
            out = unique_path(out_dir, f"{base}-p{pno + 1:03d}", ".pdf")
            dst = fitz.open()
            dst.insert_pdf(src, from_page=pno, to_page=pno)
            save_doc(dst, out)
            dst.close()
            outputs.append(out)
    src.close()
    return {
        "ok": True,
        "outputs": outputs,
        "stats": {"pages": total, "mode": mode, "n": n, "parts": len(outputs)},
    }


def optimize_images(doc, dpi, quality):
    """
    图片优化：
      1) 优先用 Pillow 手动降采样 + JPEG 重编码（可精确控制 dpi / 质量，
         绕开 PyMuPDF 1.28 的 rewrite_images 在 dpi_target>72 时的参数 bug）；
      2) 没有 Pillow 时回退到 fitz 自带的 rewrite_images（目标 dpi 收敛到 72）。
    只替换「确实更小」的图片，绝不把图越压越大。
    返回 (模式, 替换张数)。
    """
    # 收集每张图片在所有页面上显示的最大宽度（pt）
    display = {}
    for page in doc:
        for info in page.get_images(full=True):
            xref = info[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            w = max((r.width for r in rects), default=0)
            if w > display.get(xref, 0):
                display[xref] = w

    try:
        from PIL import Image  # noqa: WPS433 (可选依赖)
        import io
    except Exception:
        # ---- 回退：PyMuPDF 自带实现（dpi_target 固定 72，避免其内部参数 bug）
        target = min(dpi, 72)
        try:
            doc.rewrite_images(dpi_threshold=max(target, 72), dpi_target=target,
                               quality=quality, lossy=True, lossless=True, bitonal=True)
            return "fitz-fallback", len(display)
        except Exception:
            return "none", 0

    replaced = 0
    done = set()
    for page in doc:
        for info in page.get_images(full=True):
            xref = info[0]
            if xref in done:
                continue
            done.add(xref)
            try:
                pix = fitz.Pixmap(doc, xref)
                orig = doc.extract_image(xref)
                orig_len = len(orig.get("image") or b"")
                rect_w = display.get(xref) or 0
                if pix.alpha:
                    img = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
                    white = Image.new("RGB", img.size, (255, 255, 255))
                    white.paste(img, mask=img.split()[-1])
                    img = white
                else:
                    if pix.n != 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                if rect_w > 0:
                    target_px = max(1, int(round(rect_w / 72.0 * dpi)))
                else:
                    target_px = img.width
                if target_px < img.width:
                    ratio = target_px / img.width
                    img = img.resize((target_px, max(1, int(round(img.height * ratio)))),
                                     Image.LANCZOS)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=max(10, min(95, quality)), optimize=True)
                data = buf.getvalue()
                if orig_len and len(data) >= orig_len:
                    continue  # 反而更大，保留原图
                page.replace_image(xref, stream=data)
                replaced += 1
            except Exception:
                continue
    return ("pil-manual" if replaced else "none"), replaced


def act_compress(job):
    src_path = job["file"]
    out_dir = job["outDir"]
    base = job.get("base") or os.path.splitext(os.path.basename(src_path))[0]
    dpi = int(job.get("dpi") or 96)
    quality = int(job.get("quality") or 65)
    before = os.path.getsize(src_path)
    doc = fitz.open(src_path)
    image_count = sum(len(page.get_images(full=True)) for page in doc)
    # 1) 图片降采样 / 有损重编码
    rewrite_mode, replaced = optimize_images(doc, dpi, quality)
    # 2) 字体子集化（失败不影响主流程）
    try:
        doc.subset_fonts()
        subset = True
    except Exception:
        subset = False
    out = unique_path(out_dir, f"{base}-compressed", ".pdf")
    save_doc(doc, out)
    doc.close()
    after = os.path.getsize(out)
    return {
        "ok": True,
        "outputs": [out],
        "stats": {
            "beforeBytes": before,
            "afterBytes": after,
            "savedBytes": before - after,
            "ratio": round(after / before, 4) if before else None,
            "smaller": after < before,
            "images": image_count,
            "imagesReplaced": replaced,
            "rewriteMode": rewrite_mode,
            "subsetFonts": subset,
        },
    }


def act_watermark(job):
    src_path = job["file"]
    out_dir = job["outDir"]
    text = job.get("text") or "DSH"
    base = job.get("base") or os.path.splitext(os.path.basename(src_path))[0]
    fontsize = float(job.get("fontsize") or 30)
    opacity = float(job.get("opacity") or 0.16)
    fontname = pick_watermark_font(text)
    doc = fitz.open(src_path)
    placed = 0
    for page in doc:
        rect = page.rect
        tw = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
        step_x = max(tw + 60, 120)
        step_y = max(fontsize * 3.2, 90)
        y = step_y
        while y < rect.height:
            x = 24
            while x < rect.width:
                page.insert_text((x, y), text, fontname=fontname, fontsize=fontsize,
                                 color=(0.72, 0.12, 0.12), fill_opacity=opacity)
                placed += 1
                x += step_x
            y += step_y
    out = unique_path(out_dir, f"{base}-watermarked", ".pdf")
    save_doc(doc, out)
    pages = doc.page_count
    doc.close()
    return {
        "ok": True,
        "outputs": [out],
        "stats": {"pages": pages, "text": text, "placements": placed, "opacity": opacity},
    }


def act_extract_text(job):
    src_path = job["file"]
    out_dir = job["outDir"]
    base = job.get("base") or os.path.splitext(os.path.basename(src_path))[0]
    doc = fitz.open(src_path)
    outputs = []
    total_chars = 0
    for pno, page in enumerate(doc):
        text = page.get_text("text")
        total_chars += len(text)
        out = unique_path(out_dir, f"{base}-p{pno + 1:03d}", ".txt")
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(text)
        outputs.append(out)
    pages = doc.page_count
    doc.close()
    return {
        "ok": True,
        "outputs": outputs,
        "stats": {"pages": pages, "chars": total_chars},
    }


def act_extract_images(job):
    src_path = job["file"]
    out_dir = job["outDir"]
    base = job.get("base") or os.path.splitext(os.path.basename(src_path))[0]
    min_bytes = int(job.get("minBytes") or 0)
    doc = fitz.open(src_path)
    outputs = []
    skipped = 0
    for pno, page in enumerate(doc):
        for idx, info in enumerate(page.get_images(full=True), start=1):
            xref = info[0]
            try:
                img = doc.extract_image(xref)
            except Exception:
                skipped += 1
                continue
            data = img.get("image") or b""
            if len(data) < min_bytes:
                skipped += 1
                continue
            ext = "." + (img.get("ext") or "png").lower()
            out = unique_path(out_dir, f"{base}-p{pno + 1:03d}-img{idx:02d}", ext)
            with open(out, "wb") as fh:
                fh.write(data)
            outputs.append(out)
    pages = doc.page_count
    doc.close()
    return {
        "ok": True,
        "outputs": outputs,
        "stats": {"pages": pages, "images": len(outputs), "skipped": skipped},
    }


# ---------------------------------------------------------------- 演示素材

def act_make_demo(job):
    """生成演示用 PDF：两页文本 A/B、两页重图、带内嵌图片的文本页。"""
    out = job["dir"]
    os.makedirs(out, exist_ok=True)

    def text_pdf(path, label, pages=2, image=False):
        doc = fitz.open()
        for i in range(pages):
            page = doc.new_page(width=595, height=842)
            page.insert_text((64, 90), "DSH PDF Kit demo", fontsize=12,
                             color=(0.35, 0.35, 0.35))
            page.insert_text((64, 150), f"{label} / page {i + 1} of {pages}", fontsize=26)
            page.insert_text((64, 200), "Hello from dsh-pdf-kit.", fontsize=16)
            if image:
                page.insert_image(fitz.Rect(64, 260, 500, 560), stream=make_png(560, 360))
        save_doc(doc, path)
        doc.close()
        return path

    a = text_pdf(os.path.join(out, "demo-a.pdf"), "Document A", 2, image=True)
    b = text_pdf(os.path.join(out, "demo-b.pdf"), "Document B", 2)

    heavy = os.path.join(out, "demo-heavy.pdf")
    noise = make_png(1000, 750, mode="noise")
    doc = fitz.open()
    for _ in range(2):
        page = doc.new_page(width=595, height=842)
        page.insert_image(fitz.Rect(50, 60, 545, 430), stream=noise)
        page.insert_text((50, 480), "Heavy image document (for compress demo)", fontsize=14)
    save_doc(doc, heavy)
    doc.close()

    return {"ok": True, "outputs": [a, b, heavy],
            "stats": {"dir": out, "textA": a, "textB": b, "heavy": heavy}}


# ---------------------------------------------------------------- 入口

ACTIONS = {
    "probe": lambda job: {
        "ok": True,
        "stats": {
            "python": sys.version.split()[0],
            "fitz": getattr(fitz, "VersionBind", "unknown"),
        },
    },
    "make-demo": act_make_demo,
    "merge": act_merge,
    "split": act_split,
    "compress": act_compress,
    "watermark": act_watermark,
    "extract-text": act_extract_text,
    "extract-images": act_extract_images,
}


def main():
    raw = sys.stdin.read()
    try:
        job = json.loads(raw) if raw.strip() else {}
    except Exception as e:
        print(json.dumps({"ok": False, "error": "bad-job-json", "detail": str(e)},
                         ensure_ascii=False))
        sys.exit(1)
    action = job.get("action")
    fn = ACTIONS.get(action)
    if not fn:
        print(json.dumps({"ok": False, "error": "unknown-action", "detail": action,
                          "hint": "支持：" + ", ".join(sorted(ACTIONS))}, ensure_ascii=False))
        sys.exit(1)
    try:
        result = fn(job)
        print(json.dumps(result, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({
            "ok": False,
            "error": "fitz-op-failed",
            "detail": f"{type(e).__name__}: {e}",
            "trace": traceback.format_exc(limit=3).strip().splitlines()[-1],
        }, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
