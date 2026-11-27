"""Render a validated trip plan JSON to a 16:9 PowerPoint deck."""

from __future__ import annotations

import argparse
import io
import json
import os
import sys
import zipfile
from pathlib import Path

from PIL import Image, ImageEnhance, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from validate_plan import validate

NAVY = "18313A"
PAPER = "F8F5ED"
WHITE = "FFFFFF"
TEAL = "087F76"
ORANGE = "E7774F"
# Keep the original orange for decoration; use this tone for text/white labels.
ORANGE_DARK = "BC5230"
CLOSED = "B3261E"
GOLD = "F3C669"
MUTED = "61727A"
PALE = "E4EEEA"
FONT = "Microsoft YaHei"


def balanced_pages(items, capacity):
    """Keep order and avoid a nearly empty last page without shrinking text."""
    if not items:
        return
    pages = (len(items) + capacity - 1) // capacity
    count, extra = divmod(len(items), pages)
    offset = 0
    for index in range(pages):
        size = count + (index < extra)
        yield offset, items[offset:offset + size]
        offset += size


def row_y(index, count, capacity, top, pitch):
    """Center short tables in their existing content area."""
    return top + (capacity - count) * pitch / 2 + index * pitch


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def box(slide, x, y, w, h, fill, rounded=False):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background()
    if rounded:
        try:
            shape.adjustments[0] = 0.14
        except (AttributeError, IndexError):
            pass
    return shape


def text(slide, value, x, y, w, h, size=16, color=NAVY, bold=False, link=None, align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for index, line in enumerate(str(value).split("\n")):
        p = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(0)
        p.line_spacing = 1.12
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        # PowerPoint does not reliably apply a:latin to Chinese glyphs.
        rpr = run._r.get_or_add_rPr()
        for tag, face in (("a:ea", "微软雅黑" if FONT == "Microsoft YaHei" else FONT), ("a:cs", FONT)):
            element = rpr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1])
            if element is None:
                element = OxmlElement(tag)
                rpr.append(element)
            element.set("typeface", face)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
        if link:
            run.hyperlink.address = link
    return shape


def local_image(value: str, base: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else base / path


def picture(slide, path: Path, x, y, w, h, *, crop=True, brightness=1.0):
    with Image.open(path) as source:
        im = source.convert("RGB")
        size = (max(1, int(w * 180)), max(1, int(h * 180)))
        if crop:
            im = ImageOps.fit(im, size, method=Image.Resampling.LANCZOS)
        else:
            im.thumbnail(size, Image.Resampling.LANCZOS)
            canvas = Image.new("RGB", size, "#FFFFFF")
            canvas.paste(im, ((size[0] - im.width) // 2, (size[1] - im.height) // 2))
            im = canvas
        if brightness != 1.0:
            im = ImageEnhance.Brightness(im).enhance(brightness)
        stream = io.BytesIO()
        im.save(stream, format="JPEG", quality=86, optimize=True)
    stream.seek(0)
    return slide.shapes.add_picture(stream, Inches(x), Inches(y), Inches(w), Inches(h))


def heading(prs, section, title, subtitle="", dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(NAVY if dark else PAPER)
    box(slide, .55, .42, .08, .32, ORANGE)
    text(slide, section, .75, .39, 9.4, .28, 10, GOLD if dark else TEAL, True)
    text(slide, title, .6, .85, 12.1, .65, 28, WHITE if dark else NAVY, True)
    if subtitle:
        text(slide, subtitle, .62, 1.52, 12, .42, 12, "D2E0DC" if dark else MUTED)
    return slide


def footer(prs, slide, label=""):
    box(slide, .62, 7.11, 12.05, .015, "D4DCD8")
    text(slide, label, .66, 7.17, 11.1, .24, 9.5, MUTED)
    text(slide, f"{len(prs.slides):02d}", 12.11, 7.17, .55, .24, 9.5, MUTED, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body, accent=TEAL, body_size=13):
    box(slide, x, y, w, h, WHITE, True)
    box(slide, x + .2, y + .24, .08, .50, accent)
    text(slide, title, x + .43, y + .22, w - .65, .43, 18, NAVY, True)
    text(slide, body, x + .25, y + .89, w - .50, h - 1.02, body_size, MUTED)


def source_label(ids: list[str], by_id: dict[str, dict]) -> str:
    labels = [by_id[x].get("title", x) for x in ids if x in by_id]
    return "来源：" + "、".join(labels[:2]) if labels else ""


def build(plan: dict, base: Path, output: Path, *, font: str | None = None) -> int:
    global FONT
    FONT = (font or os.environ.get("TRAVEL_PLAN_FONT") or "Microsoft YaHei").strip()
    if not FONT:
        FONT = "Microsoft YaHei"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    trip, operation = plan["trip"], plan["operation"]
    weather = plan.get("weather") or {}
    sources = plan.get("sources") or []
    by_id = {x["id"]: x for x in sources}
    is_demo = bool(trip.get("demo"))
    status = operation["status"]
    status_color = {"verified_open": TEAL, "verified_closed": CLOSED, "unverified": ORANGE_DARK}[status]
    status_label = {
        "verified_open": "已核实营业 · 项目仍需逐项确认",
        "verified_closed": "已核实暂停营业 · 请改期或更换地点",
        "unverified": "条件版 · 出发前核实营业",
    }[status]

    # Cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cover = (plan.get("images") or {}).get("cover")
    if cover:
        picture(s, local_image(cover, base), 0, 0, 13.333, 7.5, brightness=.38)
    else:
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = rgb(NAVY)
        box(s, 8.8, -.8, 5.5, 9.0, TEAL, True)
        box(s, 10.5, -.5, 3.4, 8.4, ORANGE, True)
    box(s, .70, .65, .09, .56, ORANGE)
    text(s, f"{trip['date']} · {trip.get('city') or '中国'}", .96, .72, 9.7, .35, 17, GOLD, True)
    text(s, f"{trip['place']}\n{trip['people']} 人游玩计划", .94, 1.55, 10.8, 1.70, 39, WHITE, True)
    text(s, "地图路线 · 时间安排 · 行前清单 · 信息来源", .98, 4.04, 10.0, .45, 18, WHITE)
    box(s, .96, 5.12, 6.55, .69, status_color, True)
    text(s, status_label, 1.16, 5.28, 6.13, .34, 16, WHITE, True)
    if is_demo:
        text(s, "演示数据，不可作为真实出行依据", .97, 6.76, 9.8, .36, 15, WHITE, True)
    else:
        text(s, "目的地日期与当日公告为准", .97, 6.78, 9.8, .30, 11, WHITE)

    # Availability and weather
    s = heading(prs, "01 · 先决定是否出发", "当天能玩吗？先看官方信息", "开园、预约和项目运行会变化；这页决定后续行程是否可执行。")
    card(s, .65, 2.07, 5.84, 3.77, "营业与预约", f"{status_label}\n{operation.get('hours') or '营业时间待核实'}\n{operation.get('decision') or ''}", status_color, 15)
    card(s, 6.82, 2.07, 5.84, 3.77, "天气与穿着", weather.get("summary") or "目标日期天气尚未核实。\n出发前查看最新预报和预警。", TEAL, 15)
    box(s, .68, 6.11, 11.97, .55, PALE, True)
    text(s, "项目是否运行、演出是否举行，以目标日期园方公告和现场信息为准。", .90, 6.27, 11.4, .27, 13, NAVY, True)
    footer(prs, s, source_label(operation.get("source_ids", []) + weather.get("source_ids", []), by_id))

    # Candidate attractions
    attractions = plan.get("attractions") or []
    if not attractions:
        s = heading(prs, "02 · 游玩重点", "项目清单待核实", "先确认营业及项目表，再确定优先顺序。")
        card(s, .68, 2.12, 11.96, 3.7, "先看当天公告", "当前没有足够可靠的资料列出具体项目。可先使用后面的路线与时间框架，出发前补上项目。", ORANGE, 17)
        footer(prs, s)
    else:
        for page in range(0, len(attractions), 6):
            subset = attractions[page:page + 6]
            s = heading(prs, "02 · 游玩重点", "值得考虑的项目", "优先级来自体验和位置；开放状态写在每张卡片上。")
            for index, item in enumerate(subset):
                x = .65 + (index % 3) * 4.09
                y = 2.06 + (index // 3) * 2.19
                state = {
                    "verified_open": "当天已核实开放",
                    "verified_closed": "当天已核实关闭",
                    "unknown": "当天开放待核实",
                }[item.get("operation", "unknown")]
                body = f"{item.get('area') or '园区内'} · {state}\n{item.get('why') or '按兴趣和现场排队决定'}"
                accent = CLOSED if item.get("operation") == "verified_closed" else TEAL
                card(s, x, y, 3.86, 1.94, item["name"], body, accent, 12.2)
            footer(prs, s, "资料出处见末页；任何未核实项目都不是开放承诺。")

    # Map and route
    route = plan.get("route") or []
    venue_map = plan.get("map") or {}
    s = heading(prs, "03 · 路线", "按区域走，少折返", venue_map.get("edition_note") or "园内路线以现场导览图为准。")
    map_image = venue_map.get("image")
    if map_image:
        picture(s, local_image(map_image, base), .65, 2.05, 7.45, 4.88, crop=False)
    else:
        box(s, .65, 2.05, 7.45, 4.88, PALE, True)
        text(s, "地图尚未取得\n使用右侧区域路线", 1.09, 3.70, 6.55, 1.0, 24, TEAL, True, align=PP_ALIGN.CENTER)
    box(s, 8.39, 2.05, 4.25, 4.88, WHITE, True)
    text(s, "建议顺序", 8.68, 2.35, 3.65, .43, 19, NAVY, True)
    for index, stop in enumerate(route[:7]):
        y = 2.94 + index * .51
        text(s, f"{index + 1:02d}", 8.69, y, .48, .31, 13, ORANGE_DARK, True)
        text(s, stop, 9.28, y, 3.03, .36, 14, NAVY)
    footer(prs, s, source_label(venue_map.get("source_ids", []), by_id))

    # Time blocks
    itinerary = plan.get("itinerary") or []
    if not itinerary:
        s = heading(prs, "04 · 时间表", "确认营业后再安排具体时段")
        card(s, .68, 2.12, 11.96, 3.7, "先核实，再排时间", "当前缺少可用的开闭园或项目时间。不要把通用营业时间当作目标日期的承诺。", ORANGE, 17)
        footer(prs, s)
    else:
        for page, subset in balanced_pages(itinerary, 6):
            s = heading(prs, "04 · 全天时间表", "今天怎么走", "时间含步行、排队和休息容错；根据现场情况替换同一区项目。")
            for index, event in enumerate(subset):
                y = row_y(index, len(subset), 6, 2.03, .75)
                box(s, .66, y, 12.0, .65, WHITE, True)
                box(s, .66, y, 2.09, .65, ORANGE_DARK if index % 2 == 0 else TEAL, True)
                text(s, f"{event['start']}–{event['end']}", .80, y + .17, 1.80, .27, 12.5, WHITE, True, align=PP_ALIGN.CENTER)
                text(s, event["activity"], 2.95, y + .15, 5.32, .32, 15, NAVY, True)
                text(s, event.get("buffer") or "", 8.38, y + .17, 1.62, .29, 11.5, TEAL, True)
                text(s, event.get("fallback") or "", 10.09, y + .16, 2.37, .35, 10.2, MUTED)
            footer(prs, s, "排队时间为计划留量，实际排队以现场信息为准。")

    # Packing
    packing = plan.get("packing") or []
    s = heading(prs, "05 · 出发清单", "带这些，路上更从容", "按天气、场馆规则和同行人数调整。")
    if not packing:
        packing = ["证件或预约码", "手机与充电宝", "饮水", "适合步行的鞋"]
    for index, item in enumerate(packing[:12]):
        col, row = index % 3, index // 3
        x, y = .66 + col * 4.09, 2.03 + row * 1.06
        box(s, x, y, 3.87, .90, WHITE, True)
        box(s, x + .20, y + .22, .38, .39, TEAL, True)
        text(s, "✓", x + .27, y + .29, .22, .22, 12, WHITE, True, align=PP_ALIGN.CENTER)
        text(s, item, x + .73, y + .24, 2.86, .38, 13.7, NAVY, True)
    footer(prs, s, "入园携带限制以场馆当天规则为准。")

    # Cautions
    cautions = plan.get("cautions") or []
    if cautions:
        for page in range(0, len(cautions), 5):
            s = heading(prs, "06 · 注意事项", "出发前值得看一眼", "旧经验用来提醒要核实什么，不代替当天规则。")
            for index, caution in enumerate(cautions[page:page + 5]):
                item = caution if isinstance(caution, dict) else {"text": str(caution)}
                y = 2.06 + index * .89
                box(s, .67, y, 11.99, .75, WHITE, True)
                box(s, .89, y + .17, .45, .40, ORANGE_DARK if index % 2 == 0 else TEAL, True)
                text(s, str(page + index + 1), .99, y + .24, .24, .25, 11.5, WHITE, True, align=PP_ALIGN.CENTER)
                text(s, item.get("text", ""), 1.58, y + .17, 10.76, .41, 13.6, NAVY)
            footer(prs, s, "资料出处见末页；实际规定以官方公告与现场提示为准。")

    # Sources
    if sources:
        for page, subset in balanced_pages(sources, 7):
            s = heading(prs, "资料与链接", "从这里打开原始信息", "来源在目标日期前可能变化；出发前请重新查看营业与天气。")
            for index, source in enumerate(subset):
                y = row_y(index, len(subset), 7, 2.04, .64)
                box(s, .67, y, 11.99, .55, WHITE, True)
                text(s, f"{page + index + 1:02d}", .88, y + .13, .45, .26, 11.5, ORANGE_DARK, True)
                label = source.get("title", source["id"])[:43]
                text(s, label, 1.55, y + .11, 8.80, .31, 13, NAVY, True, link=source["url"])
                text(s, f"{source.get('observed_at', '')[:10]}  ↗", 10.55, y + .13, 1.75, .27, 10, TEAL, True, link=source["url"])
            footer(prs, s, "点击标题打开来源。")

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    with zipfile.ZipFile(output) as archive:
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"Corrupt PPTX member: {bad}")
    return len(prs.slides)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--font", help="Installed font family; overrides TRAVEL_PLAN_FONT (default: Microsoft YaHei)")
    args = parser.parse_args()
    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    errors = validate(plan, args.plan.resolve().parent)
    if errors:
        print(json.dumps({"valid": False, "errors": errors}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    slides = build(plan, args.plan.resolve().parent, args.output, font=args.font)
    print(json.dumps({"pptx": str(args.output.resolve()), "slides": slides, "font": FONT}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
