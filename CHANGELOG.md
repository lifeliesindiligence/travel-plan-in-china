# Changelog

## 0.2.0 — 2026-11-27

- Keep decorative orange `#E7774F`; use `#BC5230` for white labels and orange text on white cards (not small text on the paper background).
- Give verified closures a separate red `#B3261E`, retaining written status labels.
- Write explicit East Asian and complex-script font declarations. Add `--font`, `TRAVEL_PLAN_FONT`, and a keyword font option for programmatic builds; default remains Microsoft YaHei / 微软雅黑.
- Increase footer labels and page numbers to 9.5 pt.
- Balance timetable/source pages and vertically center short lists, without dropping entries or shrinking text (10 events: 5+5; 8 sources: 4+4).
- Add regression coverage for contrast, font declarations/selection, pagination, closure color, and PPTX hyperlinks.

## 0.1.0 — 2026-11-20

- First reusable version of `travel plan In China`.
- Three required inputs: traveler count, destination, and date.
- Evidence rules for opening, weather, attractions, maps, Xiaohongshu notes and comments.
- JSON validation and offline PowerPoint generation with an explicitly synthetic example.

