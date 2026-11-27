---
name: travel-plan-in-china
description: Create a source-grounded China travel or attraction day-plan PowerPoint from party size, destination, and date. Use for a requested trip PPT; not for booking or publishing.
---

# travel plan In China

Produce a usable, illustrated `.pptx` for a trip in mainland China. The only required user inputs are **number of travellers, destination, and travel date**. Resolve the destination's city and local date before searching. Do not infer that two people are a couple, or infer ages, mobility, budget, diet, or thrill tolerance from party size. Use a balanced default plan and make activity restrictions visible. Ask for an additional fact only when it is indispensable to a safe, executable plan; otherwise state the assumption in the deck.

## Finish line

Deliver a PowerPoint that a traveller can open and use: opening/booking decision, weather when forecastable, route on a real map when available, timed itinerary with meals and queue buffers, packing list, practical cautions, and clickable sources. A closed attraction produces a clear change-of-plan deck. An unverified opening produces a **conditional** plan, never a made-up opening schedule. Optional Word, media archives, or live-map links are extras only when requested or already authorized.

## Workflow

1. **Availability first.** Find the venue's official date-specific hours, reservation/ticket rules, closures, and attraction or show schedule. Check the forecast and warnings from an authoritative weather source. Record each claim's source URL, observed time, target date, and whether it is confirmed, historical, or unknown. Do not use a generic ticket page or old social post as proof of the requested day's operation. If the date is outside a reliable forecast window, say weather is pending and provide a later refresh point. Read [source rules](references/source-rules.md) for the evidence hierarchy and search limits.
2. **Gather on-site context.** Locate an official or clearly attributed venue map, entrances, areas, transit stop, food, toilets, and indoor rest areas. Use a map service for the trip **to** the venue; use the venue map for routes **inside** it. If the map edition is unknown, mark it as a reference and defer to the on-site version.
3. **Add traveller evidence.** When Xiaohongshu MCP is available, search the destination with `攻略` and `注意事项`, plus relevant local terms. Review up to ten relevant notes or videos per search purpose, deduplicate, and record the exact search, sort, collection time, likes and exclusions. For each caution note, inspect up to thirty high-like first-level comments **among the comments actually loaded** and also sample recent or dissenting comments when available. Never label this a platform-wide top ten unless the provider genuinely supports that claim. If XHS or comments are unavailable, use accessible public sources and state the smaller evidence scope. Do not stall the PPT waiting for a media downloader.
4. **Process useful videos.** If downloads are authorized and a capable tool is available, retain source links and verify downloaded files. Separate audio and use voice activity detection before Whisper so music-only clips do not create false narration. Check uncertain names and numbers against frames, captions, OCR, and official material. Treat old prices, events, and queue reports as historical experiences. See [source rules](references/source-rules.md).
5. **Plan the day.** Prioritize confirmed-open activities, nearby areas, fixed show times, rest and meals. Give realistic walking and queue **allowances**, not unsupported wait-time predictions. Add a same-area swap for a ride closure and a shorter plan for a late opening. Keep the route geographically coherent. If the venue is closed, do not schedule its attractions.
6. **Make and deliver the deck.** Build a structured plan JSON following [plan schema](references/plan-schema.md), validate it with `scripts/validate_plan.py`, then run `scripts/build_ppt.py`. Use the local PowerPoint/presentation skill or installed libraries for additional visual polish when useful. Render or open the `.pptx` and inspect the first, map, timetable and source slides; test archive integrity and hyperlinks. Deliver the first correct usable deck before optional refinements. Keep extraction, model, and programming notes out of traveller-facing slides.

## Reusable commands

```bash
python scripts/validate_plan.py plan.json
python scripts/build_ppt.py plan.json output.pptx
# Optional: choose an installed font on macOS/Linux or another workstation.
python scripts/build_ppt.py plan.json output.pptx --font "Source Han Sans SC"
```

`python-pptx` and Pillow are listed in `requirements.txt`. The scripts do not browse, download, authenticate, change proxies, or publish. The agent prepares the plan JSON from current sources; the builder only turns that evidence into a deck. A synthetic example for testing is in `examples/example-plan.json`.

## Readable PowerPoint output

Preserve the navy/teal/warm-orange/paper palette. `ORANGE` (`#E7774F`) is decorative; use `ORANGE_DARK` (`#BC5230`) for white-on-orange labels and orange text on white cards. Deep orange against `PAPER` is only about 4.39:1, so do not use it for small text directly on that background. Use `CLOSED` (`#B3261E`) for verified closures and retain the written status so color is never the only signal. Do not lighten `MUTED` without rechecking text contrast. Keep footers at least 9.5 pt.

The builder writes Latin and East Asian font declarations explicitly. Font selection is `--font` / `build(..., font=...)`, then `TRAVEL_PLAN_FONT`, then `Microsoft YaHei` (East Asian name `微软雅黑`). For another OS, choose an actually installed CJK font such as PingFang SC or Source Han Sans SC and re-render. These declarations do not embed fonts or guarantee identical layout on machines missing the font; PowerPoint has no CSS-style fallback chain here.

Timetable and source pages balance rows across pages and center shorter lists within the content area (10 events become 5+5; 8 sources become 4+4). Preserve itinerary and evidence entries rather than cutting them to fit a page. Long text still needs a visual fit check; this is row-count balancing, not automatic text-height measurement.

## Public repository boundary

Keep credentials, cookies, downloaded social videos, transcripts, third-party map files, and personal itineraries out of the GitHub repository. Store source links and attribution in each user's deliverable. Publishing the skill does not authorize publishing a user's trip or media.
