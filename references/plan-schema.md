# Plan JSON for `scripts/build_ppt.py`

The agent collects evidence; the offline builder renders it. `scripts/validate_plan.py` checks the important relationships. See `examples/example-plan.json` for a synthetic, fact-free smoke fixture.

```json
{
  "trip": {"people": 2, "place": "景点名称", "city": "城市", "date": "2027-05-12", "audience": "未指定"},
  "operation": {"status": "unverified", "hours": "待核实", "source_ids": [], "decision": "出发前核对官方售票与项目公告"},
  "weather": {"summary": "尚无可靠逐日预报", "source_ids": []},
  "map": {"image": "optional-local-image.jpg", "source_ids": ["map-source"], "edition_note": "版本日期未标"},
  "route": ["入口", "A 区", "B 区", "出口"],
  "attractions": [{"name": "项目 A", "area": "A 区", "operation": "unknown", "why": "体验特色", "source_ids": ["map-source"]}],
  "itinerary": [{"start": "09:30", "end": "10:00", "activity": "到场并核对当天公告", "buffer": "30 分钟", "fallback": "未开园则改期"}],
  "packing": ["证件或预约码", "充电宝"],
  "cautions": [{"text": "实际项目开放以现场为准", "source_ids": []}],
  "sources": [{"id": "map-source", "title": "园区导览图", "url": "https://example.org/map", "observed_at": "2027-05-01", "kind": "venue-map"}],
  "images": {"cover": "optional-local-image.jpg"}
}
```

`operation.status`: `verified_open`, `verified_closed`, `unverified`. Any verified status needs a source ID. `attractions[].operation`: `verified_open`, `verified_closed`, `unknown`; a verified attraction state needs its own source ID. A closed venue cannot have ride activities in `itinerary` (the validator catches this when an itinerary row includes an `attraction` ID). `source_ids` must resolve to entries in `sources`.

Optional sections may be empty, but the deck must stay honest. `map.image` and `images.cover` are local paths relative to the JSON file or absolute paths. A URL in an image path is not downloaded by the builder. Any numeric queue time in `buffer` is an **allowance for planning**, not a claimed live wait. Put actual live wait data in a separate source-backed field if available.
