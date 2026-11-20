"""Validate the evidence and schedule contract before rendering a trip deck."""

from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse


def _date(value: object) -> bool:
    try:
        date.fromisoformat(str(value))
        return True
    except ValueError:
        return False


def _time(value: object) -> int | None:
    try:
        parsed = datetime.strptime(str(value), "%H:%M")
        return parsed.hour * 60 + parsed.minute
    except ValueError:
        return None


def _image(path_value: object, base: Path) -> bool:
    if not isinstance(path_value, str) or not path_value:
        return False
    if urlparse(path_value).scheme in ("http", "https"):
        return False
    path = Path(path_value)
    return (path if path.is_absolute() else base / path).is_file()


def validate(plan: dict, base: Path) -> list[str]:
    errors: list[str] = []
    if not isinstance(plan, dict):
        return ["Plan root must be an object"]
    trip = plan.get("trip") or {}
    if not isinstance(trip, dict):
        return ["trip must be an object"]
    if type(trip.get("people")) is not int or not 1 <= trip["people"] <= 99:
        errors.append("trip.people must be an integer from 1 to 99")
    if not isinstance(trip.get("place"), str) or not trip["place"].strip():
        errors.append("trip.place is required")
    if not _date(trip.get("date")):
        errors.append("trip.date must be YYYY-MM-DD")

    sources = plan.get("sources") or []
    if not isinstance(sources, list):
        return errors + ["sources must be an array"]
    by_id: dict[str, dict] = {}
    for index, source in enumerate(sources):
        if not isinstance(source, dict):
            errors.append(f"sources[{index}] must be an object")
            continue
        sid = source.get("id")
        if not isinstance(sid, str) or not sid.strip() or sid in by_id:
            errors.append(f"sources[{index}].id must be unique and nonempty")
        else:
            by_id[sid] = source
        url = source.get("url")
        if not isinstance(url, str) or urlparse(url).scheme not in ("http", "https") or not urlparse(url).netloc:
            errors.append(f"sources[{index}].url must be an HTTP(S) URL")
        if not _date(str(source.get("observed_at", ""))[:10]):
            errors.append(f"sources[{index}].observed_at must begin with YYYY-MM-DD")
        if source.get("target_date") and not _date(source["target_date"]):
            errors.append(f"sources[{index}].target_date must be YYYY-MM-DD")

    def refs(obj: dict, label: str) -> list[dict]:
        ids = obj.get("source_ids") or []
        if not isinstance(ids, list) or not all(isinstance(x, str) for x in ids):
            errors.append(f"{label}.source_ids must be an array of IDs")
            return []
        missing = [x for x in ids if x not in by_id]
        if missing:
            errors.append(f"{label}.source_ids unknown: {', '.join(missing)}")
        return [by_id[x] for x in ids if x in by_id]

    operation = plan.get("operation") or {}
    if not isinstance(operation, dict):
        errors.append("operation must be an object")
        operation = {}
    status = operation.get("status")
    if status not in ("verified_open", "verified_closed", "unverified"):
        errors.append("operation.status must be verified_open, verified_closed, or unverified")
    op_sources = refs(operation, "operation")
    if status in ("verified_open", "verified_closed") and not any(
        x.get("kind") == "official-venue" and x.get("target_date") == trip.get("date")
        for x in op_sources
    ):
        errors.append("Verified operation requires an official-venue source for the trip date")

    weather = plan.get("weather") or {}
    if not isinstance(weather, dict):
        errors.append("weather must be an object")
    else:
        refs(weather, "weather")
        if weather.get("status") == "forecast" and not weather.get("source_ids"):
            errors.append("A weather forecast requires a source")

    venue_map = plan.get("map") or {}
    if not isinstance(venue_map, dict):
        errors.append("map must be an object")
    else:
        refs(venue_map, "map")
        if venue_map.get("image") and not _image(venue_map["image"], base):
            errors.append("map.image must point to an existing local image")

    images = plan.get("images") or {}
    if isinstance(images, dict) and images.get("cover") and not _image(images["cover"], base):
        errors.append("images.cover must point to an existing local image")

    attractions = plan.get("attractions") or []
    if not isinstance(attractions, list):
        errors.append("attractions must be an array")
        attractions = []
    attraction_names: dict[str, dict] = {}
    for index, attraction in enumerate(attractions):
        if not isinstance(attraction, dict) or not isinstance(attraction.get("name"), str):
            errors.append(f"attractions[{index}] needs a name")
            continue
        name = attraction["name"].strip()
        if not name or name in attraction_names:
            errors.append(f"attractions[{index}].name must be unique and nonempty")
        attraction_names[name] = attraction
        state = attraction.get("operation", "unknown")
        if state not in ("verified_open", "verified_closed", "unknown"):
            errors.append(f"attractions[{index}].operation is invalid")
        at_sources = refs(attraction, f"attractions[{index}]")
        if state in ("verified_open", "verified_closed") and not any(
            x.get("kind") == "official-venue" and x.get("target_date") == trip.get("date")
            for x in at_sources
        ):
            errors.append(f"attractions[{index}] verified state needs a dated official source")

    itinerary = plan.get("itinerary") or []
    if not isinstance(itinerary, list):
        errors.append("itinerary must be an array")
        itinerary = []
    previous_end = -1
    for index, event in enumerate(itinerary):
        if not isinstance(event, dict):
            errors.append(f"itinerary[{index}] must be an object")
            continue
        start, end = _time(event.get("start")), _time(event.get("end"))
        if start is None or end is None or end <= start:
            errors.append(f"itinerary[{index}] needs a valid same-day start/end")
        elif start < previous_end:
            errors.append(f"itinerary[{index}] overlaps the previous event")
        if end is not None:
            previous_end = end
        if not isinstance(event.get("activity"), str) or not event["activity"].strip():
            errors.append(f"itinerary[{index}].activity is required")
        attraction = event.get("attraction")
        if attraction:
            if attraction not in attraction_names:
                errors.append(f"itinerary[{index}].attraction is not in attractions")
            elif attraction_names[attraction].get("operation") == "verified_closed":
                errors.append(f"itinerary[{index}] schedules a verified-closed attraction")
            if status == "verified_closed":
                errors.append(f"itinerary[{index}] schedules an attraction while venue is closed")

    cautions = plan.get("cautions") or []
    if isinstance(cautions, list):
        for index, item in enumerate(cautions):
            if isinstance(item, dict):
                refs(item, f"cautions[{index}]")
    else:
        errors.append("cautions must be an array")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    args = parser.parse_args()
    try:
        data = json.loads(args.plan.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(json.dumps({"valid": False, "errors": [str(exc)]}, ensure_ascii=False))
        return 1
    errors = validate(data, args.plan.resolve().parent)
    print(json.dumps({"valid": not errors, "errors": errors}, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
