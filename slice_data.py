import json
from pathlib import Path

root = Path(__file__).resolve().parent
src = json.loads((root / "build" / "data.json").read_text(encoding="utf-8"))
meta = src["meta"]
schemas = src["schemas"]

def dump(name, obj):
    (root / "build" / name).write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

dump("data_new.json", {
    "meta": meta, "schemas": {"channelActual": schemas["channelActual"], "clientActual": schemas["clientActual"], "clientComparison": schemas["clientComparison"]},
    "channels": {"newDetails": src["channels"]["newDetails"], "groupNew": src["channels"]["groupNew"]},
    "clients": {"newDetails": src["clients"]["newDetails"], "groupNew": src["clients"]["groupNew"], "top100": src["clients"]["top100"], "top500": src["clients"]["top500"]},
})
dump("data_channel_comp.json", {
    "meta": meta, "schemas": {"channelActual": schemas["channelActual"], "channelComparison": schemas["channelComparison"]},
    "channels": {"fyDetails": src["channels"]["fyDetails"], "newDetails": src["channels"]["newDetails"], "comparison": src["channels"]["comparison"]},
})
dump("data_client_actuals.json", {
    "meta": meta, "schemas": {"clientActual": schemas["clientActual"]},
    "clients": {"fyDetails": src["clients"]["fyDetails"], "newDetails": src["clients"]["newDetails"]},
})
dump("data_client_comp.json", {
    "meta": meta, "schemas": {"clientComparison": schemas["clientComparison"]},
    "clients": {"matched": src["clients"]["matched"], "new": src["clients"]["new"], "lost": src["clients"]["lost"], "top100": src["clients"]["top100"], "top500": src["clients"]["top500"], "gainers": src["clients"]["gainers"], "decliners": src["clients"]["decliners"]},
})
print("slices written")
