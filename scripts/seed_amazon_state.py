import json
import sys
import time
from pathlib import Path


def main():
    state_path = Path(sys.argv[1])
    csv_paths = [Path(value) for value in sys.argv[2:]]
    state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else {}
    records = []
    for path in csv_paths:
        text = path.read_text(encoding="utf-8-sig")
        records.append({"store": "亚马逊", "name": path.name, "text": text, "trafficOnly": True})
    state["amazon"] = {
        "record": records,
        "file_name": "、".join(path.name for path in csv_paths),
        "uploaded_at": int(time.time() * 1000),
    }
    state_path.write_text(json.dumps(state, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(json.dumps({"files": len(records), "names": [path.name for path in csv_paths], "bytes": state_path.stat().st_size}, ensure_ascii=False))


if __name__ == "__main__":
    main()
