"""Generate a realistic sample call-detail file (45 columns, mixed duration formats)."""
from __future__ import annotations

import datetime as dt
import random
from pathlib import Path

import pandas as pd

SEED = 20260820
FIRST = ["John", "Maria", "Luis", "Ana", "Carlos", "Sofia", "Miguel", "Elena", "Diego",
         "Paula", "Andres", "Karla", "Ruben", "Lucia", "Hector", "Valeria", "Omar",
         "Daniela", "Ivan", "Rocio", "Pablo"]
LAST = ["Smith", "Garcia", "Lopez", "Hernandez", "Ramirez", "Torres", "Flores", "Vargas",
        "Mendoza", "Castillo", "Reyes", "Cruz", "Morales", "Ortega", "Delgado"]

HEADERS = [
    "Transaction ID", "Site", "Agent Name", "Rec Type", "Media", "Direction",
    "ANI", "DNIS", "Queue", "Start Date Time", "End Date Time", "Call Duration",
    "Hold Time", "Talk Time", "Wrap Time", "Segment", "Disposition", "Skill",
    "Team", "Supervisor", "LOB", "Campaign", "Channel", "Transfer Flag",
    "Hold Count", "Silence %", "Sentiment", "Language", "Device", "Codec",
    "Server", "Cluster", "Region", "Login ID", "Extension", "Workgroup",
    "Program", "Vendor", "Connection ID", "GenConnID", "Post_Route_Data",
    "Survey Score", "Notes", "Archive Path", "Retention Days",
]


def duration_variant(seconds: int, style: int):
    """Return the duration in one of several Excel representations."""
    if style == 0:                                   # HH:MM:SS string
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    if style == 1:                                   # Excel fractional day
        return round(seconds / 86400.0, 8)
    if style == 2:                                   # datetime.time
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        return dt.time(h, m, s)
    if style == 3:                                   # MM:SS string
        m, s = divmod(seconds, 60)
        return f"{m:02d}:{s:02d}"
    return round(seconds / 60.0, 2)                  # numeric minutes


def build(rows: int = 12500, agents: int = 42, out: str = "sample_call_detail.xlsx") -> Path:
    rng = random.Random(SEED)
    names = []
    while len(names) < agents:
        n = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
        if n not in names:
            names.append(n)

    base = dt.datetime(2026, 8, 3, 7, 0, 0)
    codes = ["ABC123", "XYZ456", "TEST789", "NEM001", "CBB404", "LPC220"]
    data = []
    for i in range(rows):
        agent = rng.choice(names)
        # occasional messy agent value (extra spaces) - must group as the same agent
        if rng.random() < 0.02:
            agent = f"  {agent.replace(' ', '  ')} "
        if rng.random() < 0.006:
            agent = ""  # missing agent -> excluded + reported

        r = rng.random()
        if r < 0.55:
            secs = rng.randint(300, 1200)
        elif r < 0.72:
            secs = rng.randint(1260, 2100)
        elif r < 0.86:
            secs = rng.randint(2160, 5400)
        else:
            secs = rng.randint(15, 299)          # too short to sample
        if rng.random() < 0.03:
            secs = rng.choice([rng.randint(1201, 1259), rng.randint(2101, 2159)])  # gap calls

        duration = duration_variant(secs, rng.choice([0, 0, 0, 1, 1, 2, 3, 4]))
        if rng.random() < 0.005:
            duration = rng.choice(["", "N/A", "--"])   # invalid duration -> reported

        rec_type = "Voice And Screen" if rng.random() < 0.66 else "Voice Only"
        if rec_type == "Voice And Screen" and rng.random() < 0.15:
            rec_type = rng.choice(["voice and screen", "VOICE AND SCREEN", " Voice And Screen "])

        start = base + dt.timedelta(minutes=rng.randint(0, 20160), seconds=rng.randint(0, 59))
        code = rng.choice(codes)
        post_route = (
            f"ROUTE={code}|TYPE={rng.choice(['TRANSFER', 'DIRECT', 'CALLBACK'])}"
            f"|SEG={rng.randint(1, 4)}"
        )
        if rng.random() < 0.04:
            post_route = ""
        if rng.random() < 0.05:
            post_route = f"{code}"  # exact-match candidate

        gen = f"GC{1000000 + i}"
        if rng.random() < 0.004 and i > 10:
            gen = f"GC{1000000 + i - 1}"   # duplicate GenConnID -> must not duplicate output

        row = {h: "" for h in HEADERS}
        row["Transaction ID"] = f"TXN{500000 + i}"
        row["Site"] = rng.choice(["GDL", "MTY", "CDMX"])
        row["Agent Name"] = agent
        row["Rec Type"] = rec_type
        row["Media"] = "Voice"
        row["Direction"] = "Inbound"
        row["ANI"] = f"555{rng.randint(1000000, 9999999)}"
        row["DNIS"] = f"800{rng.randint(1000000, 9999999)}"
        row["Queue"] = rng.choice(["BILLING", "OUTAGE", "START_STOP", "PAYMENTS"])
        row["Start Date Time"] = start
        row["End Date Time"] = start + dt.timedelta(seconds=secs)
        row["Call Duration"] = duration
        row["Hold Time"] = rng.randint(0, 240)
        row["Talk Time"] = max(secs - rng.randint(0, 120), 0)
        row["Wrap Time"] = rng.randint(10, 180)
        row["Segment"] = rng.randint(1, 3)
        row["Disposition"] = rng.choice(["Resolved", "Escalated", "Callback", "Transferred"])
        row["Skill"] = rng.choice(["EN", "ES"])
        row["Team"] = f"Team {rng.randint(1, 8)}"
        row["Supervisor"] = rng.choice(names)
        row["LOB"] = "SCE"
        row["Campaign"] = "SCE Customer Care"
        row["Channel"] = "Phone"
        row["Transfer Flag"] = rng.choice(["Y", "N"])
        row["Hold Count"] = rng.randint(0, 4)
        row["Silence %"] = round(rng.random() * 20, 1)
        row["Sentiment"] = rng.choice(["Positive", "Neutral", "Negative"])
        row["Language"] = rng.choice(["English", "Spanish"])
        row["Device"] = "WebRTC"
        row["Codec"] = "G711"
        row["Server"] = f"SRV-{rng.randint(1, 12):02d}"
        row["Cluster"] = f"CL{rng.randint(1, 4)}"
        row["Region"] = "West"
        row["Login ID"] = f"U{200000 + names.index(agent.strip().replace('  ', ' ')) if agent.strip() else 0}"
        row["Extension"] = rng.randint(3000, 3999)
        row["Workgroup"] = "WG-CARE"
        row["Program"] = "Care"
        row["Vendor"] = "TCS"
        row["Connection ID"] = f"CN{i}"
        row["GenConnID"] = gen
        row["Post_Route_Data"] = post_route
        row["Survey Score"] = rng.choice(["", 1, 2, 3, 4, 5])
        row["Notes"] = ""
        row["Archive Path"] = f"\\\\archive\\{start:%Y%m%d}\\{gen}.wav"
        row["Retention Days"] = 90
        data.append(row)

    df = pd.DataFrame(data, columns=HEADERS)
    path = Path(out)
    df.to_excel(path, index=False)
    return path


if __name__ == "__main__":
    p = build()
    print(f"Wrote {p} ({p.stat().st_size / 1024:.0f} KB)")
