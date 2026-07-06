#!/usr/bin/env python3
"""Generate expanded, internally consistent GBF datasets."""

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

SEED = 20260706
random.seed(SEED)

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "tools" / ".cache" / "raw"

PROVINCES = {
    "HCMC": {"label": "Ho Chi Minh City", "alt": "HCMC", "hub": ["HCMC1", "HCMC2"], "weight": 0.38},
    "BD": {"label": "Binh Duong", "alt": "Binh Duong", "hub": ["BD"], "weight": 0.24},
    "DN": {"label": "Dong Nai", "alt": "Dong Nai", "hub": ["DN"], "weight": 0.28},
    "LA": {"label": "Long An", "alt": "Long An", "hub": ["LA"], "weight": 0.10},
}

ETHNICITIES = ["Kinh"] * 82 + ["Khmer", "Hoa", "Tay", "Cham", ""] * 4
PROGRAMMES = [("SF", 0.78), ("PD", 0.17), ("BM", 0.05)]
COMPLETION = [("Completed", 0.62), ("Dropped", 0.26), ("Active", 0.12)]
FIRST_NAMES_M = ["An", "Binh", "Cuong", "Duc", "Huy", "Khoa", "Long", "Minh", "Nam", "Phuc", "Son", "Tam", "Vinh", "Xuan"]
FIRST_NAMES_F = ["Mai", "Lan", "Binh", "Dung", "Em", "Hoa", "Lan", "Ngoc", "Oanh", "Phuong", "Quynh", "Tuyet", "Uyen", "Van"]
LAST_NAMES = ["Nguyen", "Tran", "Le", "Pham", "Hoang", "Vo", "Dang", "Bui", "Do", "Ngo", "Ly", "Truong"]


def pick_province():
    r = random.random()
    c = 0
    for k, v in PROVINCES.items():
        c += v["weight"]
        if r <= c:
            return k
    return "HCMC"


def fmt_date(d, style=None):
  styles = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
  return d.strftime(styles[style if style is not None else random.randint(0, 2)])


def gen_beneficiaries(n=520):
    rows = []
    used_ids = set()
    for i in range(1, n + 1):
        year = 2024 if i <= 280 else 2025
        bid = f"GBF-{year}-{i+10000:05d}"
        if bid in used_ids:
            continue
        used_ids.add(bid)
        prov_key = pick_province()
        prov = PROVINCES[prov_key]
        gender = random.choice(["M", "F", "F", "M"])
        fn = random.choice(FIRST_NAMES_M if gender == "M" else FIRST_NAMES_F)
        ln = random.choice(LAST_NAMES)
        name = f"{ln} {'Thi' if gender == 'F' else 'Van'} {fn}"
        age = random.randint(16, 24)
        if i == 417:  # intentional age error
            dob = datetime(1984, 3, 15)
        else:
            dob = datetime(2026 - age, random.randint(1, 12), random.randint(1, 28))
        eth = random.choice(ETHNICITIES)
        disability = "Y" if random.random() < 0.06 else "N"
        if random.random() < 0.18:
            eth = ""
        prog = random.choices([p[0] for p in PROGRAMMES], [p[1] for p in PROGRAMMES])[0]
        hub = random.choice(prov["hub"])
        if prog == "PD":
            hub = "HCMC1"
        enroll = datetime(year, random.randint(1, 11), random.randint(1, 28))
        if prov_key == "LA":
            comp_weights = [("Completed", 0.40), ("Dropped", 0.45), ("Active", 0.15)]
        elif prov_key == "DN":
            comp_weights = [("Completed", 0.55), ("Dropped", 0.35), ("Active", 0.10)]
        else:
            comp_weights = COMPLETION
        completion = random.choices([c[0] for c in comp_weights], [c[1] for c in comp_weights])[0]
        province_label = prov["label"] if random.random() > 0.15 else prov["alt"]
        cohort = f"{prog}-{hub}-{year}-Q{(enroll.month-1)//3+1}-C{random.randint(1,5):02d}"
        rows.append({
            "beneficiary_id": bid,
            "full_name": name,
            "date_of_birth": fmt_date(dob),
            "gender": gender,
            "ethnicity": eth,
            "disability_status": disability if random.random() > 0.05 else "",
            "province": province_label,
            "programme": prog,
            "hub": hub,
            "cohort_id": cohort,
            "enrollment_date": fmt_date(enroll),
            "completion_status": completion,
        })
    # intentional duplicates
    dup = rows[0].copy()
    dup["beneficiary_id"] = rows[6]["beneficiary_id"]
    dup["province"] = "HCMC"
    rows.append(dup)
    dup2 = rows[10].copy()
    dup2["beneficiary_id"] = rows[25]["beneficiary_id"]
    rows.append(dup2)
    return rows


def placement_probability(row):
    base = 0.68
    if row["hub"] == "DN":
        base -= 0.14
    if row["hub"] == "LA":
        base -= 0.18
    if row["programme"] == "PD":
        base -= 0.12
    if row["disability_status"] == "Y":
        base -= 0.22
    if row["completion_status"] != "Completed":
        return None
    if row["province"] in ("Dong Nai", "DN") and random.random() < 0.1:
        return "Gig"
    return "Formal" if random.random() < base else None


def gen_outcomes(beneficiaries):
    rows = []
    codes = ["Y", "Yes", "1", "N", "No", "0", ""]
    for b in beneficiaries:
        if b["completion_status"] == "Active":
            continue
        placed_type = placement_probability(b)
        if b["completion_status"] == "Dropped":
            placed = random.choice(["N", "No", "0", ""])
        elif placed_type:
            placed = random.choice(["Y", "Yes", "1"])
        else:
            placed = random.choice(["N", "No", "0"])
        wage = ""
        pdate = ""
        etype = ""
        notes = ""
        if placed in ("Y", "Yes", "1"):
            wage = random.randint(4200000, 6500000)
            if b["beneficiary_id"] == "GBF-2025-10417":
                wage = 99999999
                notes = "Data entry error wage"
            pdate = fmt_date(datetime(2025, random.randint(6, 12), random.randint(1, 28)))
            etype = placed_type or "Formal"
        rows.append({
            "beneficiary_id": b["beneficiary_id"],
            "programme": b["programme"],
            "hub": b["hub"],
            "completion_status": b["completion_status"],
            "placed_90d": placed,
            "placement_date": pdate,
            "wage_vnd_monthly": wage,
            "employment_type": etype,
            "notes": notes,
        })
    return rows


def gen_attendance(beneficiaries, sessions_per=8):
    rows = []
    att_codes = ["Y", "N", "1", "0"]
    for b in beneficiaries:
        if b["programme"] == "BM":
            continue
        n_sess = random.randint(5, 10)
        for s in range(1, n_sess + 1):
            base_date = datetime(2025, 3, 1) + timedelta(days=7 * s)
            rate = 0.85 if b["completion_status"] == "Completed" else 0.55
            if b["hub"] == "DN":
                rate -= 0.12
            attended = random.choice(att_codes) if random.random() < rate else "N"
            rows.append({
                "beneficiary_id": b["beneficiary_id"],
                "session_id": f"{b['programme']}-S-{s:03d}",
                "session_date": fmt_date(base_date),
                "programme": b["programme"],
                "attended": attended,
            })
    return rows


def gen_survey(beneficiaries):
    rows = []
    for b in beneficiaries:
        if b["completion_status"] not in ("Completed", "Dropped"):
            continue
        if random.random() > 0.75:
            continue
        score = random.randint(2, 5)
        if b["hub"] == "DN" and b["completion_status"] == "Dropped":
            score = random.randint(1, 3)
        rows.append({
            "beneficiary_id": b["beneficiary_id"],
            "programme": b["programme"],
            "hub": b["hub"],
            "satisfaction_score": score,
            "would_recommend": random.choice(["Yes", "No", "Maybe"]),
            "barrier_transport": 1 if b["hub"] in ("DN", "LA") and random.random() < 0.6 else 0,
            "barrier_cost": 1 if random.random() < 0.35 else 0,
            "barrier_family": 1 if random.random() < 0.2 else 0,
            "open_comment": random.choice([
                "", "", "Good programme", "Transport too expensive",
                "Employer treated us poorly", "Mentor was helpful",
                "Waited too long to join", "Want more digital skills",
            ]),
        })
    return rows


def gen_volunteer_hours(beneficiaries):
    rows = []
    mentors = [f"M-{i:03d}" for i in range(1, 45)]
    for b in beneficiaries:
        if random.random() > 0.35:
            continue
        mid = random.choice(mentors)
        status = "Inactive" if mid == "M-099" else "Active"
        for month in range(3, 9):
            hrs = random.randint(0, 12) if b["completion_status"] == "Completed" else random.randint(0, 4)
            rows.append({
                "mentor_id": mid,
                "beneficiary_id": b["beneficiary_id"],
                "month": f"2025-{month:02d}",
                "hours_logged": hrs,
                "activity_type": random.choice(["1-on-1", "Group"]),
                "status": status,
            })
    return rows


def gen_hub_costs():
    """Monthly operating cost by hub, 2025. Dong Nai Q3 is double-entered (error)."""
    rows = []
    hubs = {"HCMC1": 118, "HCMC2": 96, "BD": 74, "DN": 68, "LA": 41}  # base VND million/month
    for hub, base in hubs.items():
        for month in range(1, 13):
            staff = base + random.randint(-6, 6)
            programme = round(staff * random.uniform(0.45, 0.6))
            transport = random.randint(2, 7)
            rows.append({
                "hub": hub,
                "month": f"2025-{month:02d}",
                "staff_cost_vnd_m": staff,
                "programme_cost_vnd_m": programme,
                "transport_support_vnd_m": transport,
                "total_vnd_m": staff + programme + transport,
            })
    # intentional error: Dong Nai September double-entered
    dn_sep = next(r for r in rows if r["hub"] == "DN" and r["month"] == "2025-09")
    rows.append(dict(dn_sep))
    return rows


def gen_funding():
    """Quarterly income by funder, 2025 (VND million). One row has a comma-format error."""
    funders = [
        ("VPBank Foundation", "Corporate", 640),
        ("Prudence Foundation", "Corporate", 210),
        ("UK aid / FCDO", "Bilateral", 380),
        ("Irish Aid", "Bilateral", 150),
        ("Save the Children VN", "INGO", 260),
        ("Plan International", "INGO", 175),
        ("Individual giving", "Philanthropy", 130),
        ("Corporate CSR (misc.)", "Corporate", 300),
        ("Provincial co-funding", "Government", 170),
        ("GBF Staffing Solutions", "Enterprise", 110),
        ("Community fundraising", "Philanthropy", 55),
        ("Alumni network", "Philanthropy", 30),
    ]
    rows = []
    for name, ftype, base in funders:
        for q in range(1, 5):
            amount = base + random.randint(-20, 25)
            amt_field = amount
            # intentional error: one row stored as text with a thousands comma
            if name == "UK aid / FCDO" and q == 3:
                amt_field = f"{amount:,}"  # e.g. "1,234" -> breaks SUM
            rows.append({
                "funder": name,
                "funder_type": ftype,
                "quarter": f"2025-Q{q}",
                "amount_vnd_m": amt_field,
                "restricted": random.choice(["Restricted", "Restricted", "Flexible"]),
            })
    return rows


def gen_waitlist():
    """Waitlist snapshot by district, January 2026. Totals ~2,400."""
    districts = [
        ("HCMC", "Binh Tan"), ("HCMC", "District 8"), ("HCMC", "District 12"),
        ("HCMC", "Go Vap"), ("HCMC", "Thu Duc"), ("HCMC", "Binh Chanh"),
        ("Binh Duong", "Di An"), ("Binh Duong", "Thuan An"), ("Binh Duong", "Thu Dau Mot"),
        ("Binh Duong", "Ben Cat"), ("Binh Duong", "Tan Uyen"),
        ("Dong Nai", "Bien Hoa"), ("Dong Nai", "Long Thanh"), ("Dong Nai", "Trang Bom"),
        ("Dong Nai", "Nhon Trach"), ("Dong Nai", "Vinh Cuu"),
        ("Long An", "Ben Luc"), ("Long An", "Duc Hoa"), ("Long An", "Can Giuoc"),
        ("Long An", "Tan An"),
        ("HCMC", "District 7"), ("HCMC", "Nha Be"),
        ("Binh Duong", "Bau Bang"), ("Dong Nai", "Cam My"),
    ]
    rows = []
    for prov, dist in districts:
        waiting = random.randint(60, 160)
        rows.append({
            "province": prov,
            "district": dist,
            "on_waitlist": waiting,
            "avg_wait_days": random.randint(70, 220),
            "youthworks_present": "Y" if prov == "Binh Duong" and random.random() < 0.6 else "N",
            "snapshot_date": "2026-01-15",
        })
    return rows


def write_csv(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    beneficiaries = gen_beneficiaries(520)
    b_fields = ["beneficiary_id", "full_name", "date_of_birth", "gender", "ethnicity",
                "disability_status", "province", "programme", "hub", "cohort_id",
                "enrollment_date", "completion_status"]
    write_csv(RAW / "beneficiary_registry.csv", beneficiaries, b_fields)

    outcomes = gen_outcomes(beneficiaries)
    o_fields = ["beneficiary_id", "programme", "hub", "completion_status", "placed_90d",
                "placement_date", "wage_vnd_monthly", "employment_type", "notes"]
    write_csv(RAW / "employment_outcomes.csv", outcomes, o_fields)

    attendance = gen_attendance(beneficiaries)
    write_csv(RAW / "programme_attendance.csv", attendance,
              ["beneficiary_id", "session_id", "session_date", "programme", "attended"])

    survey = gen_survey(beneficiaries)
    write_csv(RAW / "survey_outcomes_2025.csv", survey,
              ["beneficiary_id", "programme", "hub", "satisfaction_score", "would_recommend",
               "barrier_transport", "barrier_cost", "barrier_family", "open_comment"])

    volunteers = gen_volunteer_hours(beneficiaries)
    write_csv(RAW / "volunteer_hours.csv", volunteers,
              ["mentor_id", "beneficiary_id", "month", "hours_logged", "activity_type", "status"])

    hub_costs = gen_hub_costs()
    write_csv(RAW / "hub_costs_2025.csv", hub_costs,
              ["hub", "month", "staff_cost_vnd_m", "programme_cost_vnd_m",
               "transport_support_vnd_m", "total_vnd_m"])

    funding = gen_funding()
    write_csv(RAW / "funding_by_source.csv", funding,
              ["funder", "funder_type", "quarter", "amount_vnd_m", "restricted"])

    waitlist = gen_waitlist()
    write_csv(RAW / "geographic_waitlist.csv", waitlist,
              ["province", "district", "on_waitlist", "avg_wait_days",
               "youthworks_present", "snapshot_date"])

    print(f"Generated {len(beneficiaries)} beneficiaries")
    print(f"Generated {len(outcomes)} outcomes, {len(attendance)} attendance rows")
    print(f"Generated {len(survey)} survey rows, {len(volunteers)} volunteer rows")
    print(f"Generated {len(hub_costs)} hub-cost rows, {len(funding)} funding rows, "
          f"{len(waitlist)} waitlist rows")


if __name__ == "__main__":
    main()
