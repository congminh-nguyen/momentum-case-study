#!/usr/bin/env python3
"""Generate expanded, internally consistent GBF datasets.

Export size matches the programme figures (~800 young people served). Sized for
Excel-friendly consulting work while keeping intentional quality issues solvable:
duplicates with last_updated, coding variants, one age/wage outlier, DN cost
double-entry, inactive mentor hours, and a capacity table for scale scenarios.
"""

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

SEED = 20260717
random.seed(SEED)

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "tools" / ".cache" / "raw"

# Sweet spot for mentees: enough for stable hub rates, light enough for Excel pivots.
# Keep in sync with programme guideline + D-01 narrative headcounts.
N_BENEFICIARIES = 800

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

# Mentors M-001..M-044 active; M-099 is inactive but still has logged hours (intentional)
ACTIVE_MENTORS = [f"M-{i:03d}" for i in range(1, 45)]
INACTIVE_MENTOR = "M-099"


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


def gen_beneficiaries(n=N_BENEFICIARIES):
    rows = []
    used_ids = set()
    for i in range(1, n + 1):
        # Mostly 2025 cohort to match "served in 2025"; some 2024 carry-over
        year = 2024 if i <= 220 else 2025
        bid = f"GBF-{year}-{i + 10000:05d}"
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
        if i == 417:  # single intentional impossible age (~42 in 2026)
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
        cohort = f"{prog}-{hub}-{year}-Q{(enroll.month - 1) // 3 + 1}-C{random.randint(1, 5):02d}"
        # Stagger last_updated so duplicate resolution has a clear rule
        updated = enroll + timedelta(days=random.randint(30, 400))
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
            "last_updated": fmt_date(updated, style=1),  # ISO for the "true" row
        })
    # Intentional duplicates: same ID, conflicting province, older last_updated
    dup = rows[0].copy()
    dup["beneficiary_id"] = rows[6]["beneficiary_id"]
    dup["province"] = "HCMC"
    dup["last_updated"] = fmt_date(datetime(2024, 2, 1), style=1)
    rows.append(dup)
    dup2 = rows[10].copy()
    dup2["beneficiary_id"] = rows[25]["beneficiary_id"]
    dup2["province"] = "HCMC" if dup2["province"] != "HCMC" else "Ho Chi Minh City"
    dup2["last_updated"] = fmt_date(datetime(2024, 3, 15), style=1)
    rows.append(dup2)
    return rows


def placement_probability(row):
    """Tuned so strong hubs clear ~65%+ among Completers; DN/LA stay below on formal."""
    base = 0.76
    if row["hub"] == "HCMC1":
        base += 0.02
    if row["hub"] == "HCMC2":
        base += 0.01
    if row["hub"] == "DN":
        base -= 0.20
    if row["hub"] == "LA":
        base -= 0.22
    if row["hub"] == "BD":
        base -= 0.04
    if row["programme"] == "PD":
        base -= 0.10
    if row["programme"] == "BM":
        base -= 0.06
    if row["disability_status"] == "Y":
        base -= 0.22
    if row["completion_status"] != "Completed":
        return None
    if row["province"] in ("Dong Nai", "DN") and random.random() < 0.18:
        return "Gig"
    return "Formal" if random.random() < base else None


def gen_outcomes(beneficiaries):
    rows = []
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


def gen_attendance(beneficiaries):
    """SF/PD: classroom sessions. BM: fewer mentor-circle sessions (so Task DB-3 is possible)."""
    rows = []
    att_codes = ["Y", "N", "1", "0"]
    for b in beneficiaries:
        if b["programme"] == "BM":
            n_sess = random.randint(2, 4)
            prog_prefix = "BM"
            rate = 0.80 if b["completion_status"] == "Completed" else 0.50
        else:
            n_sess = random.randint(5, 8)
            prog_prefix = b["programme"]
            rate = 0.85 if b["completion_status"] == "Completed" else 0.55
        if b["hub"] == "DN":
            rate -= 0.12
        for s in range(1, n_sess + 1):
            base_date = datetime(2025, 3, 1) + timedelta(days=7 * s)
            attended = random.choice(att_codes) if random.random() < rate else "N"
            rows.append({
                "beneficiary_id": b["beneficiary_id"],
                "session_id": f"{prog_prefix}-S-{s:03d}",
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
        if random.random() > 0.72:
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
    for b in beneficiaries:
        if random.random() > 0.38:
            continue
        # ~10% of mentoring assignments go to inactive mentor M-099
        if random.random() < 0.10:
            mid = INACTIVE_MENTOR
            status = "Inactive"
        else:
            mid = random.choice(ACTIVE_MENTORS)
            status = "Active"
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
    # Guarantee a visible inactive block even if sampling missed it
    if not any(r["mentor_id"] == INACTIVE_MENTOR for r in rows):
        for b in beneficiaries[:6]:
            for month in range(3, 7):
                rows.append({
                    "mentor_id": INACTIVE_MENTOR,
                    "beneficiary_id": b["beneficiary_id"],
                    "month": f"2025-{month:02d}",
                    "hours_logged": random.randint(2, 8),
                    "activity_type": "1-on-1",
                    "status": "Inactive",
                })
    return rows


def gen_hub_costs():
    """Monthly operating cost by hub, 2025. Dong Nai September is double-entered."""
    rows = []
    hubs = {"HCMC1": 42, "HCMC2": 34, "BD": 26, "DN": 24, "LA": 15}
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
    dn_sep = next(r for r in rows if r["hub"] == "DN" and r["month"] == "2025-09")
    rows.append(dict(dn_sep))
    return rows


def gen_hub_capacity():
    """Staff / mentor / finance headroom for scale scenarios (Reflection feedback)."""
    # Throughput ~800 total; DN/LA intentionally tight — scaling without capacity fails.
    specs = [
        # hub, staff_fte, active_mentors, avg_caseload, monthly_burn_vnd_m, max_annual_youth, notes
        ("HCMC1", 6, 8, 42, 65, 260, "Strongest hub; limited spare mentor capacity"),
        ("HCMC2", 5, 6, 38, 55, 200, "Can absorb moderate growth with hiring lead time"),
        ("BD", 4, 5, 40, 42, 160, "Border districts contested with YouthWorks"),
        ("DN", 4, 4, 55, 40, 120, "Overloaded caseloads; quality already under strain"),
        ("LA", 2, 3, 48, 22, 60, "Pilot hub; high drop-out; thin staff bench"),
    ]
    rows = []
    for hub, fte, mentors, caseload, burn, cap, notes in specs:
        rows.append({
            "hub": hub,
            "snapshot_date": "2026-01-15",
            "staff_fte": fte,
            "active_mentors": mentors,
            "avg_caseload_per_staff": caseload,
            "monthly_operating_burn_vnd_m": burn,
            "max_annual_throughput_youth": cap,
            "open_field_vacancies": random.randint(0, 3) if hub in ("DN", "LA") else random.randint(0, 1),
            "notes": notes,
        })
    return rows


def gen_funding():
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
            if name == "UK aid / FCDO" and q == 3:
                amt_field = "1,250"  # intentional text-with-comma so SUM fails until cleaned
            rows.append({
                "funder": name,
                "funder_type": ftype,
                "quarter": f"2025-Q{q}",
                "amount_vnd_m": amt_field,
                "restricted": random.choice(["Restricted", "Restricted", "Flexible"]),
            })
    return rows


def gen_waitlist():
    """Waitlist snapshot by district, January 2026. Totals exactly 1,100."""
    districts = [
        ("HCMC", "Binh Tan", 62), ("HCMC", "District 8", 55), ("HCMC", "District 12", 58),
        ("HCMC", "Go Vap", 50), ("HCMC", "Thu Duc", 65), ("HCMC", "Binh Chanh", 48),
        ("Binh Duong", "Di An", 52), ("Binh Duong", "Thuan An", 50), ("Binh Duong", "Thu Dau Mot", 45),
        ("Binh Duong", "Ben Cat", 40), ("Binh Duong", "Tan Uyen", 42),
        ("Dong Nai", "Bien Hoa", 58), ("Dong Nai", "Long Thanh", 50), ("Dong Nai", "Trang Bom", 46),
        ("Dong Nai", "Nhon Trach", 44), ("Dong Nai", "Vinh Cuu", 38),
        ("Long An", "Ben Luc", 36), ("Long An", "Duc Hoa", 34), ("Long An", "Can Giuoc", 32),
        ("Long An", "Tan An", 30),
        ("HCMC", "District 7", 48), ("HCMC", "Nha Be", 36),
        ("Binh Duong", "Bau Bang", 38), ("Dong Nai", "Cam My", 43),
    ]
    rows = []
    for prov, dist, waiting in districts:
        rows.append({
            "province": prov,
            "district": dist,
            "on_waitlist": waiting,
            "avg_wait_days": random.randint(70, 220),
            "youthworks_present": "Y" if prov == "Binh Duong" and random.random() < 0.6 else "N",
            "snapshot_date": "2026-01-15",
        })
    assert sum(r["on_waitlist"] for r in rows) == 1100
    return rows


def write_csv(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    beneficiaries = gen_beneficiaries(N_BENEFICIARIES)
    # Related tables use unique people only; intentional ID collisions live only in the registry
    unique_beneficiaries = beneficiaries[:N_BENEFICIARIES]
    b_fields = [
        "beneficiary_id", "full_name", "date_of_birth", "gender", "ethnicity",
        "disability_status", "province", "programme", "hub", "cohort_id",
        "enrollment_date", "completion_status", "last_updated",
    ]
    write_csv(RAW / "beneficiary_registry.csv", beneficiaries, b_fields)

    outcomes = gen_outcomes(unique_beneficiaries)
    write_csv(
        RAW / "employment_outcomes.csv", outcomes,
        ["beneficiary_id", "programme", "hub", "completion_status", "placed_90d",
         "placement_date", "wage_vnd_monthly", "employment_type", "notes"],
    )

    attendance = gen_attendance(unique_beneficiaries)
    write_csv(
        RAW / "programme_attendance.csv", attendance,
        ["beneficiary_id", "session_id", "session_date", "programme", "attended"],
    )

    survey = gen_survey(unique_beneficiaries)
    write_csv(
        RAW / "survey_outcomes_2025.csv", survey,
        ["beneficiary_id", "programme", "hub", "satisfaction_score", "would_recommend",
         "barrier_transport", "barrier_cost", "barrier_family", "open_comment"],
    )

    volunteers = gen_volunteer_hours(unique_beneficiaries)
    write_csv(
        RAW / "volunteer_hours.csv", volunteers,
        ["mentor_id", "beneficiary_id", "month", "hours_logged", "activity_type", "status"],
    )

    hub_costs = gen_hub_costs()
    write_csv(
        RAW / "hub_costs_2025.csv", hub_costs,
        ["hub", "month", "staff_cost_vnd_m", "programme_cost_vnd_m",
         "transport_support_vnd_m", "total_vnd_m"],
    )

    capacity = gen_hub_capacity()
    write_csv(
        RAW / "hub_capacity_jan2026.csv", capacity,
        ["hub", "snapshot_date", "staff_fte", "active_mentors", "avg_caseload_per_staff",
         "monthly_operating_burn_vnd_m", "max_annual_throughput_youth",
         "open_field_vacancies", "notes"],
    )

    funding = gen_funding()
    write_csv(
        RAW / "funding_by_source.csv", funding,
        ["funder", "funder_type", "quarter", "amount_vnd_m", "restricted"],
    )

    waitlist = gen_waitlist()
    write_csv(
        RAW / "geographic_waitlist.csv", waitlist,
        ["province", "district", "on_waitlist", "avg_wait_days",
         "youthworks_present", "snapshot_date"],
    )

    inactive_n = sum(1 for r in volunteers if r["status"] == "Inactive")
    print(f"Generated {len(beneficiaries)} beneficiaries (target {N_BENEFICIARIES})")
    print(f"Generated {len(outcomes)} outcomes, {len(attendance)} attendance rows")
    print(f"Generated {len(survey)} survey rows, {len(volunteers)} volunteer rows "
          f"({inactive_n} inactive-mentor hours)")
    print(f"Generated {len(hub_costs)} hub-cost rows, {len(capacity)} capacity rows, "
          f"{len(funding)} funding rows, {len(waitlist)} waitlist rows")


if __name__ == "__main__":
    main()
