"""
Data preparation for the GIPA project - ALL YEARS (2017-2024).

Works directly on the ORIGINAL Georgian files (ueedata_2017.xlsx ... ueedata_2024.xlsx),
so no translated copies are needed. Put all the files in the same folder as this script.

Output:
  applicants_gipa_all_years.csv - one row per applicant per year, target = GIPA in top 3
  programs_all_years.csv        - one row per program per year (tuition, places, funding)

Run:  python prepare_all_years.py
"""
import glob
import re

import pandas as pd

from translations_en import NAMES_EN   # Georgian -> English names (separate file)

GIPA_HEI_CODE = "153"   # GIPA's code - checked: the same in every year 2017-2024
TOP_N = 3

# ------------------------------------------------------------------ translations
SHEET_APPLICANTS = "აბიტურიენტები"
SHEET_RESULTS = "შედეგები"
SHEET_SELECTED = "არჩეული პროგრამები"
SHEET_PROGRAMS = "პროგრამები"
SHEET_ENROLLMENTS = "ჩარიცხვები"

GENDER = {"ვაჟი": "Male", "გოგონა": "Female"}

SUBJECTS = {
    "ქართული ენა და ლიტ.": "georgian_language_and_lit",
    "ზოგადი უნარები": "general_aptitude",
    "ინგლისური ენა": "english_language",
    "გერმანული ენა": "german_language",
    "ფრანგული ენა": "french_language",
    "რუსული ენა": "russian_language",
    "აფხაზური ენა": "abkhaz_language",
    "ოსური ენა": "ossetian_language",
    "ისტორია": "history",
    "მათემატიკა": "mathematics",
    "გეოგრაფია": "geography",
    "ბიოლოგია": "biology",
    "ქიმია": "chemistry",
    "ფიზიკა": "physics",
    "ლიტერატურა": "literature",
    "სამოქალაქო განათლება": "civic_education",
    "ხელოვნება": "art",
}

EXAM_LANGUAGE = {"ქართ": "Georgian", "რუს": "Russian", "აზერ": "Azerbaijani",
                 "სომხ": "Armenian", "ოსური": "Ossetian"}

FUNDED_VALUES = {"ფინანსდება", "აფინანსებს სახელმწიფო"}   # "funded" / "funded by the state"

DISTRICTS = {
    'აბაშა': 'Abasha',
    'ადიგენი': 'Adigeni',
    'ამბროლაური': 'Ambrolauri',
    'ასპინძა': 'Aspindza',
    'ახალგორი': 'Akhalgori',
    'ახალქალაქი': 'Akhalkalaki',
    'ახალციხე': 'Akhaltsikhe',
    'ახმეტა': 'Akhmeta',
    'ბათუმი': 'Batumi',
    'ბაღდათი': 'Baghdati',
    'ბოლნისი': 'Bolnisi',
    'ბორჯომი': 'Borjomi',
    'გაგრა': 'Gagra',
    'გალი': 'Gali',
    'გარდაბანი': 'Gardabani',
    'გორი': 'Gori',
    'გუდაუთა': 'Gudauta',
    'გულრიფში': 'Gulripshi',
    'გურჯაანი': 'Gurjaani',
    'დედოფლისწყარო': 'Dedoplistskaro',
    'დმანისი': 'Dmanisi',
    'დუშეთი': 'Dusheti',
    'ვანი': 'Vani',
    'ზესტაფონი': 'Zestaponi',
    'ზუგდიდი': 'Zugdidi',
    'თბილისი': 'Tbilisi',
    'თეთრიწყარო': 'Tetritskaro',
    'თელავი': 'Telavi',
    'თერჯოლა': 'Terjola',
    'თიანეთი': 'Tianeti',
    'კასპი': 'Kaspi',
    'ლაგოდეხი': 'Lagodekhi',
    'ლანჩხუთი': 'Lanchkhuti',
    'ლენტეხი': 'Lentekhi',
    'მარნეული': 'Marneuli',
    'მარტვილი': 'Martvili',
    'მესტია': 'Mestia',
    'მცხეთა': 'Mtskheta',
    'ნინოწმინდა': 'Ninotsminda',
    'ოზურგეთი': 'Ozurgeti',
    'ონი': 'Oni',
    'ოჩამჩირე': 'Ochamchire',
    'რუსთავი': 'Rustavi',
    'საბერძნეთი': 'Greece',
    'საგარეჯო': 'Sagarejo',
    'სამტრედია': 'Samtredia',
    'საჩხერე': 'Sachkhere',
    'სენაკი': 'Senaki',
    'სიღნაღი': 'Sighnaghi',
    'სოხუმი': 'Sokhumi',
    'ტყვარჩელი': 'Tkvarcheli',
    'ტყიბული': 'Tkibuli',
    'ფოთი': 'Poti',
    'ქარელი': 'Kareli',
    'ქედა': 'Keda',
    'ქობულეთი': 'Kobuleti',
    'ქუთაისი': 'Kutaisi',
    'ყაზბეგი': 'Kazbegi',
    'ყვარელი': 'Kvareli',
    'შუახევი': 'Shuakhevi',
    'ჩოხატაური': 'Chokhatauri',
    'ჩხოროწყუ': 'Chkhorotsku',
    'ცაგერი': 'Tsageri',
    'ცხინვალი': 'Tskhinvali',
    'წალენჯიხა': 'Tsalenjikha',
    'წალკა': 'Tsalka',
    'წყალტუბო': 'Tskaltubo',
    'ჭიათურა': 'Chiatura',
    'ხარაგაული': 'Kharagauli',
    'ხაშური': 'Khashuri',
    'ხელვაჩაური': 'Khelvachauri',
    'ხობი': 'Khobi',
    'ხონი': 'Khoni',
    'ხულო': 'Khulo',
    'აფხაზეთი': 'Abkhazia',
    'საქართველო': 'Georgia',
    'ხანია': 'Chania (Greece)',
    'ჯავა': 'Java',
    'დ.': 'D. (unclear in source)',
    'გლდანი': 'Gldani',
    'დიდუბე': 'Didube',
    'ვაკე': 'Vake',
    'ისანი': 'Isani',
    'კრწანისი': 'Krtsanisi',
    'მთაწმინდა': 'Mtatsminda',
    'ნაძალადევი': 'Nadzaladevi',
    'საბურთალო': 'Saburtalo',
    'სამგორი': 'Samgori',
    'ჩუღურეთი': 'Chughureti',
}


# ------------------------------------------------------------------ English names
CITY = {"თბილისი": "Tbilisi", "ბათუმი": "Batumi", "ქუთაისი": "Kutaisi", "ახალციხე": "Akhaltsikhe",
        "ახალქალაქი": "Akhalkalaki", "გორი": "Gori", "თელავი": "Telavi", "ზუგდიდი": "Zugdidi",
        "სიღნაღი": "Sighnaghi", "ფოთი": "Poti", "რუსთავი": "Rustavi", "ოზურგეთი": "Ozurgeti",
        "მარნეული": "Marneuli", "ხულო": "Khulo", "ქობულეთი": "Kobuleti"}


def split_program_name(name):
    """Split a Georgian program name into: core field, city, teaching language.
    e.g. 'ბიზნესის ადმინისტრირება (ინგლისურენოვანი) (ქ. ბათუმი)'
         -> core 'ბიზნესის ადმინისტრირება', city 'Batumi', language 'English-taught'"""
    s = re.sub(r"\s+", " ", str(name).strip().strip(",'\"„“ "))
    lang = None
    if "ინგლისურენოვან" in s:
        lang = "English-taught"
    elif "რუსულენოვან" in s:
        lang = "Russian-taught"
    elif "ინგლისური კომპონენტით" in s:
        lang = "Georgian with English component"
    city = None
    m = re.search(r"\(\s*ქ\.\s*([^\),]+)", s) or re.search(r"\(\s*(" + "|".join(CITY) + r")\s*\)", s)
    if m:
        frag = m.group(1).strip()
        # the source text is sometimes cut off ("ქ. თბ"), so also match the start of a city name
        city = CITY.get(frag) or next((en for ka, en in CITY.items() if ka.startswith(frag)), None)
    core = s if re.match(r"^\s*(\d|ა\))", s) else re.split(r"[\(\.]\s|\(|/| - |\.\s*მოდულ", s)[0]
    core = re.sub(r"(საბაკალავრო\s*)?(საგანმანათლებლო\s*)?(პროგ[რ]?ამა|პროგარამა)$", "", core.strip()).strip()
    core = re.sub(r"\s*(საბაკალავრო|საგანმანათლებლო)$", "", core).strip(" ,.;-")
    core = core.replace("ინგლისურენოვანი", "").replace("რუსულენოვანი", "").strip(" ,.;-")
    return core or s, city, lang


def add_english_names(programs):
    """Adds English columns to the programs table (translation happens AFTER preparation)."""
    ka = programs["program_name_ka"].fillna("").str.strip()
    parts = ka.apply(split_program_name)
    programs["program_field_en"] = [NAMES_EN.get(c, c) for c, _, _ in parts]
    programs["program_city"] = [c for _, c, _ in parts]
    programs["teaching_language"] = [l if l else "Georgian" for _, _, l in parts]
    programs["program_name_en"] = [NAMES_EN.get(full, field) for full, field in zip(ka, programs["program_field_en"])]
    programs["hei_name_en"] = programs["hei_name_ka"].fillna("").str.strip().map(lambda h: NAMES_EN.get(h, h))
    # One stable English name per university code (the name used in its most recent year)
    latest = programs.sort_values("year").groupby("hei_code")["hei_name_en"].last()
    programs["hei_name_standard"] = programs["hei_code"].map(latest)
    order = ["year", "program_key", "hei_code", "hei_name_standard", "hei_name_en", "program_name_en",
             "program_field_en", "program_city", "teaching_language", "places", "tuition", "funded",
             "n_chose_first", "n_chose_top3", "n_chose_any", "n_enrolled", "n_enrolled_with_grant",
             "fill_rate", "avg_enrolled_score", "min_enrolled_score", "is_gipa", "hei_name_ka", "program_name_ka"]
    return programs[order]


# ------------------------------------------------------------------ helpers
def clean_id(series):
    """Exam IDs are stored inconsistently (text with leading zeros vs numbers).
    Remove spaces, a trailing '.0' and leading zeros so they always match."""
    return (series.astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .str.lstrip("0"))


def split_district(raw):
    """'თბილისი/ვაკე' -> district Tbilisi, subdistrict Vake.
    'აფხაზეთი/გაგრა' -> region Abkhazia, district Gagra."""
    if pd.isna(raw):
        return pd.Series({"district": "Unknown", "tbilisi_subdistrict": None})
    parts = [p.strip() for p in str(raw).strip().rstrip(",").split("/")]
    if parts[0] == "თბილისი":
        sub = DISTRICTS.get(parts[1], parts[1]) if len(parts) > 1 else None
        return pd.Series({"district": "Tbilisi", "tbilisi_subdistrict": sub})
    name = parts[-1]
    return pd.Series({"district": DISTRICTS.get(name, name), "tbilisi_subdistrict": None})


# ------------------------------------------------------------------ one year
def prepare_year(path, year):
    sheets = pd.read_excel(path, sheet_name=[SHEET_APPLICANTS, SHEET_RESULTS, SHEET_SELECTED,
                                             SHEET_PROGRAMS, SHEET_ENROLLMENTS], dtype=str)
    for df in sheets.values():
        df.columns = df.columns.astype(str).str.strip()

    # --- programs: aggregate to the 7-digit program (code = program + elective subject digits)
    p = sheets[SHEET_PROGRAMS].dropna(subset=["პროგრამის კოდი"]).copy()
    p["program_key"] = p["პროგრამის კოდი"].str.strip().str[:7]
    p["places"] = pd.to_numeric(p["ადგილების რაოდენობა"], errors="coerce")
    p["tuition"] = pd.to_numeric(p["წლიური გადასახადი"], errors="coerce")
    programs = (p.groupby("program_key")
                .agg(hei_code=("უსდ კოდი", lambda s: s.astype(str).str.strip().iloc[0]),
                     hei_name_ka=("უსდ", "first"),
                     program_name_ka=("პროგრამა", "first"),
                     places=("places", "sum"),
                     tuition=("tuition", "max"),
                     funded=("ფინანსდება", lambda s: s.isin(FUNDED_VALUES).any()))
                .reset_index())
    programs["is_gipa"] = programs["hei_code"] == GIPA_HEI_CODE
    programs.insert(0, "year", year)

    # --- target (column names differ between years, so use positions: ID, priority, program)
    s = sheets[SHEET_SELECTED].iloc[:, :3].copy()
    s.columns = ["exam_id", "priority", "program_key"]
    s = s.dropna(subset=["exam_id", "program_key"])
    s["exam_id"] = clean_id(s["exam_id"])
    s["priority"] = pd.to_numeric(s["priority"], errors="coerce")
    s["is_gipa"] = s["program_key"].astype(str).str.strip().str[:3] == GIPA_HEI_CODE
    target = s.groupby("exam_id").agg(n_choices=("priority", "size"), gipa_any=("is_gipa", "any"))
    target["gipa_top3"] = s[s["priority"] <= TOP_N].groupby("exam_id")["is_gipa"].any()
    target["gipa_top3"] = target["gipa_top3"].fillna(False).astype(bool)

    # --- demand per program: how many applicants chose it, and at which priority
    s["program_key"] = s["program_key"].astype(str).str.strip()
    demand = s.groupby("program_key").agg(
        n_chose_any=("exam_id", "nunique"),
        n_chose_first=("priority", lambda p: (p == 1).sum()),
        n_chose_top3=("priority", lambda p: (p <= TOP_N).sum()),
    )

    # --- enrollments per program: how many were finally admitted (and with a grant)
    e = sheets[SHEET_ENROLLMENTS].copy()
    e["program_key"] = e["პროგრამა"].astype(str).str.strip().str[:7]
    e["grant"] = pd.to_numeric(e["გრანტი"], errors="coerce")
    e["comp_score"] = pd.to_numeric(e["საკონკურსო ქულა"], errors="coerce")
    enrolled = e.groupby("program_key").agg(
        n_enrolled=("საგამოცდო", "size"),
        n_enrolled_with_grant=("grant", lambda g: (g > 0).sum()),
        avg_enrolled_score=("comp_score", "mean"),
        min_enrolled_score=("comp_score", "min"),
    )

    programs = programs.merge(demand, on="program_key", how="left").merge(enrolled, on="program_key", how="left")
    count_cols = ["n_chose_any", "n_chose_first", "n_chose_top3", "n_enrolled", "n_enrolled_with_grant"]
    programs[count_cols] = programs[count_cols].fillna(0).astype(int)
    programs[["avg_enrolled_score", "min_enrolled_score"]] = programs[["avg_enrolled_score", "min_enrolled_score"]].round(1)
    # share of places that were filled (can be above 1 when places is underreported)
    programs["fill_rate"] = (programs["n_enrolled"] / programs["places"].where(programs["places"] > 0)).round(2)

    # --- exam scores: one column per subject
    r = sheets[SHEET_RESULTS].iloc[:, :8].copy()
    r.columns = ["exam_id", "subject_code", "variant", "subject_ka", "raw", "appeal", "equated", "scaled"]
    r = r.dropna(subset=["exam_id", "subject_ka"])
    r["exam_id"] = clean_id(r["exam_id"])
    r["score"] = pd.to_numeric(r["scaled"], errors="coerce")
    base = r["subject_ka"].str.replace(r"\s*\(.*\)\s*$", "", regex=True).str.strip()
    r["subject"] = base.map(SUBJECTS).fillna(base)
    r["exam_lang"] = r["subject_ka"].str.extract(r"\((.*)\)")[0].map(EXAM_LANGUAGE)
    scores = r.pivot_table(index="exam_id", columns="subject", values="score", aggfunc="max")
    took = scores.notna().add_prefix("took_")
    scores = scores.add_prefix("score_")
    aptitude_lang = (r[r["subject"] == "general_aptitude"].groupby("exam_id")["exam_lang"]
                     .first().rename("aptitude_language"))

    # --- applicant details
    a = sheets[SHEET_APPLICANTS].copy()
    a["exam_id"] = clean_id(a["საგამოცდო"])
    a = a.drop_duplicates("exam_id").set_index("exam_id")
    a["gender"] = a["სქესი"].map(GENDER)
    a["age"] = year - pd.to_numeric(a["დაბადების წელი"], errors="coerce")
    a = a.join(a["რაიონი"].apply(split_district))
    a["is_tbilisi"] = a["district"] == "Tbilisi"
    a = a[["gender", "age", "district", "tbilisi_subdistrict", "is_tbilisi"]]

    table = target.join(a, how="left").join(scores, how="left").join(took, how="left") \
                  .join(aptitude_lang, how="left").reset_index()
    took_cols = [c for c in table.columns if c.startswith("took_")]
    table[took_cols] = table[took_cols].fillna(False).astype(bool)
    table.insert(0, "year", year)
    return table, programs


# ------------------------------------------------------------------ all years
files = sorted(f for f in glob.glob("ueedata_20*.xlsx")
               if re.fullmatch(r"ueedata_20\d\d\.xlsx", f.split("/")[-1].split("\\")[-1]))
if not files:
    raise SystemExit("No files found. Put ueedata_20XX.xlsx in the same folder as this script.")

all_tables, all_programs = [], []
for path in files:
    year = int(re.search(r"20\d\d", path).group())
    print(f"Preparing {year} ...", flush=True)
    t, p = prepare_year(path, year)
    all_tables.append(t)
    all_programs.append(p)
    print(f"   {len(t):,} applicants, GIPA in top {TOP_N}: {t['gipa_top3'].sum():,} "
          f"({t['gipa_top3'].mean():.2%}), GIPA programs: {p['is_gipa'].sum()}")

applicants = pd.concat(all_tables, ignore_index=True)
took_cols = [c for c in applicants.columns if c.startswith("took_")]
applicants[took_cols] = applicants[took_cols].fillna(False).astype(bool)
programs = pd.concat(all_programs, ignore_index=True)
print("Adding English names ...")
programs = add_english_names(programs)

# The same exam ID can appear in different years, so (year, exam_id) is the unique key.
applicants.to_csv("applicants_gipa_all_years.csv", index=False, encoding="utf-8-sig")
programs.to_csv("programs_all_years.csv", index=False, encoding="utf-8-sig")

print(f"\nDone: {len(applicants):,} applicant rows, {len(programs):,} program rows")
print(f"GIPA in top {TOP_N} overall: {applicants['gipa_top3'].sum():,} ({applicants['gipa_top3'].mean():.2%})")
