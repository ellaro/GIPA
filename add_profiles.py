"""
Adds the translated program profiles (2019) to our programs table.

Input : programs_all_years.csv            (created by prepare_all_years.py)
        program_profiles*.xlsx            (the profiles file translated by the team)
Output: programs_2019_with_profiles.csv   (one row per profiled program, 2019)

Run AFTER prepare_all_years.py:   python add_profiles.py
"""
import glob
import os

import numpy as np
import pandas as pd

# ------------------------------------------------------------------ find the files
profile_files = sorted(glob.glob("program_profiles*.xlsx"), key=os.path.getmtime)
if not profile_files:
    raise SystemExit("No program_profiles*.xlsx file found in this folder.")
PROFILES_FILE = profile_files[-1]          # the most recently saved one
print(f"Using profiles file: {PROFILES_FILE}")

prof = pd.read_excel(PROFILES_FILE)
prof.columns = prof.columns.str.strip()
if prof.astype(str).apply(lambda c: c.str.contains("[\u10A0-\u10FF]")).any().any():
    print("WARNING: this file still contains Georgian text - is it the translated version?")

# ------------------------------------------------------------------ 1. join key
# PROGRAM_CODE lost its leading zeros (10112 -> "0010112"). Don't use UNIVERSITY CODE, it is wrong.
prof["program_key"] = prof["PROGRAM_CODE"].astype(int).astype(str).str.zfill(7)

# ------------------------------------------------------------------ 2. clean-up maps
YES_NO = {"Yes": 1, "No": 0}

PROGRAM_FORMAT = {   # many spellings -> 4 clear groups
    "Local program with local lecturers only": "Local, local lecturers only",
    "Local program, local lecturers only": "Local, local lecturers only",
    "Local program, with regular visits by foreign lecturers": "Local, with foreign lecturers",
    "Local program, with regular involvement of foreign lecturers": "Local, with foreign lecturers",
    "Program in cooperation with a foreign university, with regular visits by foreign lecturers": "Foreign partner, with foreign lecturers",
    "Program in cooperation with a foreign university, without regular visits by foreign lecturers": "Foreign partner, no foreign lecturers",
    "Program in cooperation with a foreign university": "Foreign partner, no foreign lecturers",
    "Joint program with a foreign university (Georgian diploma)": "Joint program",
    "Joint program with a foreign university": "Joint program",
}

CONDITION = {        # building / campus condition as a 1-4 score
    "Needs renovation": 1,
    "In normal condition": 2, "Normal": 2,
    "Good": 3, "Newly renovated (old-style)": 3,
    "Newly renovated (modern)": 4,
}

LOCATION = {
    "In the center": "City center", "City center": "City center",
    "In a prestigious district of the city (not the center)": "Prestigious district",
    "In a prestigious location": "Prestigious district",
    "In a non-prestigious district of the city (not in the suburbs or the center)": "Other district",
    "In the suburbs": "Suburbs",
}

EVENTS_PER_YEAR = {  # frequency text -> approx. number of events per academic year
    "Weekly": 36, "2-3 times per month": 22, "Monthly": 9,
    "2-3 times per semester": 5, "Once per semester": 2, "2-3 times a year": 2.5,
    "Once a year": 1, "Not organized": 0, "No": 0,
    # "Yes" / "Other" say nothing about frequency -> left empty
}


def to_number(series):
    """Text -> number. Handles '4,5' (comma decimal) and '-' (missing)."""
    return pd.to_numeric(series.astype(str).str.strip().str.replace(",", ".", regex=False)
                         .replace({"-": np.nan, "nan": np.nan}), errors="coerce")


def fix_thousands(series):
    """Counts like 7.132 were typed with a European thousands separator -> 7132."""
    s = to_number(series)
    return s.where(s.isna() | (s % 1 == 0) | (s >= 1000), s * 1000).round()


def zero_to_missing(series):
    """In this file 0 usually means 'no information', not a real zero."""
    return to_number(series).replace(0, np.nan)


# ------------------------------------------------------------------ 3. build the clean table
c = pd.DataFrame({"program_key": prof["program_key"]})

# program level
c["prog_year_started"] = zero_to_missing(prof["YEAR_STARTED"])
c["prog_format"] = prof["PROGRAM_FORMAT"].str.strip().map(PROGRAM_FORMAT).fillna("Other/Unknown")
c["prog_has_foreign_partner"] = c["prog_format"].str.startswith(("Foreign partner", "Joint")).astype(int)
c["prog_has_foreign_lecturers"] = c["prog_format"].str.contains("with foreign lecturers").astype(int)
c["prog_building_condition"] = prof["BUILDING_CONDITION"].str.strip().map(CONDITION)
c["prog_building_location"] = prof["BUILDING_LOCATION"].str.strip().map(LOCATION).fillna("Other/Unknown")
for col in ["HAS_STUDENT_SPACES", "HAS_CAFFETERIA", "HAS_OUTSIDE_SPACE", "HAS_FREE_PARKING", "HAS_PAID_PARKING"]:
    c["prog_" + col.lower()] = prof[col].str.strip().map(YES_NO)
c["prog_practice_share"] = zero_to_missing(prof["PRACTICE_SHARE_CURRICULUM"])
c["prog_employment_rate"] = zero_to_missing(prof["EMPLOYEMENT_PERCENTAGE"])

# university level
c["uni_year_founded"] = zero_to_missing(prof["YEAR_UNI_FOUNDED"])
c["uni_total_students"] = zero_to_missing(prof["UNI_NO_OF_TOTAL_STUDENTS"])
c["uni_staff"] = zero_to_missing(prof["UNI_NO_OF_STAFF"])
c["uni_admin_staff"] = zero_to_missing(prof["UNI_NO_OF_ADMIN_STAFF"])
c["uni_professors"] = zero_to_missing(prof["UNI_NO_OF_PROFESSORS"])
c["uni_students_per_professor"] = (c["uni_total_students"] / c["uni_professors"]).round(1)
c["uni_fb_likes"] = fix_thousands(prof["FB_LIKES_UNI"])
c["uni_fb_followers"] = fix_thousands(prof["FB_UNI_FOLLOWERS"])
c["uni_fb_rating"] = zero_to_missing(prof["UNI_FB_RATING"])
c["uni_fb_rating_voters"] = zero_to_missing(prof["UNI_RATING_VOTERS"])
c["uni_web_global_rank"] = to_number(prof["UNI_RATING_GLOBAL_RANK"])     # lower = more web traffic
c["uni_web_local_rank"] = fix_thousands(prof["UNI__RATING_LOCAL_RANK"])
c["uni_campuses"] = to_number(prof["UNI_NO_OF_CAMPUSES"])
c["uni_campuses_in_tbilisi"] = to_number(prof["UNI_NO_OF_CAMPUSES_IN_TBILISI"])
c["uni_campus_condition"] = prof["CAMPUS_CONDITION"].str.strip().map(CONDITION)
c["uni_campus_location"] = prof["CAMPUS_LOCATION"].str.strip().map(LOCATION).fillna("Other/Unknown")
campus_cols = {"HAS_STUDENT_HOUSING": "uni_has_student_housing", "HAS_STUDENT_SPACES.1": "uni_has_student_spaces",
               "HAS_CAFFETERIA.1": "uni_has_caffeteria", "HAS_OUTSIDE_SPACE.1": "uni_has_outside_space",
               "HAS_GYM": "uni_has_gym", "HAS_POOL": "uni_has_pool", "HAS_STADIUM": "uni_has_stadium",
               "HAS_FREE_PARKING.1": "uni_has_free_parking", "HAS_PAID_PARKING.1": "uni_has_paid_parking"}
for src, dst in campus_cols.items():
    c[dst] = prof[src].str.strip().map(YES_NO)

# money
c["uni_has_scholarship"] = prof["HAS_SCHOLARSHIP"].str.strip().map(YES_NO)
c["uni_max_scholarship"] = to_number(prof["MAX_SCHOLARSHIP"])
c.loc[(c["uni_has_scholarship"] == 1) & (c["uni_max_scholarship"] == 0), "uni_max_scholarship"] = np.nan
c["uni_has_financing"] = prof["HAS_FINANCING"].str.strip().map(YES_NO)
c["uni_max_financing"] = to_number(prof["MAX_FINANCING"])
c.loc[(c["uni_has_financing"] == 1) & (c["uni_max_financing"] == 0), "uni_max_financing"] = np.nan
c["uni_financing_type"] = prof["FINANCING_TYPE"].astype(str).str.strip().replace({"0": "Unknown", "nan": "Unknown"})

# international
for region in ["USA", "WEST_EUROPE", "EAST_EUROPE", "EAST_ASIA", "FUSSR", "OTHER"]:
    c["uni_exchange_" + region.lower()] = prof["HAS_EXCHANGE_" + region].str.strip().map(YES_NO)
c["uni_exchange_regions"] = c[[x for x in c.columns if x.startswith("uni_exchange_")]].sum(axis=1, min_count=1)
c["uni_exchange_students"] = to_number(prof["NO_EXCHANGE_STUDENTS"]).fillna(
    prof["NO_EXCHANGE_STUDENTS"].astype(str).str.contains("not currently", case=False).map({True: 0}))

# student life (events per year)
for col in ["HOLDS_CONFERENCES", "STUDENT_LIFE_BY_ADMIN", "STUDENT_LIFE_BY_STUDENTS",
            "SPORT_LIFE_BY_ADMIN", "SPORT_LIFE_BY_STUDENTS"]:
    c["uni_" + col.lower() + "_per_year"] = prof[col].astype(str).str.strip().map(EVENTS_PER_YEAR)
c["uni_number_of_programs"] = to_number(prof["UNI_NO_PROGRAMS"])

# ------------------------------------------------------------------ 4. join to the 2019 programs
programs = pd.read_csv("programs_all_years.csv", dtype={"program_key": str, "hei_code": str})
# If the CSV was opened and saved in Excel, leading zeros are lost ("0010102" -> "10102")
# and True/False becomes the text TRUE/FALSE. Repair both so the join still works.
programs["program_key"] = programs["program_key"].str.strip().str.zfill(7)
programs["hei_code"] = programs["hei_code"].str.strip().str.zfill(3)
for col in ["is_gipa", "funded"]:
    programs[col] = programs[col].astype(str).str.strip().str.upper().map({"TRUE": True, "FALSE": False})
p2019 = programs[programs["year"] == 2019]
result = p2019.merge(c, on="program_key", how="inner")

missing = set(c["program_key"]) - set(result["program_key"])
if missing:
    print(f"WARNING: {len(missing)} profiled programs not found in 2019: {sorted(missing)[:10]}")

result.to_csv("programs_2019_with_profiles.csv", index=False, encoding="utf-8-sig")
print(f"Done: {len(result)} of {len(c)} profiled programs joined "
      f"({int(result['is_gipa'].sum())} GIPA programs) -> programs_2019_with_profiles.csv")
