import pandas as pd

programs = pd.read_csv("programs_all_years.csv", dtype={"program_key": str, "hei_code": str})

# התחומים שהבנות בחרו (מילות מפתח, כדי לתפוס גם וריאציות של השם)
RELEVANT = ["computer science", "sociolog", "political science", "international relations",
            "journalism", "mass communication", "economic", "business administration",
            "law", "jurisprudence", "psycholog", "digital telecommunication", "public relations",
            "management", "financ", "visual arts", "audio-visual", "audiovisual",
            "visual communication", "digital marketing", "digital media"]

# מוציאים גם אם יש מילה רלוונטית (למשל "Construction Management", "Hotel Management")
EXCEPT = ["engineer", "transport", "logistic", "freight", "construct", "agri", "agro", "farming",
          "aviation", "air ", "maritime", "port ", "food", "industrial", "sport", "health",
          "music", "forest", "tourism", "hotel", "hospitality", "quality"]

field = programs["program_field_en"].fillna("").str.lower()
programs["is_relevant"] = (
    (field.str.contains("|".join(RELEVANT)) & ~field.str.contains("|".join(EXCEPT)))
    | programs["is_gipa"]          # תוכניות של GIPA תמיד נשארות
)

relevant = programs[programs["is_relevant"]]
relevant.to_csv("programs_relevant.csv", index=False, encoding="utf-8-sig")
print(f"Kept {len(relevant):,} of {len(programs):,} programs")