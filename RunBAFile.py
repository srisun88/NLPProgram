import pandas as pd
import spacy
import re

nlp = spacy.load("en_core_web_sm")

# -----------------------------------------
# GLOBAL CONFIG
# -----------------------------------------
TBD_TERMS = ["tbd", "to be decided", "to be determined"]

AMBIGUOUS_TERMS = [
    "may", "might", "should", "could", "possibly", "approximately",
    "etc", "and/or", "various", "usually", "commonly", "sometime"
] + TBD_TERMS

NFR_KEYWORDS = {
    "performance": ["seconds", "load", "response", "latency"],
    "security": ["encrypt", "secure", "authentication"],
    "compatibility": ["chrome", "firefox", "mobile", "browser"],
    "error_handling": ["error", "invalid", "fail"],
    "availability": ["99%", "uptime", "sla"]
}

# -----------------------------------------
# GHERKIN REWRITE (NEW LOGIC)
# -----------------------------------------
def gherkin_rewrite(ac_text):
    ac = ac_text.strip().lower()

    # Full rewrite
    if not any(k in ac for k in ["given", "when", "then"]):
        return (
            f"Given a valid user session, "
            f"When {ac_text.split()[0].lower()} action is triggered, "
            f"Then {ac_text}"
        )

    # Partial rewrite: missing WHEN
    if ("given" in ac) and ("when" not in ac) and ("then" in ac):
        return ac_text.replace("Given", "Given") + "\nWhen action is triggered"

    # Partial rewrite: missing THEN
    if ("given" in ac) and ("when" in ac) and ("then" not in ac):
        return ac_text.replace("When", "When") + "\nThen expected outcome"

    # Already Gherkin → return original
    return ac_text


# -----------------------------------------
# GHERKIN CHECKER
# -----------------------------------------
def gherkin_checker(ac_text):
    if not isinstance(ac_text, str) or len(ac_text.strip()) == 0:
        return "Fail", "No AC provided."

    ac = ac_text.lower()

    if any(term in ac for term in TBD_TERMS):
        return "Fail", "AC contains TBD."

    required = ["given", "when", "then"]
    missing = [word for word in required if word not in ac]

    if missing:
        return "Fail", f"Missing Gherkin elements: {missing}"

    if not (ac.find("given") < ac.find("when") < ac.find("then")):
        return "Fail", "Incorrect GWT order"

    return "Pass", "Valid Gherkin format"


# -----------------------------------------
# AMBIGUITY CHECKER
# -----------------------------------------
def check_ambiguity(text):
    if not isinstance(text, str):
        return "Fail", "No AC", "Critical"

    t = text.lower()
    doc = nlp(t)
    found = set()

    for token in doc:
        if token.text in AMBIGUOUS_TERMS:
            found.add(token.text)

    if found:
        if any(term in t for term in TBD_TERMS):
            return "Fail", f"TBD: {found}", "Critical"
        return "Fail", f"Ambiguous: {found}", "Medium"

    return "Pass", "Clear", "None"


# -----------------------------------------
# INVEST
# -----------------------------------------
def invest_scoring(text):
    if not isinstance(text, str):
        return {}, "Invalid input"

    t = text.lower()
    score = {
        "I_Independent": "Pass",
        "N_Negotiable": "Pass",
        "V_Valuable": "Pass",
        "E_Estimable": "Pass",
        "S_Small": "Pass",
        "T_Testable": "Pass"
    }

    if re.search(r"\b(story|depends|subtask|part\s*\d)\b", t):
        score["I_Independent"] = "Fail"

    ui_terms = ["click button", "use react", "database table", "html"]
    if any(term in t for term in ui_terms):
        score["N_Negotiable"] = "Fail"

    if "so that" not in t:
        score["V_Valuable"] = "Fail"

    if not re.search(r"\d+", t):
        score["E_Estimable"] = "Fail"

    if len(text.split()) > 80 or text.lower().count(" and ") > 2:
        score["S_Small"] = "Fail"

    if not any(k in t for k in ["given", "then", "expected", "result"]):
        score["T_Testable"] = "Fail"

    total_fails = list(v for v in score.values() if v == "Fail")
    summary = "Pass" if len(total_fails) == 0 else f"{len(total_fails)} INVEST weaknesses"

    return score, summary


# -----------------------------------------
# NFR CHECK
# -----------------------------------------
def nfr_check(text):
    if not isinstance(text, str):
        return "Fail", "No AC"

    found = []
    for category, words in NFR_KEYWORDS.items():
        if any(w in text.lower() for w in words):
            found.append(category)

    if found:
        return "Pass", f"NFR: {list(set(found))}"

    return "Fail", "No NFR"


# -----------------------------------------
# AUTOMATION MAPPING
# -----------------------------------------
def automation_mapping(text):
    if not isinstance(text, str):
        return "Low", "Not automatable"

    t = text.lower()
    score = 0

    if all(k in t for k in ["given", "when", "then"]):
        score += 2
    if re.search(r"\d+", t):
        score += 2
    if "error" in t or "invalid" in t:
        score += 1

    if score >= 4:
        return "High", "Strong BDD candidate"
    if score >= 2:
        return "Medium", "Partial automation"

    return "Low", "Weak automation"


# -----------------------------------------
# MAIN EXECUTION
# -----------------------------------------
df = pd.read_excel("BA_Acceptance_Criteria.xlsx")
results = []

for _, row in df.iterrows():
    ac = row["Acceptance Criteria"]

    g_score, g_msg = gherkin_checker(ac)
    a_score, a_msg, _ = check_ambiguity(ac)
    invest_dict, invest_msg = invest_scoring(ac)
    nfr_score, nfr_msg = nfr_check(ac)
    auto_score, auto_msg = automation_mapping(ac)

    # 🔥 New rewritten Gherkin output
    improved_story = gherkin_rewrite(ac)

    results.append([
        ac,
        g_score, g_msg,
        a_score, a_msg,
        invest_dict["I_Independent"],
        invest_dict["N_Negotiable"],
        invest_dict["V_Valuable"],
        invest_dict["E_Estimable"],
        invest_dict["S_Small"],
        invest_dict["T_Testable"],
        invest_msg,
        nfr_score, nfr_msg,
        auto_score, auto_msg,
        improved_story
    ])

output_cols = [
    "Original AC",
    "Gherkin Score", "Gherkin Message",
    "Ambiguity Score", "Ambiguity Message",
    "INVEST I", "INVEST N", "INVEST V", "INVEST E", "INVEST S", "INVEST T", "INVEST Summary",
    "NFR Score", "NFR Message",
    "Automation Score", "Automation Message",
    "Improved Story"
]

out = pd.DataFrame(results, columns=output_cols)
out.to_excel("Agile_Quality_INVEST.xlsx", index=False)

print("✔ Completed with Gherkin rewrite applied.")