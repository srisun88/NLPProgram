import pandas as pd
import spacy
import re

# Load NLP model
nlp = spacy.load("en_core_web_sm")

# -------------------------------------------------------
# GLOBAL LIST (Ambiguity + TBD)
# -------------------------------------------------------
TBD_TERMS = ["tbd", "to be decided", "to be determined"]
AMBIGUOUS_TERMS = [
    "may", "might", "should", "could", "possibly", "approximately",
    "etc", "and/or", "various", "usually", "commonly",
    "sometime"
] + TBD_TERMS


# -------------------------------------------------------
# 1. Gherkin Structure Checker (Now includes TBD FAIL)
# -------------------------------------------------------
def gherkin_checker(ac_text):
    if not isinstance(ac_text, str) or len(ac_text.strip()) == 0:
        return "Fail", "No AC provided."

    ac = ac_text.lower()

    # NEW: Immediate fail for incomplete (TBD)
    if any(term in ac for term in TBD_TERMS):
        return "Fail", "AC contains TBD / incomplete information."

    required = ["given", "when", "then"]

    missing = [word for word in required if word not in ac]
    if missing:
        return "Fail", f"Missing Gherkin keywords: {missing}"

    # Check order
    if not (ac.find("given") < ac.find("when") < ac.find("then")):
        return "Fail", "Incorrect Gherkin order (should be Given → When → Then)."

    return "Pass", "Valid Gherkin format."


# -------------------------------------------------------
# 2. Ambiguity + Severity Checker (TBD → Critical)
# -------------------------------------------------------
def check_ambiguity_with_severity(text):
    if not isinstance(text, str) or len(text.strip()) == 0:
        return "Fail", "No AC provided.", "Critical"

    text_lower = text.lower()
    doc = nlp(text_lower)

    found = []

    # Single-word token-based detection
    for token in doc:
        if token.text in AMBIGUOUS_TERMS:
            found.append(token.text)

    # Multi-word TBD detection
    for phrase in TBD_TERMS:
        if phrase in text_lower:
            found.append(phrase)

    found = list(set(found))  # unique terms

    # TBD → Critical severity
    if any(term in text_lower for term in TBD_TERMS):
        return "Fail", f"Incomplete data (TBD) detected: {found}", "Critical"

    # Normal ambiguity
    if found:
        return "Fail", f"Ambiguous words detected: {found}", "Medium"

    return "Pass", "No ambiguity detected.", "None"


# -------------------------------------------------------
# 3. Gherkin Rewriter
# -------------------------------------------------------
def rewrite_gherkin(ac_text):
    text = ac_text.strip()

    # Skip rewrite if TBD present
    if any(term in text.lower() for term in TBD_TERMS):
        return "Cannot rewrite — AC contains TBD/incomplete information."

    if not any(k in text.lower() for k in ["given", "when", "then"]):
        return f"Given {text},\nWhen user performs the action,\nThen the expected result should occur."

    lines = text.split("\n")
    formatted = []

    for line in lines:
        original = line.strip()
        lower = original.lower()

        if lower.startswith("given"):
            formatted.append("Given " + original[5:].strip().capitalize())
        elif lower.startswith("when"):
            formatted.append("When " + original[4:].strip().capitalize())
        elif lower.startswith("then"):
            formatted.append("Then " + original[4:].strip().capitalize())
        else:
            formatted.append(original.capitalize())

    return "\n".join(formatted)


# -------------------------------------------------------
# 4. Final Severity Merger
# -------------------------------------------------------
def calculate_final_severity(gherkin_msg, amb_severity):
    if "TBD" in gherkin_msg or "incomplete" in gherkin_msg.lower():
        return "Critical"

    if amb_severity == "Critical":
        return "Critical"

    if "Missing Gherkin" in gherkin_msg or "Incorrect Gherkin" in gherkin_msg:
        return "High"

    if amb_severity == "Medium":
        return "Medium"

    return "None"


# -------------------------------------------------------
# 5. MAIN EXECUTION
# -------------------------------------------------------
df = pd.read_excel("BA_Acceptance_Criteria.xlsx")
results = []

for index, row in df.iterrows():
    ac = row["Acceptance Criteria"]

    # Checks
    gherkin_score, gherkin_msg = gherkin_checker(ac)
    amb_score, amb_msg, amb_severity = check_ambiguity_with_severity(ac)

    # Final severity combining both
    final_severity = calculate_final_severity(gherkin_msg, amb_severity)

    # Improved gherkin rewrite
    rewritten_ac = rewrite_gherkin(ac)

    # Append results
    results.append([
        ac,
        gherkin_score, gherkin_msg,
        amb_score, amb_msg,
        final_severity,
        rewritten_ac
    ])

# Output column names
output_columns = [
    "Original AC",
    "Gherkin Score", "Gherkin Message",
    "Ambiguity Score", "Ambiguity Message",
    "Severity Level",
    "Improved Gherkin AC"
]

# Export to Excel
output = pd.DataFrame(results, columns=output_columns)
output.to_excel("Gherkin_Analysis_Output.xlsx", index=False)

print("✔ Gherkin AC analysis completed successfully with updated TBD handling!")
