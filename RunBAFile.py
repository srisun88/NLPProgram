import pandas as pd
import spacy
import re

# Load NLP model
nlp = spacy.load("en_core_web_sm")

# -------------------------------------------------------
# 1. Gherkin Structure Checker (Given/When/Then)
# -------------------------------------------------------
def gherkin_checker(ac_text):
    if not isinstance(ac_text, str) or len(ac_text.strip()) == 0:
        return "Fail", "No AC provided."

    ac = ac_text.lower()
    required = ["given", "when", "then"]

    # Missing keywords
    missing = [word for word in required if word not in ac]
    if missing:
        return "Fail", f"Missing Gherkin keywords: {missing}"

    # Check correct order: Given → When → Then
    given_pos = ac.find("given")
    when_pos  = ac.find("when")
    then_pos  = ac.find("then")

    if not (given_pos < when_pos < then_pos):
        return "Fail", "Incorrect Gherkin order (should be Given → When → Then)."

    return "Pass", "Valid Gherkin format."


# -------------------------------------------------------
# 2. Ambiguity + TBD Severity Checker
# -------------------------------------------------------
AMBIGUOUS_TERMS = [
    "may", "might", "should", "could", "possibly", "approximately",
    "etc", "and/or", "various", "usually", "commonly",
    "tbd", "sometime", "to be decided", "to be determined"
]

def check_ambiguity_with_severity(text):
    text_lower = text.lower()
    doc = nlp(text_lower)
    found = []

    for token in doc:
        if token.text in AMBIGUOUS_TERMS:
            found.append(token.text)

    # Detect multi-word TBD phrases
    if "to be determined" in text_lower:
        found.append("to be determined")
    if "to be decided" in text_lower:
        found.append("to be decided")

    # Severity assignment
    if any(term in text_lower for term in ["tbd", "to be decided", "to be determined"]):
        return "Fail", f"Incomplete data containing: {list(set(found))}", "Critical"

    if found:
        return "Fail", f"Ambiguous words detected: {list(set(found))}", "Medium"

    return "Pass", "No ambiguity detected.", "None"


# -------------------------------------------------------
# 3. Gherkin Rewriter
# -------------------------------------------------------
def rewrite_gherkin(ac_text):
    text = ac_text.strip()

    # If user did not provide any G/W/T keywords
    if not any(k in text.lower() for k in ["given", "when", "then"]):
        return f"Given {text},\nWhen user performs the action,\nThen expected result should occur."

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
# Severity Merger (Combines Gherkin + Ambiguity)
# -------------------------------------------------------
def calculate_final_severity(gherkin_score, gherkin_msg, amb_severity):
    if "Missing Gherkin" in gherkin_msg or "Incorrect Gherkin" in gherkin_msg:
        return "High"

    if amb_severity == "Critical":
        return "Critical"

    if amb_severity == "Medium":
        return "Medium"

    return "None"


# -------------------------------------------------------
# Main Execution
# -------------------------------------------------------
df = pd.read_excel("BA_Acceptance_Criteria.xlsx")
results = []

for index, row in df.iterrows():
    ac = row["Acceptance Criteria"]

    gherkin_score, gherkin_msg = gherkin_checker(ac)
    amb_score, amb_msg, amb_severity = check_ambiguity_with_severity(ac)

    # Final combined severity score
    final_severity = calculate_final_severity(gherkin_score, gherkin_msg, amb_severity)

    rewritten_ac = rewrite_gherkin(ac)

    results.append([
        ac,
        gherkin_score, gherkin_msg,
        amb_score, amb_msg,
        final_severity,
        rewritten_ac
    ])

output_columns = [
    "Original AC",
    "Gherkin Score", "Gherkin Message",
    "Ambiguity Score", "Ambiguity Message",
    "Severity Level",
    "Improved Gherkin AC"
]

output = pd.DataFrame(results, columns=output_columns)
output.to_excel("Gherkin_Analysis_Output.xlsx", index=False)

print("Gherkin-based AC analysis completed successfully with severity levels!")
