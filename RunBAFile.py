import pandas as pd
import spacy
import re

# Load NLP model
nlp = spacy.load("en_core_web_sm")

# -------------------------------------------------------
# 1. Check Structure using NLP (Role, Action, Value)
# -------------------------------------------------------
def check_structure(story):
    doc = nlp(story)

    # Basic pattern check
    pattern = r"As (a|an) .* I want .* so that .*"
    basic = bool(re.search(pattern, story, re.IGNORECASE))

    # Extract role, action, value using dependency parsing
    role = None
    action = None
    benefit = None

    for token in doc:
        if token.dep_ == "nsubj":
            role = token.text
        if token.dep_ == "ROOT":
            action = token.lemma_
        if token.dep_ == "advcl" or token.text.lower() == "so":
            benefit = token.text

    if basic:
        score = "Pass"
        msg = "Follows correct user story pattern."
    else:
        score = "Fail"
        msg = "Does not follow 'As a... I want... so that...' format."

    return score, msg


# -------------------------------------------------------
# 2. Ambiguity Detection using NLP
# -------------------------------------------------------
AMBIGUOUS_TERMS = [
    "may", "might", "should", "could", "possibly", "approximately",
    "etc", "and/or", "various", "usually", "commonly"
]

def check_ambiguity(text):
    doc = nlp(text.lower())
    found = []

    for token in doc:
        if token.text in AMBIGUOUS_TERMS:
            found.append(token.text)

    if found:
        return "Fail", f"Ambiguous words detected: {list(set(found))}"
    else:
        return "Pass", "No ambiguity detected."


# -------------------------------------------------------
# 3. INVEST Analysis using NLP
# -------------------------------------------------------
def invest_analyzer(text):
    doc = nlp(text)

    issues = []

    # I - Independent
    if "and" in text.lower():
        issues.append("Story may contain multiple requirements (check 'Independent').")

    # N - Negotiable
    # Very long stories may not be negotiable
    if len(text.split()) > 40:
        issues.append("Story too long, might not be negotiable.")

    # V - Valuable: check for "so that"
    if "so that" not in text.lower():
        issues.append("Missing 'so that' value clause.")

    # E - Estimable: too vague
    vague = any(word in text.lower() for word in ["sometime", "eventually", "in future"])
    if vague:
        issues.append("Contains vague time expressions (affects Estimable).")

    # S - Small: length check
    if len(text.split()) > 50:
        issues.append("User story too large; break it down (Small).")

    # T - Testable: check AC present
    # Testable validated separately — but we add soft check
    if "verify" not in text.lower() and "validate" not in text.lower():
        issues.append("Story might not be fully testable.")

    if issues:
        return "Fail", "; ".join(issues)
    else:
        return "Pass", "Meets INVEST guidelines."


# -------------------------------------------------------
# 4. Acceptance Criteria Checker (Gherkin)
# -------------------------------------------------------
def ac_checker(ac_text):
    if not isinstance(ac_text, str) or len(ac_text.strip()) == 0:
        return "Fail", "No acceptance criteria provided."

    gherkin_keywords = ["given", "when", "then"]
    ac_lower = ac_text.lower()

    if all(keyword in ac_lower for keyword in gherkin_keywords):
        return "Pass", "Valid Gherkin-style AC."
    else:
        return "Fail", "AC missing Gherkin keywords (Given/When/Then)."


# -------------------------------------------------------
# 5. Rewrite Story using NLP Cleanup
# -------------------------------------------------------
def rewrite(story):
    story = story.strip()

    # Ensure correct format
    if not story.lower().startswith("as a"):
        story = "As a user, " + story[0].lower() + story[1:]

    # Ensure commas
    story = story.replace("I want", ", I want").replace("so that", ", so that")

    # Capitalization cleanup
    story = story[0].upper() + story[1:]

    return story


# -------------------------------------------------------
# Main Execution
# -------------------------------------------------------
df = pd.read_excel("BA_Acceptance_Criteria.xlsx")

results = []

for index, row in df.iterrows():

    story = row["User Story"]
    ac = row["Acceptance Criteria"]

    structure_score, structure_msg = check_structure(story)
    amb_score, amb_msg = check_ambiguity(story)
    invest_score, invest_msg = invest_analyzer(story)
    ac_score, ac_msg = ac_checker(ac)

    rewritten = rewrite(story)

    results.append([
        story,
        structure_score, structure_msg,
        amb_score, amb_msg,
        invest_score, invest_msg,
        ac_score, ac_msg,
        rewritten
    ])


output_columns = [
    "Story", "Structure Score", "Structure Message",
    "Ambiguity Score", "Ambiguity Message",
    "INVEST Score", "INVEST Message",
    "AC Score", "AC Message",
    "Improved Story"
]

output = pd.DataFrame(results, columns=output_columns)
output.to_excel("UserStoryAnalysis_NLP.xlsx", index=False)

print("NLP-based user story analysis completed successfully!")