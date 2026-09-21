from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES
from app_core import validate_gorus_analysis

SPEC = "Claim language here. Sensor A is positioned in the main body. Sensor B is positioned in an external attachment."

def base_analysis():
    return {
        "amendment_required": True,
        "amendments": [{
            "claim_number": "1",
            "reason": "distinguishing limitation",
            "basis_quote": "Sensor A is positioned in the main body.",
            "change_type": "technical",
            "direct_support": [
                {"added_feature": "Sensor A is positioned in the main body", "basis_quote": "Sensor A is positioned in the main body."},
                {"added_feature": "Sensor B is positioned in an external attachment", "basis_quote": "Sensor B is positioned in an external attachment."},
            ],
            "old_text": "Claim language here.",
            "new_text": "Claim language here. Sensor A is positioned in the main body and Sensor B is positioned in an external attachment.",
        }],
        "technical_contributions": [],
        "customer_defence_points": [],
        "description_prior_art_updates": [],
    }

def test_version_and_direct_support_rule():
    assert APP_VERSION == "v5.4.72"
    assert RULESET_VERSION == "2026-09-18.v64"
    low = GORUS_RULES.casefold()
    assert "dolaylı" in low and "direct_support" in low and "doğrudan" in low

def test_direct_support_can_come_from_multiple_paragraphs():
    validate_gorus_analysis(base_analysis(), SPEC)

def test_technical_amendment_without_direct_support_fails():
    a=base_analysis(); a["amendments"][0]["direct_support"]=[]
    try:
        validate_gorus_analysis(a, SPEC)
    except ValueError as e:
        assert "direct_support" in str(e)
    else:
        raise AssertionError("expected fail")

def test_nonexistent_direct_basis_fails():
    a=base_analysis(); a["amendments"][0]["direct_support"][1]["basis_quote"]="implied but not written"
    try:
        validate_gorus_analysis(a, SPEC)
    except ValueError as e:
        assert "Doğrudan revizyon dayanağı" in str(e)
    else:
        raise AssertionError("expected fail")
