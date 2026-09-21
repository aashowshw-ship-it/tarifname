from pathlib import Path
import pytest
import app_core as app


def test_reference_display_order_numeric_then_symbolic_then_method_is_separate_block():
    rows = [
        {"number":"K","name":"Arayan"},
        {"number":"160","name":"Sınıf koşullu modül"},
        {"number":"62","name":"Ağırlıklı dal"},
        {"number":"60","name":"Havuzlama başlığı"},
        {"number":"80b","name":"Karar birleştirici"},
        {"number":"80a","name":"Yayın birleştirici"},
        {"number":"O","name":"Operatör"},
        {"number":"61","name":"Dikkat dalı"},
    ]
    ordered = app._ordered_reference_elements(rows)
    assert [r["number"] for r in ordered] == ["60","61","62","80a","80b","160","K","O"]


def test_generic_system_whole_cannot_be_numbered_element():
    draft = {"elements":[{"number":"1","name":"Sistem"}], "system_claim":{}, "dependent_system_claims":[]}
    with pytest.raises(ValueError, match="buluşun bütünü"):
        app._validate_reference_role_rules(draft, "Türkçe")


def test_human_actor_reference_cannot_be_claim_element_but_terminal_can():
    bad = {
        "elements":[{"number":"O","name":"Operatör"}],
        "system_claim":{"preamble":"Teknik bir sistem", "elements":["operatör (O),"], "closing":"içermesidir."},
        "dependent_system_claims":[],
    }
    with pytest.raises(ValueError, match="insan rolü"):
        app._validate_reference_role_rules(bad, "Türkçe")

    good = {
        "elements":[{"number":"10","name":"Operatör terminali"}],
        "system_claim":{"preamble":"Teknik bir sistem", "elements":["operatör terminali (10),"], "closing":"içermesidir."},
        "dependent_system_claims":[],
    }
    app._validate_reference_role_rules(good, "Türkçe")


def test_dependent_claim_type_closure_rejects_generic_structure_and_mismatch():
    draft = {"elements":[{"number":"20","name":"Kanal çözme modülü"}]}
    with pytest.raises(ValueError, match="bir yapıda olmasıdır"):
        app._validate_dependent_system_claim_type_closure([
            "İstem 1’e uygun sistem olup, özelliği; kanal çözme modülünün (20), dar bant veriyi işleyen bir yapıda olmasıdır."
        ], draft, "Türkçe")
    with pytest.raises(ValueError, match="bir modül olmasıdır"):
        app._validate_dependent_system_claim_type_closure([
            "İstem 1’e uygun sistem olup, özelliği; kanal çözme modülünün (20), dar bant veriyi işleyen bir birim olmasıdır."
        ], draft, "Türkçe")
    app._validate_dependent_system_claim_type_closure([
        "İstem 1’e uygun sistem olup, özelliği; kanal çözme modülünün (20), dar bant veriyi işleyen bir modül olmasıdır."
    ], draft, "Türkçe")


def test_additional_independent_method_claims_are_exposed_to_validators():
    draft = {
        "method_claim":{"preamble":"Birinci yöntem için yeterince uzun teknik bir giriş bağlamı sağlayan yöntem", "steps":["birinci işlemin yapılması (1001)"]},
        "dependent_method_claims":["İstem 1’e uygun yöntem olup, özelliği; ek işlem adımını içermesidir."],
        "additional_method_claims":[{
            "claim":{"preamble":"İkinci bağımsız yöntem için yeterince uzun teknik bir giriş bağlamı sağlayan yöntem", "steps":["ikinci işlemin yapılması (1030)"], "closing":"işlem adımlarını içermesidir."},
            "dependent_claims":[]
        }]
    }
    assert len(app._all_method_claim_objects(draft)) == 2
    assert len(app._all_method_dependent_claims(draft)) == 1


def test_text_heavy_source_figure_is_not_automatically_excluded_anymore():
    audit = {
        "final_use":"include",
        "text_heavy_nonreference_numbering":True,
        "nonreference_numeric_or_step_marks":["A1","B2"],
    }
    assert app._figure_audit_requests_exclusion(audit) is False


def test_figure_flow_receives_extra_instruction_and_geometry_fail_closed_fields():
    source = Path(app.__file__).read_text(encoding="utf-8")
    assert "extra_instruction=extra_instruction" in source
    assert 'bool(audit.get("text_cleanup_required"))' in source
    for token in ["arrows_preserved", "box_edges_preserved", "connections_preserved", "only_requested_text_changed"]:
        assert token in source
    assert '"technical_term_clarity_passed"' in source


def test_symbolic_context_refs_are_not_forced_into_system_claims():
    draft = {
        "elements":[
            {"number":"10","name":"Medya çoğaltma birimi"},
            {"number":"K","name":"Arayan"},
            {"number":"O","name":"Operatör"},
            {"number":"R","name":"Ses kaydı"},
        ],
        "system_claim":{
            "preamble":"Telefon kanalı üzerinde sentetik ses tespiti gerçekleştiren bir sistem",
            "elements":["medya çoğaltma birimi (10),"],
            "closing":"içermesidir.",
        },
        "dependent_system_claims":[],
    }
    app._validate_all_elements_covered_in_claims(draft)


def test_numeric_technical_ref_still_must_be_covered_in_system_claims():
    draft = {
        "elements":[
            {"number":"10","name":"Medya çoğaltma birimi"},
            {"number":"20","name":"Kanal çözme modülü"},
        ],
        "system_claim":{
            "preamble":"Telefon kanalı üzerinde sentetik ses tespiti gerçekleştiren bir sistem",
            "elements":["medya çoğaltma birimi (10),"],
            "closing":"içermesidir.",
        },
        "dependent_system_claims":[],
    }
    with pytest.raises(ValueError, match="Kanal çözme modülü"):
        app._validate_all_elements_covered_in_claims(draft)


def test_explicit_source_figure_exclusion_requires_new_fail_closed_reason():
    audit = {
        "final_use":"exclude_only_if_nontechnical_or_unrecoverable",
        "exclusion_reason":"yalnız dekoratif logo olup teknik akış/geometri içermiyor",
        "unique_technical_information":[],
        "spec_coverage_evidence":[],
    }
    assert app._figure_audit_requests_exclusion(audit) is True
    assert app._validate_source_figure_exclusion(audit, {}) is None
    bad = dict(audit, exclusion_reason="kötü")
    assert app._validate_source_figure_exclusion(bad, {}) is not None


def test_figure_audit_prompt_receives_exact_extra_instruction():
    prompt = app._figure_reference_audit_prompt(
        {"elements":[], "method_steps":[]},
        1,
        "Türkçe",
        "Çizimler (algoritmalar) üzerinde yazı olmayacak şekilde düzenlenmeli",
    )
    assert "Çizimler (algoritmalar) üzerinde yazı olmayacak şekilde düzenlenmeli" in prompt
