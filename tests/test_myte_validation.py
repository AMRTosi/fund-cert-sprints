from __future__ import annotations

import pytest

from sprint_cert_automation.domain.myte_validation import (
    EXCEL_REQUIRED_FIELDS,
    MYTE_REQUIRED_FIELDS,
    validate_payloads,
)


@pytest.fixture
def excel() -> dict[str, str]:
    return {
        "reviewee": "alice", "email": "alice@example.com", "subco": "0",
        "horasQuincena1": "80", "horasQuincena2": "72",
    }


@pytest.fixture
def myte() -> dict[str, str]:
    return {
        "reviewee": "alice", "trStatus": "Processed", "chargeCode": "X(COMK3001)",
        "totalHours": "66.4", "workingHours": "80",
        "percentageTotalHoursVsWorkingHours": "83%",
    }


def test_exact_passed_contract_and_identity(excel, myte) -> None:
    excel["reviewee"] = " Alice "
    excel["email"] = " Original@example.com "
    assert validate_payloads([excel], [myte], 1) == [{
        "reviewee": " Alice ", "email": " Original@example.com ",
        "validationStatus": "Passed", "sendNotification": False,
        "horasQuincena1": "80", "horasQuincena2": "72",
        "trStatus": "Processed", "chargeCode": "X(COMK3001)",
        "totalHours": "66.4", "workingHours": "80",
        "percentageTotalHoursVsWorkingHours": "83%",
        "comments": "", "notificationMessage": "",
    }]


@pytest.mark.parametrize("subco,expected", [("1", []), (" 1 ", []), ("0", ["Passed"]), ("SC01", ["Passed"])])
def test_subco_exclusion(excel, myte, subco, expected) -> None:
    excel["subco"] = subco
    assert [r["validationStatus"] for r in validate_payloads([excel], [myte], 1)] == expected


def test_missing_myte_and_extra_myte_keep_excel_order(excel, myte) -> None:
    rows = [dict(excel, reviewee="missing"), dict(excel, reviewee="ignored", subco="1"), excel]
    results = validate_payloads(rows, [dict(myte, reviewee="extra"), myte], 1)
    assert [r["reviewee"] for r in results] == ["missing", "alice"]
    assert results[0] == {
        "reviewee": "missing", "email": "alice@example.com", "validationStatus": "Failed",
        "horasQuincena1": "80", "horasQuincena2": "72",
        "trStatus": "", "chargeCode": "", "totalHours": "", "workingHours": "",
        "percentageTotalHoursVsWorkingHours": "",
        "sendNotification": False,
        "comments": "El reviewee registrado en el fichero excel no ha sido encontrado en myTE",
        "notificationMessage": "",
    }


@pytest.mark.parametrize("period", [1, 2])
@pytest.mark.parametrize("status,working,percentage,passed,notify,comment", [
    (" Processed ", " 80,04 ", " 84,00% ", True, False, ""),
    (" New ", "invalid", "invalid", False, True, "myTE no enviado. TR Status es New o Draft."),
    (" Draft ", "invalid", "invalid", False, True, "myTE no enviado. TR Status es New o Draft."),
    (" Pending ", "invalid", "invalid", False, False, "Estado diferente de Processed. Revisar"),
    (
        " Processed ", " 79,00 ", "invalid", False, False,
        "El numero de WorkingHours registradas en myTE es diferente al previsto en Excel.",
    ),
    (
        " Processed ", " 80,04 ", " 87,00% ", False, False,
        "El porcentaje < 87,00% > no cumple la politica definida para el chargeCode.",
    ),
])
def test_all_branches_preserve_exact_source_strings(
    excel, myte, period, status, working, percentage, passed, notify, comment,
) -> None:
    excel.update(reviewee=" Alice ", email=" Original@example.com ", extra="not exported")
    excel["horasQuincena1"] = "unparsed" if period == 2 else " 80,00 "
    excel["horasQuincena2"] = "unparsed" if period == 1 else " 80,00 "
    myte.update(
        reviewee=" ALICE ", trStatus=status, chargeCode=" FUNDAE (COMK3001) ",
        totalHours=" not numeric ", workingHours=working,
        percentageTotalHoursVsWorkingHours=percentage, extra="not exported",
    )
    other_myte = dict(myte, reviewee="bob", chargeCode="COMK3002", totalHours="different")
    assert validate_payloads([excel], [other_myte, myte], period) == [{
        "reviewee": " Alice ",
        "email": " Original@example.com ",
        "validationStatus": "Passed" if passed else "Failed",
        "horasQuincena1": excel["horasQuincena1"],
        "horasQuincena2": excel["horasQuincena2"],
        "trStatus": status,
        "chargeCode": " FUNDAE (COMK3001) ",
        "totalHours": " not numeric ",
        "workingHours": working,
        "percentageTotalHoursVsWorkingHours": percentage,
        "comments": comment,
        "sendNotification": notify,
        "notificationMessage": "",
    }]


@pytest.mark.parametrize("status,notify,comment", [
    ("New", True, "myTE no enviado. TR Status es New o Draft."),
    (" Draft ", True, "myTE no enviado. TR Status es New o Draft."),
    ("Pending", False, "Estado diferente de Processed. Revisar"),
    ("processed", False, "Estado diferente de Processed. Revisar"),
])
def test_status_short_circuits_numeric_validation(excel, myte, status, notify, comment) -> None:
    excel["horasQuincena1"] = "invalid"
    myte.update(trStatus=status, workingHours="invalid", percentageTotalHoursVsWorkingHours="invalid")
    result = validate_payloads([excel], [myte], 1)[0]
    assert result["validationStatus"] == "Failed"
    assert result["sendNotification"] is notify
    assert result["comments"] == comment
    assert result["notificationMessage"] == ""


@pytest.mark.parametrize("period,hours,working,passed", [
    (1, "80", "80.04", True), (1, "80", "80.06", False),
    (1, "94,0", "94.0", True), (2, "72", "72", True),
    (2, "72", "80", False), (1, "80.25", "80.2", True),
    (1, "80.35", "80.3", True),
])
def test_hours_and_original_float_rounding(excel, myte, period, hours, working, passed) -> None:
    excel["horasQuincena1" if period == 1 else "horasQuincena2"] = hours
    myte["workingHours"] = working
    result = validate_payloads([excel], [myte], period)[0]
    assert result["validationStatus"] == ("Passed" if passed else "Failed")
    if not passed:
        assert result["comments"] == "El numero de WorkingHours registradas en myTE es diferente al previsto en Excel."
        assert result["sendNotification"] is False


def test_hours_mismatch_does_not_parse_percentage(excel, myte) -> None:
    myte.update(workingHours="79", percentageTotalHoursVsWorkingHours="invalid")
    assert validate_payloads([excel], [myte], 1)[0]["validationStatus"] == "Failed"


@pytest.mark.parametrize("code,percentage,passed", [
    ("COMK3001", "81.9%", False), ("COMK3001", "82%", True),
    ("COMK3001", "86.0", True), ("COMK3001", "86.1%", False),
    ("COMK3002", "90.9", False), ("COMK3002", "91%", True),
    ("COMK3002", "93%", True), ("COMK3002", "93.1%", False),
    ("COMK3002", "92,0", True), ("UNKNOWN", "83%", False),
    ("comk3001", "83%", False), ("COMK3001 COMK3002", "92%", True),
])
def test_percentage_boundaries_and_code_precedence(excel, myte, code, percentage, passed) -> None:
    myte.update(chargeCode=code, percentageTotalHoursVsWorkingHours=percentage)
    result = validate_payloads([excel], [myte], 1)[0]
    assert result["validationStatus"] == ("Passed" if passed else "Failed")
    assert result["comments"] == (
        "" if passed else f"El porcentaje <{percentage}> no cumple la politica definida para el chargeCode."
    )
    assert result["sendNotification"] is False


@pytest.mark.parametrize("reviewee", ["patricia.c.martin", " E.A.MAGANA.RAMIREZ "])
@pytest.mark.parametrize("percentage,passed", [("79.9%", False), ("80,0%", True), ("80.1%", False)])
def test_special_reviewees(excel, myte, reviewee, percentage, passed) -> None:
    excel["reviewee"] = reviewee
    myte.update(reviewee=reviewee.strip().lower(), percentageTotalHoursVsWorkingHours=percentage)
    assert validate_payloads([excel], [myte], 1)[0]["validationStatus"] == ("Passed" if passed else "Failed")


@pytest.mark.parametrize("source,fields", [("excel", EXCEL_REQUIRED_FIELDS), ("myte", MYTE_REQUIRED_FIELDS)])
@pytest.mark.parametrize("bad_type", [False, True])
def test_all_required_fields_abort_even_excluded_records(excel, myte, source, fields, bad_type) -> None:
    excel["subco"] = "1"
    for field in fields:
        excel_row, myte_row = excel.copy(), myte.copy()
        row = excel_row if source == "excel" else myte_row
        if bad_type:
            row[field] = 1
        else:
            del row[field]
        with pytest.raises(ValueError, match=f"'{field}'"):
            validate_payloads([excel_row], [myte_row], 1)


@pytest.mark.parametrize("source", ["excel", "myte"])
def test_duplicates_abort_even_excluded_or_unmatched(excel, myte, source) -> None:
    excel["subco"] = "1"
    a, b = [excel], [myte]
    if source == "excel":
        a.append(dict(excel, reviewee=" ALICE "))
    else:
        b = [dict(myte, reviewee="extra"), dict(myte, reviewee=" EXTRA ")]
    with pytest.raises(ValueError, match=f"Reviewee duplicado en {source}_extracted_filePath"):
        validate_payloads(a, b, 1)


@pytest.mark.parametrize("a,b", [({}, []), ([], {}), ([1], []), ([], [None])])
def test_invalid_root_or_item(a, b) -> None:
    with pytest.raises(ValueError, match="debe ser"):
        validate_payloads(a, b, 1)


@pytest.mark.parametrize("field", ["horasQuincena1", "workingHours", "percentageTotalHoursVsWorkingHours"])
def test_invalid_numeric_value_aborts(excel, myte, field) -> None:
    (excel if field == "horasQuincena1" else myte)[field] = "invalid"
    with pytest.raises(ValueError, match=f"Valor numerico invalido en '{field}'"):
        validate_payloads([excel], [myte], 1)


def test_total_hours_unused_and_empty_identity_preserved(excel, myte) -> None:
    excel.update(reviewee="", email="")
    myte.update(reviewee="", totalHours="not numeric")
    assert validate_payloads([excel], [myte], 1)[0]["validationStatus"] == "Passed"
    assert validate_payloads([], [], 2) == []


def test_unmatched_myte_contract_still_aborts(excel, myte) -> None:
    myte["reviewee"] = "extra"
    del myte["workingHours"]
    with pytest.raises(ValueError, match="'workingHours'"):
        validate_payloads([excel], [myte], 1)


def test_nonselected_hours_not_parsed_and_processed_trimmed(excel, myte) -> None:
    excel["horasQuincena2"] = "not numeric"
    myte["trStatus"] = " Processed "
    assert validate_payloads([excel], [myte], 1)[0]["validationStatus"] == "Passed"


@pytest.mark.parametrize("hours,percentage,passed", [
    ("nan", "84%", False), ("inf", "84%", True),
    ("-80", "84%", True), ("80", "nan", False), ("80", "inf", False),
])
def test_original_special_float_values(excel, myte, hours, percentage, passed) -> None:
    excel["horasQuincena1"] = hours
    myte.update(workingHours=hours, percentageTotalHoursVsWorkingHours=percentage)
    assert validate_payloads([excel], [myte], 1)[0]["validationStatus"] == ("Passed" if passed else "Failed")


@pytest.mark.parametrize("period", [0, 3, -1, True, "1"])
def test_invalid_period_rejected_at_domain_boundary(period) -> None:
    with pytest.raises(ValueError, match="period-number"):
        validate_payloads([], [], period)
