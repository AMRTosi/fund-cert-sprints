from __future__ import annotations

from typing import Literal, TypedDict


EXCEL_REQUIRED_FIELDS = ("reviewee", "email", "subco", "horasQuincena1", "horasQuincena2")
MYTE_REQUIRED_FIELDS = (
    "reviewee", "trStatus", "chargeCode", "totalHours", "workingHours",
    "percentageTotalHoursVsWorkingHours",
)
SPECIAL_80_PERCENT_REVIEWEES = {"patricia.c.martin", "e.a.magana.ramirez"}


class MyteValidationResult(TypedDict):
    reviewee: str
    email: str
    validationStatus: Literal["Passed", "Failed"]
    horasQuincena1: str
    horasQuincena2: str
    trStatus: str
    chargeCode: str
    totalHours: str
    workingHours: str
    percentageTotalHoursVsWorkingHours: str
    comments: str
    sendNotification: bool
    notificationMessage: str


def validate_period_number(period_number: int) -> None:
    if type(period_number) is not int or period_number not in (1, 2):
        raise ValueError("period-number must be 1 or 2")


def _assert_array_payload(payload: object, name: str) -> list[dict[str, object]]:
    if not isinstance(payload, list):
        raise ValueError(f"{name} debe ser un array de objetos JSON")
    items: list[dict[str, object]] = []
    for index, item in enumerate(payload):
        if not isinstance(item, dict):
            raise ValueError(f"{name}[{index}] debe ser un objeto JSON")
        items.append(item)
    return items


def _validate_contracts(
    items: list[dict[str, object]], fields: tuple[str, ...], name: str,
) -> list[dict[str, str]]:
    validated: list[dict[str, str]] = []
    for index, item in enumerate(items):
        errors: list[str] = []
        record: dict[str, str] = {}
        for field in fields:
            if field not in item:
                errors.append(f"Falta el campo obligatorio '{field}'")
            else:
                value = item[field]
                if not isinstance(value, str):
                    errors.append(f"El campo '{field}' debe ser de tipo str")
                else:
                    record[field] = value
        if errors:
            raise ValueError(f"{name}[{index}]: {' | '.join(errors)}")
        validated.append(record)
    return validated


def _index_by_reviewee(items: list[dict[str, str]], name: str) -> dict[str, dict[str, str]]:
    indexed: dict[str, dict[str, str]] = {}
    for item in items:
        raw_reviewee = item["reviewee"]
        key = raw_reviewee.strip().lower()
        if key in indexed:
            raise ValueError(f"Reviewee duplicado en {name}: '{raw_reviewee}'")
        indexed[key] = item
    return indexed


def _parse_decimal(value: str, field_name: str, reviewee: str) -> float:
    try:
        return float(value.strip().replace("%", "").replace(",", "."))
    except ValueError as exc:
        raise ValueError(
            f"Valor numerico invalido en '{field_name}' para reviewee '{reviewee}': {value}"
        ) from exc


def _result(
    excel: dict[str, str], myte: dict[str, str] | None,
    comments: str = "", send_notification: bool = False,
) -> MyteValidationResult:
    return {
        "reviewee": excel["reviewee"],
        "email": excel["email"],
        "validationStatus": "Failed" if comments else "Passed",
        "horasQuincena1": excel["horasQuincena1"],
        "horasQuincena2": excel["horasQuincena2"],
        "trStatus": myte["trStatus"] if myte is not None else "",
        "chargeCode": myte["chargeCode"] if myte is not None else "",
        "totalHours": myte["totalHours"] if myte is not None else "",
        "workingHours": myte["workingHours"] if myte is not None else "",
        "percentageTotalHoursVsWorkingHours": (
            myte["percentageTotalHoursVsWorkingHours"] if myte is not None else ""
        ),
        "comments": comments,
        "sendNotification": send_notification,
        "notificationMessage": "",
    }


def validate_payloads(
    payload_a: object, payload_b: object, period_number: int,
) -> list[MyteValidationResult]:
    validate_period_number(period_number)
    excel_name = "excel_extracted_filePath"
    myte_name = "myte_extracted_filePath"
    excel_items = _assert_array_payload(payload_a, excel_name)
    myte_items = _assert_array_payload(payload_b, myte_name)
    excel_records = _validate_contracts(excel_items, EXCEL_REQUIRED_FIELDS, excel_name)
    myte_records = _validate_contracts(myte_items, MYTE_REQUIRED_FIELDS, myte_name)
    # Validate all contracts and duplicates before exclusions, as in the original service.
    excel_by_reviewee = _index_by_reviewee(excel_records, excel_name)
    myte_by_reviewee = _index_by_reviewee(myte_records, myte_name)
    results: list[MyteValidationResult] = []

    for key, excel in excel_by_reviewee.items():
        if excel["subco"].strip() == "1":
            continue
        myte = myte_by_reviewee.get(key)
        if myte is None:
            results.append(_result(
                excel, None, "El reviewee registrado en el fichero excel no ha sido encontrado en myTE",
            ))
            continue

        status = myte["trStatus"].strip()
        if status in {"New", "Draft"}:
            results.append(_result(excel, myte, "myTE no enviado. TR Status es New o Draft.", True))
            continue
        if status != "Processed":
            results.append(_result(excel, myte, "Estado diferente de Processed. Revisar"))
            continue

        hours_field = "horasQuincena1" if period_number == 1 else "horasQuincena2"
        excel_hours = round(_parse_decimal(excel[hours_field], hours_field, excel["reviewee"]), 1)
        working_hours = round(
            _parse_decimal(myte["workingHours"], "workingHours", excel["reviewee"]), 1,
        )
        if excel_hours != working_hours:
            results.append(_result(
                excel, myte, "El numero de WorkingHours registradas en myTE es diferente al previsto en Excel.",
            ))
            continue

        percentage_raw = myte["percentageTotalHoursVsWorkingHours"]
        percentage = _parse_decimal(
            percentage_raw, "percentageTotalHoursVsWorkingHours", excel["reviewee"],
        )
        code = myte["chargeCode"]
        if "COMK3002" in code:
            passed = 91.0 <= percentage <= 93.0
        elif "COMK3001" in code:
            if key in SPECIAL_80_PERCENT_REVIEWEES:
                passed = percentage == 80.0
            else:
                passed = 82.0 <= percentage <= 86.0
        else:
            passed = False
        comments = "" if passed else (
            f"El porcentaje <{percentage_raw}> no cumple la politica definida para el chargeCode."
        )
        results.append(_result(excel, myte, comments))
    return results
