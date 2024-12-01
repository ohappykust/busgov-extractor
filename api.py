from typing import Any

import requests
import xlsxwriter
from colorama import Back
from future.backports.datetime import datetime

from tqdm import tqdm

from consts import (
    TARGETED_FUNDS_OPS_BUDGET_HEADERS,
    TARGETED_FUNDS_OPS_SUBSIDIES_HEADERS,
    BASIC_ORGS_HEADERS,
    QUALITY_ORGS_HEADERS,
    BUILDING_EXEC_INFO_HEADERS,
    UNAVAILABLE_ORGS_HEADERS
)


def get_optional_field_value(value):
    return value if value else "-"


def download_data(
    regions: list[str],
    vgu_name: list[str],
    vgu_ids: list[str],
    areas: list[str] | None = None,
    city: list[str] | None = None
) -> dict[str, Any] | None:
    print("(1/4) Загрузка всех организаций")

    all_orgs_url = (
        "https://bus.gov.ru/public-rest/api/orgunique/extendedSearchOrgUnique?"
        "orderAttributeName=conformity&orderDirectionASC=false&searchTermCondition=or&pageNumber=1&pageSize=100000"
    )

    regions_query_params = "&".join([f"regions={val}" for val in regions])
    vgu_name_query_params = "&".join([f"vguName={val}" for val in vgu_name])
    vgu_ids_query_params = "&".join([f"vguIds={val}" for val in vgu_ids])

    if areas:
        all_orgs_url = f"{all_orgs_url}&areas={areas}"
    if city:
        all_orgs_url = f"{all_orgs_url}&city={city}"

    all_orgs_url = f"{all_orgs_url}&{regions_query_params}&{vgu_name_query_params}&{vgu_ids_query_params}"

    headers = {
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/93.0.4577.63 Safari/537.36"
    }
    all_orgs_response = requests.get(all_orgs_url, headers=headers)

    if not all_orgs_response.ok:
        print(Back.RED + "Произошла ошибка при загрузке всех организаций.")
        input("Нажмите Enter чтобы выйти...")
        exit(-1)

    all_orgs_data = all_orgs_response.json()

    if not all_orgs_data.get('orgs'):
        print(Back.RED + "Не найдено ни одной организации по заданным фильтрам.")
        input("Нажмите Enter чтобы выйти...")
        exit(-2)

    all_orgs_agency_ids = [org["agencyId"] for org in all_orgs_data["orgs"]]

    print("(2/4) Загрузка информации об организациях")

    basic_orgs_data_url = "https://bus.gov.ru/public-rest/api/agency/compare?selectedYear=2023&compareAgencyIds="
    basic_orgs_data = {}
    unavailable_orgs_basic_data = []

    for agency_id in tqdm(all_orgs_agency_ids):
        basic_org_data_response = requests.get(f"{basic_orgs_data_url}{agency_id}", headers=headers)

        if not basic_org_data_response.ok:
            unavailable_orgs_basic_data.append(
                next(org for org in all_orgs_data["orgs"] if org["agencyId"] == agency_id)
            )
            continue

        basic_orgs_data[agency_id] = basic_org_data_response.json()

    print("(3/4) Загрузка оценок качества организаций")

    quality_orgs_data_url = "https://bus.gov.ru/public-rating/api/ratingCompare/commonInfo"
    compare_agency_ids_query_params = "&".join([f"compareAgencyIds={val}" for val in all_orgs_agency_ids])

    quality_orgs_data_response = requests.get(
        f"{quality_orgs_data_url}?{compare_agency_ids_query_params}", headers=headers
    )

    if not quality_orgs_data_response.ok:
        print("Произошла ошибка при загрузке оценок качества организаций.")
        input("Нажмите Enter чтобы выйти...")
        exit(-3)

    quality_orgs_data = quality_orgs_data_response.json()

    return {
        "allOrgsData": all_orgs_data,
        "allOrgsAgencyIds": all_orgs_agency_ids,
        "basicOrgsData": basic_orgs_data,
        "unavailableOrgsBasicData": unavailable_orgs_basic_data,
        "qualityOrgsData": quality_orgs_data,
    }


def write_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    rows: list[list[str]],
    headers: list[str]
) -> None:
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.write_row('A1', headers)

    for idx, row in enumerate(rows):
        for col, val in enumerate(row):
            if col == 0:
                worksheet.write_url(idx + 1, 0, f"https://bus.gov.ru/info-card/{row[0]}", string=str(row[0]))
                continue
            worksheet.write(idx + 1, col, val)

    column_settings = [{"header": header} for header in headers]
    worksheet.add_table(0, 0, len(rows), len(headers) - 1, {
        "columns": column_settings,
        "autofilter": True
    })

    worksheet.autofit()


def generate_xlsx(
    regions: list[str],
    vgu_name: list[str],
    vgu_ids: list[str],
    areas: list[str] | None = None,
    city: list[str] | None = None
) -> None:
    download_data_result = download_data(regions, vgu_name, vgu_ids, areas, city)
    if not download_data_result:
        return

    print("(4/4) Формирование Excel файла")

    basic_orgs_data = download_data_result["basicOrgsData"]
    quality_orgs_data = download_data_result["qualityOrgsData"]

    basic_orgs_data_rows = []
    quality_info_rows = []
    building_exec_info_rows = []
    targeted_funds_ops_budget_rows = []
    targeted_funds_ops_subsidies_rows = []
    unavailable_orgs_rows = []

    for agency_id, current_org in tqdm(basic_orgs_data.items()):
        quality_org_data = quality_orgs_data.get(str(agency_id))

        org_name = current_org["agenciesData"]["commontab.name"][0]
        short_org_name = current_org["agenciesData"]["commontab.short_name"][0]
        public_law_education = current_org["agenciesData"]["commontab.ppo.name"][0]
        func_body = current_org["agenciesData"]["commontab.founderAgency.shortClientName"][0]
        grbs_code = current_org["agenciesData"]["commontab.rgbs.code.chapter"][0]
        rbs_agency_name = current_org["agenciesData"]["commontab.rbsAgency.name"][0]
        org_type = current_org["agenciesData"]["commontab.agency.type"][0]
        org_kind = current_org["agenciesData"]["commontab.agency.kind"][0]
        okato = current_org["agenciesData"]["commontab.okato"][0]
        okfs_ownership_type = current_org["agenciesData"]["commontab.okfs.name"][0]
        okfs_ownership_kind = current_org["agenciesData"]["commontab.okopf.kind"][0]
        actual_location_address = current_org["agenciesData"]["commontab.agencyAddress"][0]
        supervisor = current_org["agenciesData"]["commontab.manager"][0]
        telephone = current_org["agenciesData"]["commontab.manager.phone"][0]
        url = current_org["agenciesData"]["commontab.website"][0]
        email = current_org["agenciesData"]["commontab.email"][0]
        parent_name = current_org["agenciesData"]["commontab.branch.parent.name"][0]
        act_type = current_org["agenciesData"]["commontab.act.type"][0]
        approver_organization_name = current_org["agenciesData"]["commontab.act.approverOrganizationName"][0]
        act_date = current_org["agenciesData"]["commontab.act.date"][0]
        act_number = current_org["agenciesData"]["commontab.act.number"][0]
        act_name = current_org["agenciesData"]["commontab.act.name"][0]

        basic_orgs_data_rows.append([
            agency_id,
            get_optional_field_value(org_name),
            get_optional_field_value(short_org_name),
            get_optional_field_value(public_law_education),
            get_optional_field_value(func_body),
            get_optional_field_value(grbs_code),
            get_optional_field_value(rbs_agency_name),
            get_optional_field_value(org_type),
            get_optional_field_value(org_kind),
            get_optional_field_value(okato),
            get_optional_field_value(okfs_ownership_type),
            get_optional_field_value(okfs_ownership_kind),
            get_optional_field_value(actual_location_address),
            get_optional_field_value(supervisor),
            get_optional_field_value(telephone),
            get_optional_field_value(url),
            get_optional_field_value(email),
            get_optional_field_value(parent_name),
            get_optional_field_value(act_type),
            get_optional_field_value(approver_organization_name),
            get_optional_field_value(act_date),
            get_optional_field_value(act_number),
            get_optional_field_value(act_name)
        ])

        if quality_org_data.get("scopeWithRatingsDtos"):
            rating_details = quality_org_data["scopeWithRatingsDtos"][0]["ratingDetailsDto"][0]
            quality_info_rows.append([
                agency_id,
                get_optional_field_value(org_name),
                get_optional_field_value(short_org_name),
                quality_org_data["ratingYear"],
                get_optional_field_value(
                    rating_details["organizationGroup"]["groupName"]
                    if rating_details["organizationGroup"] else None),
                get_optional_field_value(rating_details["globalPlaceValue"]),
                get_optional_field_value(rating_details["opennessValue"]),
                get_optional_field_value(rating_details["comfortValue"]),
                get_optional_field_value(rating_details["timeoutValue"]),
                get_optional_field_value(rating_details["goodwillValue"]),
                get_optional_field_value(rating_details["contentmentValue"])
            ])

        agencies_tasks = current_org.get("agenciesTasks", {}).get(f"value_{agency_id}", [])
        for i in range(0, len(agencies_tasks), 3):
            if "itemData" in agencies_tasks[i]:
                building_exec_info_rows.append([
                    agency_id,
                    get_optional_field_value(org_name),
                    get_optional_field_value(short_org_name),
                    "Услуги",
                    agencies_tasks[i]["itemData"],
                    agencies_tasks[i + 1]["itemData"],
                    agencies_tasks[i + 2]["itemData"]
                ])

        agencies_works = current_org.get("agenciesWorks", {}).get(f"value_{agency_id}", [])
        for i in range(0, len(agencies_works), 2):
            if "itemData" in agencies_works[i] and agencies_works[i]["itemData"]:
                building_exec_info_rows.append([
                    agency_id,
                    get_optional_field_value(org_name),
                    get_optional_field_value(short_org_name),
                    "Работы",
                    agencies_works[i]["itemData"],
                    agencies_works[i + 1]["itemData"],
                    "-"
                ])

        for funds_budget_op in current_org.get("budgetInvestmentsTable"):
            targeted_funds_ops_budget_rows.append([
                agency_id,
                get_optional_field_value(org_name),
                get_optional_field_value(short_org_name),
                current_org["budgetOperation"]["budget.operation.okato"][0],
                current_org["budgetOperation"]["budget.operation.year"][0],
                current_org["budgetOperation"]["budget.operation.sum.planned.all"][0],
                current_org["budgetOperation"]["budget.operation.subsidies.all"][0],
                funds_budget_op[0]["name"],
                funds_budget_op[0]["sum"],
            ])

        for funds_subsidies_op in current_org.get("budgetSubsidiesTable"):
            targeted_funds_ops_subsidies_rows.append([
                agency_id,
                get_optional_field_value(org_name),
                get_optional_field_value(short_org_name),
                current_org["budgetOperation"]["budget.operation.okato"][0],
                current_org["budgetOperation"]["budget.operation.year"][0],
                current_org["budgetOperation"]["budget.operation.sum.planned.all"][0],
                current_org["budgetOperation"]["budget.operation.subsidies.all"][0],
                funds_subsidies_op["code"][0],
                funds_subsidies_op["grantName"][0],
                funds_subsidies_op["sumPlannedReceips"][0]
            ])

    for org in download_data_result["unavailableOrgsBasicData"]:
        unavailable_orgs_rows.append([
            org["agencyId"],
            org["fullName"],
            org["fullAddress"],
            org["phone"],
            org["webSite"]
        ])

    workbook = xlsxwriter.Workbook(f"Данные bus.gov.ru от {datetime.now().strftime('%d.%m.%Y %H:%M')}.xlsx")

    write_sheet(workbook, "Общая информация", basic_orgs_data_rows, BASIC_ORGS_HEADERS)
    write_sheet(workbook, "Независимая оценка качества", quality_info_rows, QUALITY_ORGS_HEADERS)
    write_sheet(workbook, "Гос. здание и его исполнения", building_exec_info_rows, BUILDING_EXEC_INFO_HEADERS)
    write_sheet(workbook, "Операции с бюджет. инвестициями", targeted_funds_ops_budget_rows,
                TARGETED_FUNDS_OPS_BUDGET_HEADERS)
    write_sheet(workbook, "Операции с субсидиями", targeted_funds_ops_subsidies_rows,
                TARGETED_FUNDS_OPS_SUBSIDIES_HEADERS)
    write_sheet(workbook, "Организации без данных", unavailable_orgs_rows, UNAVAILABLE_ORGS_HEADERS)

    workbook.close()
