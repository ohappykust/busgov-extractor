from typing import Any
from urllib.parse import urlparse, parse_qs

from colorama import init, Back

from api import generate_xlsx

init(autoreset=True)

def parse_url(url) -> dict[str, Any] | bool:
    url = urlparse(url)
    params = parse_qs(url.query)

    regions = params.get("regions", [None])[0]
    areas = params.get("areas", [None])[0]
    city = params.get("city", [None])[0]
    vgu_name = params.get("vguName", [None])[0]
    vgu_ids = params.get("vguIds", [None])[0]

    if regions:
        regions = regions.split(', ')
    if areas:
        areas = areas.split(', ')
    if vgu_name:
        vgu_name = vgu_name.split(', ')
    if vgu_ids:
        vgu_ids = vgu_ids.split(', ')

    if areas and areas[0] == "empty":
        areas = None
    if city == "empty":
        city = None

    if not regions or regions[0] == "empty":
        return False
    if not vgu_name or vgu_name[0] == "empty":
        return False

    return {
        "regions": regions,
        "areas": areas,
        "city": city,
        "vgu_name": vgu_name,
        "vgu_ids": vgu_ids
    }


def callback(url: str) -> dict[str, Any] | None:
    if not (parsed_url_data := parse_url(url)):
        print(Back.RED + "Некорректная ссылка. Убедитесь, что выбран регион и вид учреждения в фильтре.")
        return None
    return parsed_url_data


def main():
    while True:

        url = input("Введите ссылку на страницу с данными: ")
        if not (parsed_url_data := parse_url(url)):
            print(Back.RED + "Некорректная ссылка. Убедитесь, что выбран регион и вид учреждения в фильтре.")
            continue

        print(
            " #########################################################\n",
            "#                     Полученные данные                 #\n",
            "#########################################################\n",
            "Регионы: " + ", ".join(parsed_url_data["regions"]) + "\n",
            "Области: " + (", ".join(parsed_url_data["areas"]) if parsed_url_data["areas"] else "Не указано") + "\n",
            "Город: " + str(parsed_url_data["city"] if parsed_url_data["city"] else "Не указано") + "\n",
            "Названия ВГУ: " + ", ".join(parsed_url_data["vgu_name"]) + "\n",
            "Идентификаторы ВГУ: " + ", ".join(parsed_url_data["vgu_ids"]) + "\n",
            "#########################################################\n",
        )

        confirm = input("Начать выгрузку данных? (Д/н): ")
        if confirm.lower() and confirm.lower() != "д" and confirm.lower() != "y":
            continue

        generate_xlsx(parsed_url_data["regions"], parsed_url_data["vgu_name"], parsed_url_data["vgu_ids"], parsed_url_data["areas"], parsed_url_data["city"])
        print(Back.LIGHTGREEN_EX + "Excel файл успешно сформирован!")
        break


if __name__ == "__main__":
    main()
