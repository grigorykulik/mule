import requests
import pandas as pd
import openpyxl
import constants

tradeable_items = ["Tattoo", "Omen", "DivinationCard", "Artifact", "Oil", "Incubator", "UniqueWeapon",
                   "UniqueArmour", "UniqueAccessory", "UniqueFlask", "UniqueJewel", "UniqueRelic", "ClusterJewel"]

url_template_items = "https://poe.ninja/api/data/itemoverview?league={}&type={}"


def get_prices_from_league(league_name, tradeable_item):
    url = url_template_items.format(league_name, tradeable_item)
    responce = requests.get(url)

    if responce.status_code == 200:
        return responce.json()
    else:
        print("Error: ")


def get_names_and_prices(item_list):
    result = []

    for item in item_list:
        item_id = item["id"]

        item_links = item.get("links", "")
        if item_links == "":
            item_name = item["name"]
        else:
            item_name = item["name"] + ", " + str(item["links"]) + "L"

        item_details_id = item.get("detailsId", "")
        if "relic" in item_details_id:
            item_name = item_name + "(relic)"

        if item.get("variant", "") != "":
            item_name += (", " + item["variant"])

        if item.get("itemType", "") == "Jewel":
            item_name += (", " + item["detailsId"].split("-")[-2])

        item_base_type = item["baseType"]
        item_price = item["chaosValue"]
        dictionary = {"id": item_id, "name": item_name, "base type": item_base_type, "price": item_price}
        result.append(dictionary)
    return result


def find_league_in_standard(item_id, name_and_price_standard):
    chaos_value = "ITEM NOT FOUND"

    for item in name_and_price_standard:
        if item["id"] == item_id:
            chaos_value = item["price"]

    return chaos_value


def merge_results(name_and_price_league, name_and_price_standard):
    merged_list = []

    for league_item in name_and_price_league:
        item_id = league_item["id"]
        item_name = league_item["name"]
        item_base_type = league_item["base type"]
        item_price_league = league_item["price"]
        item_price_standard = find_league_in_standard(item_id, name_and_price_standard)

        merged_entry = {"id": item_id, "name": item_name, "base type": item_base_type, "price in league": item_price_league,
                        "price in standard": item_price_standard}
        merged_list.append(merged_entry)
    return merged_list


def get_price_diff(item_list):
    for item in item_list:
        if item["price in standard"] != "ITEM NOT FOUND":
            item["price diff"] = float(item["price in standard"]) - float(item["price in league"])
        else:
            item["price diff"] = "THIS ITEM IS NOT IN THE STANDARD LEAGUE"
    return item_list


def create_output_file():
    file_name = "results.xlsx"
    pd.DataFrame().to_excel(file_name, index=False)


def write_data_to_file(item_list):
    with pd.ExcelWriter("results.xlsx", mode="a", engine="openpyxl") as writer:
        df = pd.DataFrame(item_list)
        df.to_excel(writer, sheet_name=tradeable_item, index=False)


def format_resulting_file():
    file_path = "results.xlsx"
    wb = openpyxl.load_workbook(file_path)
    worksheet_names = wb.sheetnames

    for sheet_name in worksheet_names:
        ws = wb[sheet_name]
        ws.column_dimensions["A"].width = 10

        for column_letter in ['B', 'C', 'F']:
            ws.column_dimensions[column_letter].width = 50

        for column_letter in ['D', 'E']:
            ws.column_dimensions[column_letter].width = 20

    wb.save(file_path)


if __name__ == "__main__":
    print(constants.banner)
    create_output_file()

    print(constants.league_type)

    user_input = input()

    while user_input != '1' and user_input != '2':
        print("Unknown league type")
        user_input = input()

    if user_input == '1':
        current_league = "Ancestor"
    elif user_input == '2':
        current_league = 'Hardcore+Ancestor'

    for tradeable_item in tradeable_items:

        print("Retrieving {} from the current league".format(tradeable_item))
        item_list_league = get_prices_from_league(current_league, tradeable_item)["lines"]

        print("Extracting prices from the server response.")
        name_and_price_league = get_names_and_prices(item_list_league)

        print("Retrieving {} from the standard league".format(tradeable_item))
        item_list_standard = get_prices_from_league("Standard", tradeable_item)["lines"]

        print("Extracting prices from the server response.")
        name_and_price_standard = get_names_and_prices(item_list_standard)

        print("Comparing prices.")
        merged_lists = merge_results(name_and_price_league, name_and_price_standard)
        merged_lists = get_price_diff(merged_lists)

        print("Writing data to file.")
        write_data_to_file(merged_lists)
        print("Done. \n")

    format_resulting_file()
