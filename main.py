import datetime
import math
import os
import time
from json import JSONDecodeError
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlsxwriter
from bs4 import BeautifulSoup


def preformated(unformated):
    splitter = unformated.split()
    day = splitter[2]
    month = splitter[3]
    year = splitter[4]
    if month == 'Jan':
        month = '1'
    elif month == 'Feb':
        month = '2'
    elif month == 'Mar':
        month = '3'
    elif month == 'Apr':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'Jun':
        month = '6'
    elif month == 'Jul':
        month = '7'
    elif month == 'Aug':
        month = '8'
    elif month == 'Sep':
        month = '9'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


page_number = 0
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive"
}
dest_codes = ["ALCAN", "DUBAI", "FAR.E", "AUSTL", "BAHAM", "BERMU", "ATLCO", "CARIB", "CUBAN", "EUROP", "HAWAI",
              "PACIF", "ISLAN", "SOPAC", "T.ATL", "TPACI"]


def get_pages_count(current_code):
    max_pages = 0
    tmpurl = "https://secure.royalcaribbean.com/ajax/cruises/searchbody?destinationRegionCode_" + current_code + \
             "=true&currentPage=0&action=update"
    tmppage = requests.get(tmpurl, headers=headers)
    tmpsoup = BeautifulSoup(tmppage.text, "lxml")
    for tmplink in tmpsoup.find_all("h3", {"class": "matching-cruises hide-for-small-only"}):
        all_itineraries = math.ceil(int(tmplink.text.split()[0]))
        max_pages = math.ceil(all_itineraries / 10)
    return max_pages


total_cruises = 0
list_to_write = []
url_list = []
packages = []
mini_list = []

print("Downloading itinerary codes")


def ged_codes(code):
    pages_count = get_pages_count(code)
    start_page = 0
    while start_page <= pages_count:
        url = "https://secure.royalcaribbean.com/ajax/cruises/searchbody?vacationTypeCode_CO=true&destinationRegionCode_" + code \
              + "=true&currentPage=" + str(start_page) + "&action=update"
        page = requests.get(url, headers=headers)
        soup = BeautifulSoup(page.text, "lxml")
        for link_line in soup.find_all("a"):
            if "cruises" in str(link_line.get("href")):
                if "-" in str(link_line.get("href")) and "currencyCode" in str(link_line.get("href")):
                    src = str(link_line.get("href"))
                    src = src.replace("/cruises/", "").replace("2F", "")
                    elements = src.split("?")
                    tmp_string = '' + elements[0] + ',' + code
                    if tmp_string in url_list:
                        pass
                    else:
                        url_list.append(tmp_string)
                        print("Processing itinerary...", tmp_string)
        start_page += 1


pool2 = ThreadPool(5)
pool2.map(ged_codes, dest_codes)
pool2.close()
pool2.join()

print("Processing....")


def get_destination(param):
    if param == 'ALCAN':
        return ['Alaska', 'A']
    elif param == 'DUBAI':
        return ['Arabian Gulf', 'AG']
    elif param == 'FAR.E':
        return ['Asia', 'O']
    elif param == 'AUSTL':
        return ['AU/NZ', 'P']
    elif param == 'BAHAM':
        return ['Bahamas', 'BH']
    elif param == 'BERMU':
        return ['Bermuda', 'BM']
    elif param == 'ATLCO':
        return ['Can/New En', 'NN']
    elif param == 'CARIB':
        return ['Carib', 'C']
    elif param == 'CUBAN':
        return ['Cuba', 'S']
    elif param == 'EUROP':
        return ['Europe', 'E']
    elif param == 'HAWAI':
        return ['Hawaii', 'H']
    elif param == 'PACIF':
        return ['Pacific', 'A']
    elif param == 'ISLAN':
        return ['Repositioning', 'R']
    elif param == 'SOPAC':
        return ['South Pacific', 'I']
    elif param == 'T.ATL':
        return ['Transatlantic', 'E']
    elif param == 'TPACI':
        return ['Transpacific', 'I']


def calculate_days(sail_date_param, number_of_nights_param):
    date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
    try:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param.split("-")[1]))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def get_vessel_id(ves_name):
    if ves_name == "Anthem of the Seas":
        return "859"
    elif ves_name == "Ovation of the Seas":
        return "931"
    elif ves_name == "Quantum of the Seas":
        return "860"
    elif ves_name == "Allure of the Seas":
        return "717"
    elif ves_name == "Harmony of the Seas":
        return "941"
    elif ves_name == "Oasis of the Seas":
        return "691"
    elif ves_name == "Freedom of the Seas":
        return "502"
    elif ves_name == "Independence of the Seas":
        return "581"
    elif ves_name == "Liberty of the Seas":
        return "561"
    elif ves_name == "Adventure of the Seas":
        return "378"
    elif ves_name == "Explorer of the Seas":
        return "217"
    elif ves_name == "Mariner of the Seas":
        return "417"
    elif ves_name == "Navigator of the Seas":
        return "408"
    elif ves_name == "Voyager of the Seas":
        return "237"
    elif ves_name == "Brilliance of the Seas":
        return "399"
    elif ves_name == "Jewel of the Seas":
        return "432"
    elif ves_name == "Radiance of the Seas":
        return "225"
    elif ves_name == "Serenade of the Seas":
        return "416"
    elif ves_name == "Enchantment of the Seas":
        return "216"
    elif ves_name == "Grandeur of the Seas":
        return "218"
    elif ves_name == "Legend of the Seas":
        return "219"
    elif ves_name == "Rhapsody of the Seas":
        return "228"
    elif ves_name == "Vision of the Seas":
        return "236"
    elif ves_name == "Majesty of the Seas":
        return "220"
    elif ves_name == "Empress of the Seas":
        return "999"
    else:
        return ""


def format_date_for_dateline(osd):
    splitter = osd.split()
    day = splitter[0]
    month = splitter[1]
    year = splitter[2]
    if month == 'Jan':
        month = '01'
    elif month == 'Feb':
        month = '02'
    elif month == 'Mar':
        month = '03'
    elif month == 'Apr':
        month = '04'
    elif month == 'May':
        month = '05'
    elif month == 'Jun':
        month = '06'
    elif month == 'Jul':
        month = '07'
    elif month == 'Aug':
        month = '08'
    elif month == 'Sep':
        month = '09'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s-%s-%s' % (year, month, day)
    return final_date


def parse(index):
    ship = index.split(",")
    url = "http://www.royalcaribbean.com/ajax/cruise/inlinepricing/" + ship[
        0] + "?currencyCode=USD&sCruiseType=CO&_=1481317060902"
    cruise_info = {}
    try:
        page = requests.get(url, headers=headers)
        cruise_info = page.json()
    except JSONDecodeError:
        print("retry...")
        time.sleep(1)
        try:
            page = requests.get(url, headers=headers)
            cruise_info = page.json()
        except JSONDecodeError:
            print("retry...")
            time.sleep(1)
            try:
                page = requests.get(url, headers=headers)
                cruise_info = page.json()
            except JSONDecodeError:
                print("Skipping.................")
                print("Skipped params in first request: " + ship[0] + "-" + ship[1])
    inline_pricing = cruise_info["inlinePricing"]
    rows = inline_pricing["rows"]
    for r in rows:
        cruise_line_name = "Royal Caribbean"
        cruise_id = "14"
        itinerary_id = ""
        oceanview_bucket_price = (r["priceItems"][1]["price"])
        if oceanview_bucket_price is not None:
            oceanview_bucket_price = str.replace(oceanview_bucket_price, "$", "").replace(",", "")
        else:
            oceanview_bucket_price = "N/A"
        balcony_bucket_price = (r["priceItems"][2]["price"])
        if balcony_bucket_price is not None:
            balcony_bucket_price = str.replace(balcony_bucket_price, "$", "").replace(",", "")
        else:
            balcony_bucket_price = "N/A"
        suite_bucket_price = (r["priceItems"][3]["price"])
        if suite_bucket_price is not None:
            suite_bucket_price = str.replace(suite_bucket_price, "$", "").replace(",", "")
        else:
            suite_bucket_price = "N/A"
        interior_bucket_price = (r["priceItems"][0]["price"])
        if interior_bucket_price is not None:
            interior_bucket_price = str.replace(interior_bucket_price, "$", "").replace(",", "")
        else:
            interior_bucket_price = "N/A"
        sail_date = preformated(r["dateLabel"])
        url = "https://secure.royalcaribbean.com/ajax/cruise/" \
              "pricing/" + cruise_info["title"].strip().replace("/", '').replace(" ", '').replace('.', '') + "-" + \
              cruise_info["packageId"].strip() + \
              "?currencyCode=USD&sCruiseType=CO&sailDate=" + sail_date + "&_=1481040091122"
        if url in packages:
            print("duplicate")
            continue
        else:
            packages.append(url)
        try:
            page = requests.get(url.replace(' ', ''), headers=headers)
            details = page.json()
        except JSONDecodeError:
            print("retry...")
            try:
                page = requests.get(url.replace(' ', ''), headers=headers)
                details = page.json()
            except JSONDecodeError:
                print("retry...")
                try:
                    page = requests.get(url.replace(' ', ''), headers=headers)
                    details = page.json()
                except JSONDecodeError:
                    print("Skipping in second request..................")
                    print(url.replace(' ', ''))
                    continue
        brochure_name = details["title"]
        vessel_name = details['shipText']
        number_of_nights = int(details["title"].split()[0])
        destination = get_destination(ship[1])
        destination_code = destination[1]
        destination_name = destination[0]
        vessel_id = get_vessel_id(vessel_name)
        return_date = calculate_days(sail_date, str(number_of_nights))
        temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                itinerary_id, brochure_name, number_of_nights, sail_date, return_date, interior_bucket_price,
                oceanview_bucket_price, balcony_bucket_price, suite_bucket_price]
        print("Processing: ", brochure_name)

        mini_list.append(temp)


results = list(set(url_list))
pool = ThreadPool(5)
pool.map(parse, results)
pool.close()
pool.join()
list_to_write.append(mini_list)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Royal Caribbean.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship_entry in data_array:
        column_count = 0
        for en in ship_entry:
            if column_count == 0:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 1:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 2:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 3:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 4:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 5:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 6:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 7:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 8:
                try:
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 9:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 10:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 11:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 12:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 13:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 14:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            column_count += 1
        row_count += 1
    workbook.close()
    pass


write_file_to_excell(mini_list)
