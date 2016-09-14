from openpyxl import load_workbook, Workbook
import usaddress
import os

fixture_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Fixtures")
infile = os.path.join(fixture_folder, "CityNames.xlsx")
wb = load_workbook(infile, read_only=True)
cities_ws = wb.get_sheet_by_name("Sheet1")


def remove_baggage(in_list):
    # Remove dates from the address
    rem = []
    for x in in_list:
        if r"/" in x:
            rem.append(x)

    # Remove keywords from the address
    keywords = ["ACRES", "DSD", "AC", "&", "TRACT", "MTG", "HM", "INVEST", "HONDA", "CHEV"]
    for x in in_list:
        if x in keywords:
            rem.append(x)

    # Remove values with ( or )
    for x in in_list:
        if "(" in x or ")" in x:
            rem.append(x)

    for r in rem:
        if r in in_list:
            in_list.remove(r)

    return in_list


class Parser:
    def __init__(self, address):
        self.description = address

    def parse(self):
        """return a list [hn, sn, apt, city, state, zipcode]"""
        address = self.description
        return

    def classify(self):
        """return a single index[0] value from the three following lists"""
        city = []
        state = []
        zipcode = []

        address = self.description
        address = address.replace(",", "")
        address = address.upper()
        split = address.split()

        cleaned = remove_baggage(split)

        # Test for the state
        states = ["GA", "FL", "SC", "NC", "TN", "VA", "WV", "KY", "AL"]
        for x in cleaned:
            if x in states:
                state.append(x)
        if len(state) > 1:
            print("More than one state parsed for {}".format(address))
        if len(state):
            for r in state:
                cleaned.remove(r)

        # Test for the zip code
        for x in cleaned:
            if u'{}'.format(x).isnumeric():
                if len(x) == 5:
                    zipcode.append(x)
        if len(zipcode) > 1:
            print("More than one zipcode parsed for {}".format(address))
        if len(zipcode):
            for r in zipcode:
                cleaned.remove(r)

        # Test for the city by matching with the census gazetteer table
        for x in cleaned:
            if x.isalpha() and len(x) > 2:
                # use the state to filter cities table and cross check city name
                for row in cities_ws.rows:
                    s = row[0].value.upper()
                    n = row[1].value
                    if s in state:
                        city_name = []
                        for c in ["city", "town"]:
                            if c in n:
                                city_name.extend(n.replace(c, "").rstrip().upper().split())
                                break
                        if len(city_name):
                            if x in city_name:
                                # Cross check the entire city name with the excel table
                                index_num = cleaned.index(x)
                                if len(city_name) > 1:
                                    fullname = " ".join(city_name)
                                    try:
                                        scnd_part = cleaned[index_num + 1]
                                        if "{} {}".format(x, scnd_part) == fullname:
                                            city.append(fullname)
                                            break
                                    except IndexError:
                                        print("Address does not match fullname")

                                else:
                                    fullname = city_name[0]
                                    if x == fullname:
                                        city.append(fullname)
                                        break

        if len(city) > 1:
            print("More than one city parsed for {} :: {}".format(address, city))
            return ["", "", ""]

        if len(city):
            for r in city:
                if r in cleaned:
                    cleaned.remove(r)

        answer = []
        try:
            answer.append(city[0])
        except IndexError:
            answer.append("")
        try:
            answer.append(state[0])
        except IndexError:
            answer.append("")
        try:
            answer.append(zipcode[0])
        except IndexError:
            answer.append("")

        return answer

    def usaddress_tag(self):
        address = self.description
        address = address.replace(",", "")
        address = address.upper()
        split = address.split()

        cleaned = remove_baggage(split)
        address = " ".join(cleaned)
        parsed = usaddress.parse(address)
        return parsed

if __name__ == "__main__":
    parser = Parser("178 RICHARDSON ST CROSS, SC 29436")
    x = parser.classify()
    y = parser.usaddress_tag()
    print x
    print y


