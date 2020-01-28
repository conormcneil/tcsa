#!/usr/bin/env python3

import sys
import csv
from os.path import splitext
import re
import jinja2
import xlsxwriter

env = jinja2.Environment(
    loader = jinja2.FileSystemLoader('templates')
)
TABS = [
    {
        "name": "all",
        "search": r'Produce',
        "replacements": [
            (r' - Rotation',''),
            (r' - Sunflower',''),
            (r' - Plain',''),
            (r' \(Weekly\)',''),
            (r' Goat',''),
            (r'^, ',''),
            (r', $','')
        ],
        "members": []
    },
    {
        "name": "bread",
        "search": r'Bread',
        "replacements": [
            (r' - Rotation',''),
            (r' - Plain/Herb',''),
            (r' - Sunflower',''),
            (r' - Plain',''),
            (r' \(Weekly\)',''),
            (r'[0-9] Goat Cheese',''),
            (r'[0-9] Produce',''),
            (r'[0-9] Sprouts',''),
            (r'[0-9] Mushrooms',''),
            (r'[0-9] Tortillas',''),
            (r'[0-9] Newsletter',''),
            (r', ',''),
            (r', $','')
        ],
        "members": []
    },
    {
        "name": "cheese",
        "search": r'Cheese',
        "replacements": [
            (r' - Sunflower',''),
            (r' \(Weekly\)',''),
            (r'[0-9] Bread',''),
            (r'[0-9] Produce',''),
            (r'[0-9] Sprouts',''),
            (r'[0-9] Mushrooms',''),
            (r'[0-9] Tortillas',''),
            (r'[0-9] Newsletter',''),
            (r', ',''),
            (r', $','')
        ],
        "members": []
    },
    {
        "name": "mushrooms",
        "search": r'Mushrooms',
        "replacements": [
            (r' - Rotation',''),
            (r' - Plain/Herb',''),
            (r' - Sunflower',''),
            (r' - Plain',''),
            (r' \(Weekly\)',''),
            (r'[0-9] Goat Cheese',''),
            (r'[0-9] Produce',''),
            (r'[0-9] Sprouts',''),
            (r'[0-9] Bread',''),
            (r'[0-9] Tortillas',''),
            (r'[0-9] Newsletter',''),
            (r', ',''),
            (r', $','')
        ],
        "members": []
    },
    {
        "name": "sprouts",
        "search": r'Sprouts',
        "replacements": [
            (r' - Plain/Herb',''),
            (r' - Plain',''),
            (r' \(Weekly\)',''),
            (r'[0-9] Goat Cheese',''),
            (r'[0-9] Produce',''),
            (r'[0-9] Bread',''),
            (r'[0-9] Mushrooms',''),
            (r'[0-9] Tortillas',''),
            (r'[0-9] Newsletter',''),
            (r'^, ',''),
            (r', $','')
        ],
        "members": []
    },
    {
        "name": "tortillas",
        "search": r'Tortillas',
        "replacements": [
            (r' - Rotation',''),
            (r' - Plain/Herb',''),
            (r' - Sunflower',''),
            (r' - Plain',''),
            (r' \(Weekly\)',''),
            (r'[0-9] Goat Cheese',''),
            (r'[0-9] Produce',''),
            (r'[0-9] Sprouts',''),
            (r'[0-9] Mushrooms',''),
            (r'[0-9] Bread',''),
            (r'[0-9] Newsletter',''),
            (r', ',''),
            (r', $','')
        ],
        "members": []
    }
]

def main(filename):
    make_workbook(filename)
    # TESTING:
    # make_csv(filename)

def make_workbook(filename):
    for record in records(filename):
        properties = parse_record(record)
        sort_record(properties)
    return save_xlsx(filename)

# useful for debugging
def make_csv(filename):
    for record in records(filename):
        properties = parse_record(record)
        sort_record(properties)
    save_as_csv()

def save_as_csv():
    template = env.get_template("sheets/list.csv")
    for tab in TABS:
        members = tab['members']
        for m in members:
            m['order'] = replace(m['order'],tab['replacements'])
        filename = tab['name'] + '.csv'
        with open(filename, "w") as outfile:
            outfile.write(template.render(members=members))

def save_xlsx(filename):
    xlsx_filename = splitext(filename)[0] + ".xlsx"
    workbook = xlsxwriter.Workbook(xlsx_filename)
    for tab in TABS:
        worksheet = workbook.add_worksheet(tab['name'])
        row = 0
        col = 0

        # write header row first
        worksheet.write(row, col  , "Check In"  )
        worksheet.write(row, col+1, "Last Name" )
        worksheet.write(row, col+2, "First Name")
        worksheet.write(row, col+3, "Order"     )
        row += 1

        for m in tab['members']:
            m['order'] = replace(m['order'],tab['replacements'])
            worksheet.write(row, col  , ""             )
            worksheet.write(row, col+1, m['last_name'] )
            worksheet.write(row, col+2, m['first_name'])
            worksheet.write(row, col+3, m['order']     )
            row += 1

    workbook.close()
    print(xlsx_filename)
    return xlsx_filename

def sort_record(properties):
    for tab in TABS:
        if tab['name'] in properties['tabs']:
            tab['members'].append(props_for_tab(properties,tab))

def props_for_tab(props,tab):
    local = {
        'last_name': props['last_name'],
        'first_name': props['first_name'],
        'order': replace(props['order'],tab['replacements'])
    }
    return local

def records(filename):
    with open(filename) as csvfile:
        csvreader = csv.reader(csvfile)
        return [x for x in csvreader][1:]

def parse_record(record):
    pickup_site, first_name, last_name, order = record
    order = parse_order(order)['order']
    tabs = parse_order(order)['tabs']
    return {
        "pickup_site": pickup_site,
        "first_name": first_name,
        "last_name": last_name,
        "order": order,
        "tabs": tabs
    }

def replace(order,replacements):
    local_order = order
    for old, new in replacements:
        local_order = re.sub(old,new,local_order)
    return local_order

def parse_order(order):
    REPLACEMENTS = [
        ('xShare',''),
        ('xBag','')
    ]
    order = replace(order,REPLACEMENTS)
    tabs = []
    for tab in TABS:
        if re.search(tab['search'],order) != None:
            tabs.append(tab['name'])
    return {'order':order,'tabs':tabs}

if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
