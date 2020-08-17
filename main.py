import gspread
import json
import re
import os
import time
from datetime import datetime
from notion.client import NotionClient

notion_token = os.getenv('NOTION_TOKEN')
gsheet_id = os.getenv('GSHEET_ID')
notion_page = os.getenv('NOTION_PAGE')
gc = gspread.service_account(filename='C:\\Users\\Josh\Dropbox\\gcpcreds.json')
ss = gc.open_by_key(gsheet_id)


def get_notion_data():
    notion = list()
    client = NotionClient(token_v2=notion_token)
    cv = client.get_collection_view(notion_page)
    for row in cv.collection.get_rows():

        # Convert named status to numbers for sorting
        status_key = {
            'Red': '1',
            'Yellow': '2',
            'Green': '3',
            'Delayed': '4',
            'Complete': 'Complete'
        }

        # Convert hyperlink to not look like garbage
        link_match = re.match(r'(http|https)://([A-Za-z\-.]*)/', row.links)
        if link_match:
            hyperlink = '=HYPERLINK("%s","%s")' % (row.links, link_match.group(2))
        else:
            hyperlink = ''

        item = {
            'Department': row.department,
            'Project': ', '.join(row.project),
            'Name': row.name,
            'Status': status_key[row.status],
            'Description': row.description,
            'Primary Person': ', '.join(row.primary_person),
            'Date Last Updated': row.date_last_updated.strftime('%Y-%m-%d'),
            'Links': hyperlink,
            'Tags': ', '.join(row.tags)
        }
        notion.append(item)
    return notion


def reset_format(sheet):
    sheet.format('2:1000', {
        'wrapStrategy': 'WRAP',
        'textFormat': {
            'fontFamily': 'Calibri',
            'fontSize': 11
        },
        'backgroundColor': {
            'red': 1.0,
            'green': 1.0,
            'blue': 1.0
        }
    })


def format_status(header, data, sheet):
    status_index = header.index('Status')
    status_column = chr(ord('@') + status_index + 1)
    row_index = 2
    with open('status_key.json', 'r') as f:
        status_key = json.load(f)
    for row in data:
        status = row[status_index]
        cell = '%s%s' % (status_column, row_index)
        sheet.format(cell, status_key[status]['format'])
        sheet.update(cell, status_key[status]['value'])
        row_index += 1
        time.sleep(1.5)


def format_date(header, data, sheet):
    date_index = header.index('Date Last Updated')
    date_column = chr(ord('@') + date_index + 1)
    row_index = 2
    for row in data:
        if not row[date_index]:
            continue
        epoch = int(row[date_index])
        date_string = datetime.fromtimestamp(epoch).strftime('%d-%b-%y')
        cell = '%s%s' % (date_column, row_index)
        sheet.update(cell, date_string)
        row_index +=1
        time.sleep(1)


def in_progress(notion):
    data = []
    sheet = ss.get_worksheet(0)
    header = sheet.row_values(1)
    ss.values_clear('In Progress!2:1000')
    reset_format(sheet)

    for entry in notion:
        if entry['Status'] == 'Complete':
            continue
        row = []
        for column in header:
            row.append(entry[column])
        data.append(row)

    # Sort Priority
    data = sorted(data,key=lambda x: x[header.index('Date Last Updated')], reverse=True)
    data = sorted(data,key=lambda x: x[header.index('Name')])
    data = sorted(data,key=lambda x: x[header.index('Status')])
    data = sorted(data,key=lambda x: x[header.index('Project')])
    data = sorted(data,key=lambda x: x[header.index('Department')])

    sheet.update('A2', data, value_input_option='USER_ENTERED')
    format_status(header, data, sheet)


def completed(notion):
    data = []
    sheet = ss.get_worksheet(1)
    header = sheet.row_values(1)
    ss.values_clear('Completed!2:1000')
    reset_format(sheet)

    for entry in notion:
        if entry['Status'] != 'Complete':
            continue
        row = []
        for column in header:
            row.append(entry[column])
        data.append(row)

    # Sort Priority
    data = sorted(data,key=lambda x: x[header.index('Name')])
    data = sorted(data,key=lambda x: x[header.index('Project')])
    data = sorted(data,key=lambda x: x[header.index('Date Last Updated')], reverse=True)
    data = sorted(data,key=lambda x: x[header.index('Department')])

    sheet.update('A2', data, value_input_option='USER_ENTERED')
    status_column = chr(ord('@') + header.index('Status') + 1)
    with open('status_key.json', 'r') as f:
        status_key = json.load(f)
    num_rows = len(data) + 1
    sheet.format(
        '{status_column}2:{status_column}{num_rows}'.format(
            status_column=status_column,
            num_rows=num_rows
        ), status_key['Complete']['format'])


def main():
    notion = get_notion_data()
    in_progress(notion)
    completed(notion)


if __name__ == '__main__':
    main()
