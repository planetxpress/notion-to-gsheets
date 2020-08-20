import datetime
import gspread
import json
import re
import os
import time
from google.cloud import secretmanager
from notion.client import NotionClient

credentials_key = os.getenv('CREDENTIALS_KEY')
notion_token = os.getenv('NOTION_TOKEN')
gsheet_id = os.getenv('GSHEET_ID')
notion_page = os.getenv('NOTION_PAGE')
project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
secret_client = secretmanager.SecretManagerServiceClient()
secret_path =  secret_client.secret_version_path(
    project_id, credentials_key, 'latest'
    )
service_account = json.loads(
    secret_client.access_secret_version(secret_path).payload.data.decode('UTF-8')
    )
gspread_auth = gspread.auth.ServiceAccountCredentials.from_service_account_info(
    info=service_account, scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
        ]
    )
gc = gspread.Client(auth=gspread_auth)
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
        if row.status not in status_key:
            status = ''
        else:
            status = status_key[row.status]

        item = {
            'Department': row.department,
            'Project': ', '.join(row.project),
            'Name': row.name,
            'Status': status,
            'Description': row.description,
            'Primary Person': ', '.join(row.primary_person),
            'Date Last Updated': row.date_last_updated,
            'Links': hyperlink,
            'Tags': ', '.join(row.tags)
        }
        for i in item:
            if not item[i]:
                item[i] = ''

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
    with open('status_key.json', 'r') as f:
        status_key = json.load(f)
    row_index = 1
    for row in data:
        row_index += 1
        status = row[status_index]
        if not status:
            continue
        cell = '%s%s' % (status_column, row_index)
        sheet.format(cell, status_key[status]['format'])
        sheet.update(cell, status_key[status]['value'])
        time.sleep(2)


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

    for index, row in enumerate(data):
        row[header.index('Date Last Updated')] = row[header.index('Date Last Updated')].strftime('%d %b %y')
        data[index] = row

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

    for index, row in enumerate(data):
        row[header.index('Date Last Updated')] = row[header.index('Date Last Updated')].strftime('%d %b %y')
        data[index] = row

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


def trigger(event, context):
    main()


if __name__ == '__main__':
    main()
