import requests
import json
import pytest
import xlsxwriter

@pytest.mark.usefixtures("api_url", "api_key")
class TestGetMail:

#Send request and get a response from server
    def test_getMail(self, api_url, api_key):
        response = requests.get(api_url + "/api/addresses/mailsacfortesting@mailsac.com/messages", headers = api_key)
        assert response.status_code == 200, 'Status code is invalid.'
        response_json = json.loads(response.text)
        print(response_json)

#Print out email info to the console
        elements = len([item.get('_id') for item in response_json])
        for x in range(0, elements):
            print("Id", x + 1)
            print("Name: ", response_json[x]['from'][0]['name'])
            print("Email: ", response_json[x]['from'][0]['address'])
            print("Subject: ", response_json[x]['subject'])
            print("Original Id: ", response_json[x]['_id'])

#Save email info to the excel file
        titles = ["Id", "Name", "Email", "Subject", "Original Id"]
        row = 0
        column = 0
        workbook = xlsxwriter.Workbook('Email_info.xlsx')
        worksheet = workbook.add_worksheet()
        for item in titles:
            worksheet.write(row, column, item)
            column += 1

        row = 1
        column = 0
        for x in range(0, elements):
            worksheet.write(row, column, x + 1)
            worksheet.write(row, column + 1, response_json[x]['from'][0]['name'])
            worksheet.write(row, column + 2, response_json[x]['from'][0]['address'])
            worksheet.write(row, column + 3, response_json[x]['subject'])
            worksheet.write(row, column + 4, response_json[x]['_id'])
            row += 1

        workbook.close()