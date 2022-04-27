import datetime
from dateutil.parser import parse
import aiohttp
import asyncio
import json
import time
from bs4 import BeautifulSoup as bs
import os
import binascii
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook, load_workbook


base_url = 'https://apps.daysmartrecreation.com/dash/admin/index.php'
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 ' \
             'Safari/537.36 '


class DashRequests:
    def __init__(self, sema, session):
        self.sema = sema
        self.session = session
        self.all_searches = self.csrf_code = self.input_control_params = None
        self.locations = {}

    def encode_multipart_formdata(self, fields):
        boundary = binascii.hexlify(os.urandom(16)).decode('ascii')

        body = (
                "".join("--%s\r\n"
                        "Content-Disposition: form-data; name=\"%s\"\r\n"
                        "\r\n"
                        "%s\r\n" % (boundary, field, value)
                        for field, value in fields.items()) +
                "--%s--\r\n" % boundary
        )

        content_type = "multipart/form-data; boundary=%s" % boundary

        return body, content_type

    async def dash_login(self, company_id, username, password):
        url = f'{base_url}?Action=Auth/validateLogin.json&extension=json&company='
        data, content_type = self.encode_multipart_formdata({'_method': 'POST', 'company_code': company_id,
                                                             'username': username, 'password': password})

        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': content_type,
            'DNT': '1',
            'Origin': 'https://apps.daysmartrecreation.com',
            'Referer': 'https://apps.daysmartrecreation.com/dash/admin/index.php?&ver=8.0',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.post(url, headers=headers, data=data) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        json_content = json.loads(content)
                        success = json_content.get('success')
                        if success == 'Login Successful':
                            return True

                        err_msg = json_content.get('messages')
                        if err_msg:
                            return err_msg
                except:
                    loop_count += 1
                    await asyncio.sleep(2)

            return False

    async def get_locations(self):
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'DNT': '1',
            'Referer': 'https://apps.daysmartrecreation.com/dash/admin/index.php?Action=Auth/login&company=canlan',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': USER_AGENT,
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        params = {
            'Action': 'Cashregister',
        }

        async with self.sema:
            async with self.session.get(f'{base_url}', headers=headers, params=params) as request:
                response = await request.content.read()
                content = response.decode('utf-8')
                html_content = bs(content, 'html.parser')
                location_tag = html_content.find('select', attrs={'name': 'NewFacilityID'}).find_all('option')
                for location in location_tag:
                    self.locations[location.text] = location.get('value')
                return self.locations

    async def dash_all_search(self):
        search_url = f'{base_url}?Action=all/search.json'
        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'DNT': '1',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.get(search_url, headers=headers) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        json_content = json.loads(content)
                        self.all_searches = json_content.get('data')
                        return self.all_searches
                except:
                    loop_count += 1
                    await asyncio.sleep(2)

            return False

    async def jasper_login(self):
        jasper_url = 'https://jasper.daysmartrecreation.com/jasperserver-pro/JavaScriptServlet'
        headers = {
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'DNT': '1',
            'FETCH-CSRF-TOKEN': '1',
            'Origin': 'https://jasper.daysmartrecreation.com',
            'Referer': 'https://jasper.daysmartrecreation.com/jasperserver-pro/login.html',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.post(jasper_url, headers=headers) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        self.csrf_code = str(content).replace("OWASP_CSRFTOKEN:", "").strip()
                        return True
                except:
                    loop_count += 1
                    await asyncio.sleep(2)
            return False

    async def report_view(self, report_url):
        report_view_url = f'{base_url}{report_url}'
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'DNT': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': USER_AGENT,
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.get(report_view_url, headers=headers) as request:
                        response = await request.content.read()
                        self.report_referer_url = str(request.url)
                        content = response.decode('utf-8')
                        report_html = bs(content, "html.parser")
                        title = report_html.find('title').text
                        if title == 'Jasper Report - DaySmart Rec Admin':
                            return True
                        else:
                            return False
                except:
                    loop_count += 0
                    await asyncio.sleep(2)
            return False

    async def load_report(self, report, report_url_):
        report_action = str(report_url_).replace("/view", "")
        report_unit = str(report_action).replace('?Action=Jasper/report', '').replace('-', '/')
        report_unit = urllib.parse.quote_plus(report_unit)
        decorate = 'decorate=no'
        theme = 'theme=default'

        report_url = f'{base_url}{report_action}&{decorate}&{theme}&report_unit={report_unit}'

        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'DNT': '1',
            'Referer': str(self.report_referer_url),
            'Sec-Fetch-Dest': 'iframe',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': USER_AGENT,
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        self.input_control_params = None

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.get(report_url, headers=headers) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        report_html = bs(content, "html.parser")
                        title = report_html.find('title').text
                        if str(report) in title:
                            self.referer_url = str(request.real_url)
                            self.input_control_params = dict(request._url.query)
                            return report_html
                        loop_count += 1
                except:
                    loop_count += 1
                    await asyncio.sleep(2)

            return False

    async def input_controls(self, report_url_):
        report_unit = str(report_url_).replace('?Action=Jasper/report', '').replace('-', '/').replace('/view', '')
        input_url = f'https://jasper.daysmartrecreation.com/jasperserver-pro/rest_v2/reports{report_unit}/inputControls/'

        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US',
            'Connection': 'keep-alive',
            'DNT': '1',
            'OWASP_CSRFTOKEN': self.csrf_code,
            'Origin': 'https://jasper.daysmartrecreation.com',
            'Referer': str(self.referer_url),
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest, OWASP CSRFGuard Project',
            'X-Suppress-Basic': 'true',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        json_data = {}
        for input_key, input_value in self.input_control_params.items():
            json_data[input_key] = [input_value]

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.post(input_url, headers=headers, json=json_data) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        input_control = json.loads(content).get('inputControl')
                        return input_control
                except:
                    loop_count += 1
                    await asyncio.sleep(2)

            return False

    async def report_values(self, report_url_, input_control):
        all_items = [item.get('id') for item in input_control]
        report_items = ';'.join(all_items)

        report_unit = str(report_url_).replace('?Action=Jasper/report', '').replace('-', '/').replace('/view', '')

        url = f'https://jasper.daysmartrecreation.com/jasperserver-pro/rest_v2/reports{report_unit}/inputControls/' \
              f'{report_items}/values'

        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US',
            'Connection': 'keep-alive',
            'DNT': '1',
            'OWASP_CSRFTOKEN': self.csrf_code,
            'Origin': 'https://jasper.daysmartrecreation.com',
            'Referer': str(self.referer_url),
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest, OWASP CSRFGuard Project',
            'X-Suppress-Basic': 'true',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        json_data = {}
        for input_key, input_value in self.input_control_params.items():
            json_data[input_key] = [input_value]

        loop_count = 0
        async with self.sema:
            while loop_count < 3:
                try:
                    async with self.session.post(url, headers=headers, json=json_data) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        return True
                except:
                    loop_count += 1
                    await asyncio.sleep(2)
            return False

    async def report_flow(self, report, report_num, report_date, other_month):
        flow_url = 'https://jasper.daysmartrecreation.com/jasperserver-pro/flow.html'
        headers = {
            'Accept': 'text/html, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'DNT': '1',
            'OWASP_CSRFTOKEN': self.csrf_code,
            'Origin': 'https://jasper.daysmartrecreation.com',
            'Referer': str(self.referer_url),
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': USER_AGENT,
            'X-Requested-With': 'XMLHttpRequest, OWASP CSRFGuard Project',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Microsoft Edge";v="100"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        params = {
            '_flowExecutionKey': f'e{str(report_num)}s1',
            '_flowId': 'viewReportFlow',
            '_eventId': 'refreshReport',
            'pageIndex': '0',
            'decorate': 'no',
            'confirm': 'true',
            'decorator': 'empty',
            'ajax': 'true',
        }

        data = self.get_report_data(report, report_date, other_month)

        loop_count = 0
        async with self.sema:
            while loop_count < 2:
                try:
                    async with self.session.post(flow_url, headers=headers, params=params, data=data) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        report_data_html = bs(content, 'html.parser')
                        span = report_data_html.find('span').text
                        if str(span).strip().replace(" ", "") == str(report).strip().replace(" ", ""):
                            return report_data_html
                        elif str(report).strip() in report_data_html.text:
                            return report_data_html
                        return False
                except:
                    loop_count += 1
                    await asyncio.sleep(2)

            return False

    def get_report_data(self, report, report_date, other_month):
        if "class" in str(report).lower().strip():
            today_ = datetime.datetime.today().strftime("%Y-%m-%d")
            data = {
                'ParaMonth': str(report_date),
                'Reportview': str(other_month).lower().strip(),
                'LocationID': '~NOTHING~',
                'Season': '~NOTHING~',
                'startDate': str(today_),
                'ShowSeasonDetails': str(other_month).lower().strip(),
            }
            return data

        elif "camp" in str(report).lower().strip() or "team" in str(report).lower().strip():
            data = {
                'startDate': str(report_date),
                'ShowSeasonDetails': str(other_month),
                'LocationID': '~NOTHING~',
                'Season': '~NOTHING~',
            }
            return data

        elif "membership" in str(report).lower().strip():
            data = {
                'ParaAsof': str(report_date),
                'Facility': '1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21',
                'Month': '~NOTHING~',
                'Year': '~NOTHING~',
            }
            return data

        elif "event" in str(report).lower().strip():
            data = {
                'AsofDate': str(report_date),
                'LocationID': '~NOTHING~',
                'EventType': '~NOTHING~',
                'GLCode': '~NOTHING~',
                'ShowEventDetail': 'true',
                'ShowEventWithoutInvoice': 'false',
            }
            return data
