import re
import os
import json
import psycopg2
import traceback
from pathlib import Path
from pprint import pprint
from datetime import datetime
from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials


def execution_time_json(file_name="execution_time.json"):
    def decorator(func):
        def wrapper(*args, **kwargs):
            time_file = f"{Path(__file__).resolve().parent}/{file_name}"
            last_execution_time = None

            # 기존 실행 시간 읽기
            if os.path.exists(time_file):
                with open(time_file, "r") as f:
                    data = json.load(f)
                    last_execution_time = data.get("last_execution_time")
            else:
                print("File does not exist. Creating new file with current time.")
                last_execution_time = datetime.now().isoformat()
                with open(time_file, "w") as f:
                    json.dump({"last_execution_time": last_execution_time}, f)

            # 함수 실행
            result = None
            try:
                result = func(*args, last_execution_time=last_execution_time, **kwargs)
            except Exception as e:
                print(f"Error in decorated function {func.__name__}: {e}")
                print(traceback.format_exc())

            # 새 실행 시간 저장
            # new_execution_time = datetime.now().replace(microsecond=0).isoformat()
            # with open(time_file, "w") as f:
            #    json.dump({"last_execution_time": new_execution_time}, f)

            return result

        return wrapper

    return decorator


class Holiwork_manager:
    def __init__(self):
        self.host = "172.20.2.30"
        self.port = 5432
        self.db = "do"
        self.user = "daoudev"
        self.pw = "Vhtm!@1357"
        self.conn = psycopg2.connect(
            host=self.host,
            port=self.port,
            database=self.db,
            user=self.user,
            password=self.pw,
        )
        self.SERVICE_ACCOUNT_FILE = f"{Path(__file__).resolve().parent}/sheet-key.json"
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        self.result_list = []
        self.cur = self.conn.cursor()

        json_file_path = Path(__file__).resolve().parent / "sheet-id.json"
        with open(json_file_path, "r", encoding="utf-8") as file:
            data = json.load(file)
        self.SPREADSHEET_ID = data.get("sheetid")

    def holiday_work(self, doc_body_id):

        data = self.html_parse_from_doc_id(doc_body_id, 0)
        # 필요한 데이터 추출
        try:
            name = data[0].strip()
            department = data[2]
            form = re.search(r"\[(.*?)\]", data[6]).group(1)
            dates = re.findall(r"\d{4}-\d{2}-\d{2}", data[5])
            remaining_vacation = None
            application_days = 1
            result = {
                "이름": name,
                "부서": department,  # 부서 정보는 'SUPERVISOR'로 지정
                "휴가 종류": form,
                "시작일": dates[0],
                "종료일": dates[0],
                "잔여연차": remaining_vacation,
                "사용일수": application_days,
            }
            self.result_list.append(list(result.values()))
        except Exception as e:
            print("데이터 추출 중 오류 holiday_work:", e)
            print(f"Error: {e}")
            print(traceback.format_exc())

    def alter_holiday(self, doc_body_id):

        data = self.html_parse_from_doc_id(doc_body_id, 1)
        # html에서 <td/>기준으로 필요한 데이터 data추출
        try:
            # print("data::", data)
            name = data[1].strip()  # 기안자 이름
            vacation_type = data[7].strip()  # 휴가 종류 origin 7
            vacation_date_range = data[8].strip()  # 휴가 날짜 및 사용일수 origin 8
            dates = re.findall(
                r"\d{4}-\d{2}-\d{2}", vacation_date_range
            )  # 날짜(yyyy-mm-dd) 추출
            vacation_days = re.findall(
                r"\d+\.?\d*", vacation_date_range.split("사용일수 :")[1]
            )  # 사용일수 숫자 추출
            remaining_vacation_match = re.search(r"잔여연차 :\s*(\d+)", data[10])
            if remaining_vacation_match:
                remaining_vacation = float(remaining_vacation_match.group(1))
            else:
                remaining_vacation = "기재안함"
            # 날짜 파싱 (휴가 날짜)
            date_range = re.search(
                r"(\d{4}-\d{2}-\d{2}) ~ (\d{4}-\d{2}-\d{2})", vacation_date_range
            )

            result = {
                "이름": name,
                "부서": data[3],  # 부서 정보는 'SUPERVISOR'로 지정
                "휴가 종류": vacation_type,
                "시작일": dates[0],
                "종료일": dates[1],
                "잔여연차": remaining_vacation,
                "사용일수": -abs(float(vacation_days[0])),
            }
            self.result_list.append(list(result.values()))
        except Exception as e:
            print("데이터 추출 중 오류 alter_holiday:", e)
            print(f"Error: {e}")
            print(traceback.format_exc())

    @execution_time_json()
    def get_elec_payment(self, last_execution_time=None):
        try:
            query = """
               SELECT 
                   drafter_name, 
                   doc_body_id, 
                   form_alias, 
                   doc_status, 
                   appr_status, 
                   title, 
                   created_at, 
                   updated_at, 
                   doc_status
               FROM 
                   go_appr_documents
               WHERE 
                   form_alias IN ('휴가계', '휴일근무신청서')
                   AND doc_status != 'CREATE'
                   AND (
                       title LIKE '%[대체휴무]%' 
                       OR title LIKE '%[휴일근무]%'
                   )
            """

            # 실행 시간 조건 추가
            if last_execution_time:
                query += f" AND updated_at > '{last_execution_time}'"

            # PostgreSQL 쿼리 실행
            self.cur.execute(query)
            rows = self.cur.fetchall()

            # 결과 처리
            for idx, row in enumerate(rows):
                if idx >= 30000000000:
                    break
                doc_body_id = row[1]
                form_distinction = row[5]
                doc_status = row[8]

                if doc_status == "COMPLETE":
                    if form_distinction.startswith("[휴일근무]"):
                        self.holiday_work(doc_body_id)
                    elif form_distinction.startswith("[대체휴무]"):
                        self.alter_holiday(doc_body_id)
                else:
                    print(f"Skipping document {doc_body_id} with status: {doc_status}")

            self.cur.close()
            self.conn.close()

            return self.result_list

        except Exception as e:
            print("데이터 추출 중 오류 get_elec_payment:", e)
            print(traceback.format_exc())

    def first_sheet(self, data, sheet_titles, service):
        try:
            sheet_name = "시트1"
            if sheet_name in sheet_titles:
                # 데이터 요청 생성
                body = {"values": data}

                # 데이터 삽입
                result = (
                    service.spreadsheets()
                    .values()
                    .append(
                        spreadsheetId=self.SPREADSHEET_ID,
                        range=f"{sheet_name}!A2",
                        valueInputOption="RAW",
                        body=body,
                        insertDataOption="INSERT_ROWS",
                    )
                    .execute()
                )

                print(f"{result.get('updates').get('updatedCells')} cells updated.")
            else:
                print(f"Sheet named '{sheet_name}' does not exist.")

        except Exception as e:
            print("An error occurred first_sheet:", e)
            print(f"Error: {e}")
            print(traceback.format_exc())

    def personal_sheet(self, data, sheet_titles, service):
        try:
            requests = []  # batchUpdate 요청 리스트

            # 시트 생성 요청
            for row in data:
                sheet_user_name = row[0]  # 시트 이름

                if sheet_user_name not in sheet_titles:
                    print(
                        f"Sheet named '{sheet_user_name}' does not exist. Creating new sheet."
                    )
                    requests.append(
                        {"addSheet": {"properties": {"title": sheet_user_name}}}
                    )
                    sheet_titles.append(sheet_user_name)

            # 배치 업데이트로 시트 생성
            if requests:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=self.SPREADSHEET_ID, body={"requests": requests}
                ).execute()
                print(f"{len(requests)} sheets created successfully.")
                requests = []

            # 시트 ID 매핑
            spreadsheet = (
                service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
            )
            sheet_id_map = {
                sheet["properties"]["title"]: sheet["properties"]["sheetId"]
                for sheet in spreadsheet["sheets"]
            }

            # 데이터 그룹화
            grouped_data = {}
            for row in data:
                sheet_user_name = row[0]
                sheet_data = row[2:]

                if sheet_user_name not in grouped_data:
                    grouped_data[sheet_user_name] = []

                grouped_data[sheet_user_name].append(sheet_data)

            # 각 사용자 시트에 데이터 추가 요청 생성
            for sheet_user_name, sheet_rows in grouped_data.items():
                rows = []
                for sheet_data in sheet_rows:
                    row_values = []
                    for cell in sheet_data:
                        if isinstance(cell, (int, float)):  # 숫자인 경우
                            row_values.append(
                                {"userEnteredValue": {"numberValue": cell}}
                            )
                        else:  # 문자열 또는 기타 타입인 경우
                            row_values.append(
                                {"userEnteredValue": {"stringValue": str(cell)}}
                            )
                    rows.append({"values": row_values})

                if sheet_user_name in sheet_id_map:
                    sheet_id = sheet_id_map[sheet_user_name]
                    requests.append(
                        {
                            "updateCells": {
                                "rows": rows,
                                "fields": "userEnteredValue",
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 2,  # A3부터 시작 인덱스
                                    "startColumnIndex": 0,
                                    "endRowIndex": 2
                                    + len(rows),  # 데이터 크기에 맞게 종료 행 계산
                                    "endColumnIndex": max(
                                        len(row) for row in sheet_rows
                                    ),
                                },
                            }
                        }
                    )

                    header_row = [
                        {"userEnteredValue": {"stringValue": "휴가 종류"}},
                        {"userEnteredValue": {"stringValue": "시작날"}},
                        {"userEnteredValue": {"stringValue": "종료날"}},
                        {"userEnteredValue": {"stringValue": "잔여 휴가"}},
                        {"userEnteredValue": {"stringValue": "사용 일수"}},
                    ]

                    requests.append(
                        {
                            "updateCells": {
                                "rows": [{"values": header_row}],
                                "fields": "userEnteredValue",
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,  # A1부터 시작
                                    "endRowIndex": 1,  # 한 행만 업데이트
                                    "startColumnIndex": 0,
                                    "endColumnIndex": len(header_row),
                                },
                            }
                        }
                    )

            # 배치 업데이트로 데이터 삽입
            if requests:
                response = (
                    service.spreadsheets()
                    .batchUpdate(
                        spreadsheetId=self.SPREADSHEET_ID, body={"requests": requests}
                    )
                    .execute()
                )
                print(f"{len(requests)} updateCells requests processed successfully.")

        except Exception as e:
            print("An error occurred in personal_sheet with batchUpdate:")
            print(f"Error: {e}")
            print(traceback.format_exc())

    def alter_holiday_count_sheet(self, data, sheet_titles, service):
        try:
            sheet_name = "잔여대체휴가"

            # "잔여대체휴가" 시트의 데이터 가져오기
            if sheet_name not in sheet_titles:
                print(f"Sheet '{sheet_name}' does not exist. MAKE '{sheet_name}'")
                return

            # '잔여대체휴가' 시트의 A열 값 확인 (이름 목록 확인)
            range_ = f"{sheet_name}!A:A"
            result = (
                service.spreadsheets()
                .values()
                .get(spreadsheetId=self.SPREADSHEET_ID, range=range_)
                .execute()
            )
            existing_names = []

            if result.get("values", []):  # A열의 기존 데이터 가져오기
                for row in result["values"]:
                    for value in row:
                        if value:
                            existing_names.append(value)
                            break

            # 업데이트 요청 생성
            requests = []

            for row in data:
                person_name = row[0]
                person_depart = row[1]
                sum_formula = f"=SUM('{person_name}'!E:E)"

                # 이미 존재하는 이름은 건너뛰기
                if person_name in existing_names:
                    print(
                        f"Name '{person_name}' already exists in the sheet '{sheet_name}'."
                    )
                    continue

                # 추가할 데이터 행
                next_empty_row = len(existing_names) + 1
                requests.append(
                    {
                        "updateCells": {
                            "rows": [
                                {
                                    "values": [
                                        {
                                            "userEnteredValue": {
                                                "stringValue": person_name
                                            }
                                        },  # A열
                                        {
                                            "userEnteredValue": {
                                                "stringValue": person_depart
                                            }
                                        },  # B열
                                        {
                                            "userEnteredValue": {
                                                "formulaValue": sum_formula
                                            }
                                        },  # C열
                                    ]
                                }
                            ],
                            "fields": "userEnteredValue",
                            "range": {
                                "sheetId": None,  # 배치 실행 전에 설정
                                "startRowIndex": next_empty_row,  # 행 인덱스 (0부터 시작)
                                "endRowIndex": next_empty_row + 1,
                                "startColumnIndex": 0,
                                "endColumnIndex": 3,
                            },
                        }
                    }
                )
                existing_names.append(person_name)  # 이름 추가하여 중복 방지

            # 시트 ID 매핑
            spreadsheet = (
                service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
            )
            sheet_id_map = {
                sheet["properties"]["title"]: sheet["properties"]["sheetId"]
                for sheet in spreadsheet["sheets"]
            }
            sheet_id = sheet_id_map.get(sheet_name)

            if sheet_id is None:
                print(f"Sheet ID for '{sheet_name}' not found.")
                return

            # sheetId를 모든 요청에 설정
            for request in requests:
                if "updateCells" in request:
                    request["updateCells"]["range"]["sheetId"] = sheet_id

            # 배치 업데이트 실행
            if requests:
                response = (
                    service.spreadsheets()
                    .batchUpdate(
                        spreadsheetId=self.SPREADSHEET_ID, body={"requests": requests}
                    )
                    .execute()
                )
                print(f"{len(requests)} rows updated in the sheet '{sheet_name}'.")

            # 업데이트 후 정렬 요청
            self.sort_sheet(sheet_name, 1, service)  # B열(1-indexed)을 기준으로 정렬

        except Exception as e:
            print(f"An error occurred alter_holiday_count_sheet: {e}")
            print(f"Error: {e}")
            print(traceback.format_exc())

    def sort_sheet(self, sheet_name, sort_column_index, service):

        try:
            # 시트 ID 가져오기
            sheet_metadata = (
                service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
            )
            sheets = sheet_metadata.get("sheets", [])
            sheet_id = None
            for sheet in sheets:
                if sheet.get("properties", {}).get("title") == sheet_name:
                    sheet_id = sheet.get("properties", {}).get("sheetId")
                    break

            if sheet_id is None:
                print(f"Sheet '{sheet_name}' not found.")
                return

            # 정렬 요청 생성
            request_body = {
                "requests": [
                    {
                        "sortRange": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 1,  # 데이터가 시작되는 행 (헤더 제외)
                                "startColumnIndex": 0,  # A열 (0부터 시작)
                                "endColumnIndex": 3,  # C열까지 포함
                            },
                            "sortSpecs": [
                                {
                                    "dimensionIndex": sort_column_index,  # 정렬 기준 열 (0부터 시작)
                                    "sortOrder": "ASCENDING",  # 오름차순 정렬
                                }
                            ],
                        }
                    }
                ]
            }

            # API 호출
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.SPREADSHEET_ID, body=request_body
            ).execute()

            print(
                f"Sheet '{sheet_name}' sorted successfully by column {sort_column_index + 1}."
            )
            print("+" * 100)
            print("+" * 45 + " FINISH " + "+" * 47)
            print("+" * 100)

        except Exception as e:
            print(f"An error occurred in sort_sheet: {e}")
            print(f"Error: {e}")
            print(traceback.format_exc())

    def append_data_sheet(self, data):
        credentials = Credentials.from_service_account_file(
            self.SERVICE_ACCOUNT_FILE, scopes=self.SCOPES
        )
        service = build("sheets", "v4", credentials=credentials)

        # 스프레드시트 정보 가져오기
        spreadsheet = (
            service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
        )

        # 시트 이름 목록 가져오기
        sheets = spreadsheet["sheets"]
        sheet_titles = [sheet["properties"]["title"] for sheet in sheets]

        try:
            for i, row in enumerate(data):  # enumerate를 사용해 인덱스 추적
                if i >= 1:
                    break
                self.first_sheet(data, sheet_titles, service)
                self.personal_sheet(data, sheet_titles, service)
                self.alter_holiday_count_sheet(data, sheet_titles, service)

        except Exception as e:
            print(f"An error occurred at append_data_sheet: {e}")
            print(f"Error: {e}")
            print(traceback.format_exc())

    def html_parse_from_doc_id(self, doc_body_id, rows_index):
        self.cur.execute(
            f"SELECT contents FROM go_appr_doc_bodies WHERE id = {doc_body_id};"
        )
        html_row = self.cur.fetchone()
        if html_row:
            html_content = html_row[0]
            soup = BeautifulSoup(html_content, "html.parser")
            rows = soup.find_all("tr")
            data = [
                row.find_all("td")[1].text.strip() for row in rows[int(rows_index) :]
            ]

        return data


if __name__ == "__main__":
    hm = Holiwork_manager()
    result_data = hm.get_elec_payment()  # 전자결재 데이터를 가져옴
    hm.append_data_sheet(result_data)  # Google Sheets에 데이터 추가
    print("DONE")
