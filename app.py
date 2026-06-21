import sqlite3
from copy import copy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# ==================================================
# 1. 기본 설정
# ==================================================

DB_PATH = "steamer.db"
TEMPLATE_PATH = "steamer_template.xlsx"

KST = ZoneInfo("Asia/Seoul")

MACHINES = [
    "11", "12", "13", "14",
    "21", "22", "23", "24",
    "51", "52", "53", "54",
]

# 11호기, 51호기만 유량 + 압력 2줄 입력
FLOW_PRESSURE_MACHINES = ["11", "51"]

# 현재 템플릿 기준 행
MACHINE_ROW_MAP = {
    "11": 11,
    "12": 12,
    "13": 13,
    "14": 14,
    "21": 15,
    "22": 16,
    "23": 17,
    "24": 18,
    "51": 19,
    "52": 20,
    "53": 21,
    "54": 22,
}


# ==================================================
# 2. DB 함수
# ==================================================

def connect_db():
    return sqlite3.connect(DB_PATH)


def create_table():
    with connect_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS steamer_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,

                check_date TEXT NOT NULL,
                check_time TEXT NOT NULL,
                machine TEXT NOT NULL,

                product TEXT,
                heat_temp REAL,
                oil_set_temp REAL,
                oil_now_temp REAL,
                steam_usage REAL,
                ct_water REAL,

                inlet_flow REAL,
                inlet_pressure REAL,

                middle_flow REAL,
                middle_pressure REAL,

                outlet_flow REAL,
                outlet_pressure REAL,

                memo TEXT,

                created_at TEXT,
                updated_at TEXT,

                UNIQUE(check_date, machine)
            )
            """
        )

        existing_columns = [
            row[1]
            for row in conn.execute(
                "PRAGMA table_info(steamer_logs)"
            ).fetchall()
        ]

        new_columns = {
            "product": "TEXT",
            "heat_temp": "REAL",
            "oil_set_temp": "REAL",
            "oil_now_temp": "REAL",
            "steam_usage": "REAL",
            "ct_water": "REAL",
            "inlet_flow": "REAL",
            "inlet_pressure": "REAL",
            "middle_flow": "REAL",
            "middle_pressure": "REAL",
            "outlet_flow": "REAL",
            "outlet_pressure": "REAL",
            "memo": "TEXT",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        }

        for column_name, column_type in new_columns.items():
            if column_name not in existing_columns:
                conn.execute(
                    f"ALTER TABLE steamer_logs "
                    f"ADD COLUMN {column_name} {column_type}"
                )

        now_text = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")

        conn.execute(
            """
            UPDATE steamer_logs
            SET created_at = ?
            WHERE created_at IS NULL
               OR created_at = ''
            """,
            (now_text,),
        )

        conn.execute(
            """
            UPDATE steamer_logs
            SET updated_at = ?
            WHERE updated_at IS NULL
               OR updated_at = ''
            """,
            (now_text,),
        )


def save_or_update_record(record):
    check_date = record[0]
    machine = record[2]

    with connect_db() as conn:
        existing_record = conn.execute(
            """
            SELECT id
            FROM steamer_logs
            WHERE check_date = ?
              AND machine = ?
            """,
            (check_date, machine),
        ).fetchone()

        if existing_record is None:
            conn.execute(
                """
                INSERT INTO steamer_logs (
                    check_date,
                    check_time,
                    machine,
                    product,
                    heat_temp,
                    oil_set_temp,
                    oil_now_temp,
                    steam_usage,
                    ct_water,
                    inlet_flow,
                    inlet_pressure,
                    middle_flow,
                    middle_pressure,
                    outlet_flow,
                    outlet_pressure,
                    memo,
                    created_at,
                    updated_at
                )
                VALUES (
                    ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?
                )
                """,
                record,
            )

            return "inserted"

        conn.execute(
            """
            UPDATE steamer_logs
            SET
                check_time = ?,
                product = ?,
                heat_temp = ?,
                oil_set_temp = ?,
                oil_now_temp = ?,
                steam_usage = ?,
                ct_water = ?,
                inlet_flow = ?,
                inlet_pressure = ?,
                middle_flow = ?,
                middle_pressure = ?,
                outlet_flow = ?,
                outlet_pressure = ?,
                memo = ?,
                updated_at = ?
            WHERE check_date = ?
              AND machine = ?
            """,
            (
                record[1],
                record[3],
                record[4],
                record[5],
                record[6],
                record[7],
                record[8],
                record[9],
                record[10],
                record[11],
                record[12],
                record[13],
                record[14],
                record[15],
                record[17],
                record[0],
                record[2],
            ),
        )

        return "updated"


def load_records(selected_date):
    date_text = selected_date.isoformat()

    conn = connect_db()
    conn.row_factory = sqlite3.Row

    records = conn.execute(
        """
        SELECT
            id,
            check_date,
            check_time,
            machine,
            product,
            heat_temp,
            oil_set_temp,
            oil_now_temp,
            steam_usage,
            ct_water,
            inlet_flow,
            inlet_pressure,
            middle_flow,
            middle_pressure,
            outlet_flow,
            outlet_pressure,
            memo,
            created_at,
            updated_at
        FROM steamer_logs
        WHERE check_date = ?
        ORDER BY CAST(machine AS INTEGER) ASC
        """,
        (date_text,),
    ).fetchall()

    conn.close()

    return records


def delete_record(record_id):
    with connect_db() as conn:
        cursor = conn.execute(
            """
            DELETE FROM steamer_logs
            WHERE id = ?
            """,
            (record_id,),
        )

        return cursor.rowcount


# ==================================================
# 3. 입력값 처리 함수
# ==================================================

def parse_number(value, label):
    value = value.strip()

    if value == "":
        return None

    try:
        number = float(value)
    except ValueError:
        raise ValueError(f"{label}은 숫자만 입력해야 합니다.")

    if "." in value:
        decimal_part = value.split(".")[1]

        if len(decimal_part) > 5:
            raise ValueError(
                f"{label}은 소수점 아래 5자리까지만 입력 가능합니다."
            )

    return number


def format_number(value):
    if value is None:
        return ""

    if isinstance(value, float):
        return f"{value:.5f}".rstrip("0").rstrip(".")

    return str(value)


def excel_value(value, dot_if_empty=False):
    if value is None:
        if dot_if_empty:
            return "."
        return None

    return value


def number_text_input(label, key, machine):
    """
    호기별로 입력창 key를 분리한다.
    예: 11_heat_temp, 12_heat_temp
    """

    return st.text_input(
        label,
        key=f"{machine}_{key}",
        placeholder="예: 123.45",
    )


# ==================================================
# 4. QR 주소에서 호기 읽기
# ==================================================

def get_machine_from_url():
    machine = st.query_params.get("machine")

    if machine in MACHINES:
        return machine

    return None


# ==================================================
# 5. Excel 보조 함수
# ==================================================

def get_writable_cell_address(worksheet, cell_address):
    """
    병합셀 내부 좌표가 들어오면 실제로 값을 쓸 수 있는
    병합 범위의 왼쪽 위 셀 주소를 반환한다.
    """

    for merged_range in worksheet.merged_cells.ranges:
        if cell_address in merged_range:
            return merged_range.start_cell.coordinate

    return cell_address


def write_cell_value(worksheet, cell_address, value):
    """
    병합셀 오류 방지용 값 입력 함수.
    """

    writable_cell_address = get_writable_cell_address(
        worksheet,
        cell_address,
    )

    worksheet[writable_cell_address] = value


def make_flow_pressure_text(flow, pressure):
    flow_text = format_number(flow)
    pressure_text = format_number(pressure)

    if flow_text == "" and pressure_text == "":
        return ""

    return f"{flow_text}\n{pressure_text}"


def write_multiline_cell(worksheet, cell_address, value):
    """
    유량/압력 두 줄 입력.
    병합셀에도 안전하게 입력.
    """

    writable_cell_address = get_writable_cell_address(
        worksheet,
        cell_address,
    )

    cell = worksheet[writable_cell_address]
    cell.value = value

    old_alignment = copy(cell.alignment)

    cell.alignment = Alignment(
        horizontal=old_alignment.horizontal,
        vertical=old_alignment.vertical,
        text_rotation=old_alignment.text_rotation,
        wrap_text=True,
        shrink_to_fit=old_alignment.shrink_to_fit,
        indent=old_alignment.indent,
    )


# ==================================================
# 6. Excel 생성 함수
# ==================================================

def make_template_excel(selected_date):
    """
    현재 steamer_template.xlsx 기준 좌표.

    날짜: A6
    점검시간: B열
    제품명: C열
    열교환기 온도: D열
    유탕온도 설정: E열
    유탕온도 현재: F열
    증기 사용량: G열
    C/T수: H열
    입구: K열
    중간: M열
    출구: O열
    """

    if not Path(TEMPLATE_PATH).exists():
        raise FileNotFoundError(
            f"{TEMPLATE_PATH} 파일을 찾을 수 없습니다. "
            "GitHub 저장소에 steamer_template.xlsx 파일을 올렸는지 확인하세요."
        )

    records = load_records(selected_date)

    workbook = load_workbook(TEMPLATE_PATH)

    if "Sheet1" in workbook.sheetnames:
        worksheet = workbook["Sheet1"]
    else:
        worksheet = workbook.worksheets[0]

    weekday_names = [
        "월", "화", "수", "목", "금", "토", "일"
    ]

    weekday = weekday_names[selected_date.weekday()]

    write_cell_value(
        worksheet,
        "A6",
        (
            f"  {selected_date.year} 년      "
            f"{selected_date.month} 월        "
            f"{selected_date.day} 일         "
            f"{weekday}요일"
        ),
    )

    for record in records:
        machine = record["machine"]

        if machine not in MACHINE_ROW_MAP:
            continue

        row = MACHINE_ROW_MAP[machine]

        write_cell_value(worksheet, f"B{row}", record["check_time"])
        write_cell_value(worksheet, f"C{row}", record["product"])

        write_cell_value(
            worksheet,
            f"D{row}",
            excel_value(record["heat_temp"]),
        )

        write_cell_value(
            worksheet,
            f"E{row}",
            excel_value(record["oil_set_temp"]),
        )

        write_cell_value(
            worksheet,
            f"F{row}",
            excel_value(record["oil_now_temp"]),
        )

        # 증기 사용량은 값이 없으면 "." 입력
        write_cell_value(
            worksheet,
            f"G{row}",
            excel_value(record["steam_usage"], dot_if_empty=True),
        )

        write_cell_value(
            worksheet,
            f"H{row}",
            excel_value(record["ct_water"]),
        )

        if machine in FLOW_PRESSURE_MACHINES:
            write_multiline_cell(
                worksheet,
                f"K{row}",
                make_flow_pressure_text(
                    record["inlet_flow"],
                    record["inlet_pressure"],
                ),
            )

            write_multiline_cell(
                worksheet,
                f"M{row}",
                make_flow_pressure_text(
                    record["middle_flow"],
                    record["middle_pressure"],
                ),
            )

            write_multiline_cell(
                worksheet,
                f"O{row}",
                make_flow_pressure_text(
                    record["outlet_flow"],
                    record["outlet_pressure"],
                ),
            )

        else:
            write_cell_value(
                worksheet,
                f"K{row}",
                excel_value(record["inlet_pressure"]),
            )

            write_cell_value(
                worksheet,
                f"M{row}",
                excel_value(record["middle_pressure"]),
            )

            write_cell_value(
                worksheet,
                f"O{row}",
                excel_value(record["outlet_pressure"]),
            )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue(), len(records)


# ==================================================
# 7. 앱 초기 설정
# ==================================================

st.set_page_config(
    page_title="증숙기 점검일지",
    page_icon="📝",
    layout="wide",
)

create_table()

qr_machine = get_machine_from_url()

if qr_machine is not None:
    default_machine_index = MACHINES.index(qr_machine)
else:
    default_machine_index = 0


# ==================================================
# 8. 화면 구성
# ==================================================

st.title("증숙기 점검일지")

tab_input, tab_records, tab_manage, tab_excel = st.tabs(
    [
        "점검 입력",
        "저장 기록",
        "기록 관리",
        "Excel 다운로드",
    ]
)


# ==================================================
# 9. 점검 입력 탭
# ==================================================

with tab_input:

    st.subheader("증숙기 점검값 입력")

    now = datetime.now(KST)

    st.info(
        f"점검 날짜/시간은 저장 버튼을 누르는 현재 시각으로 자동 저장됩니다. "
        f"현재 기준: {now.strftime('%Y-%m-%d %H:%M')}"
    )

    if qr_machine is not None:
        st.success(
            f"QR을 통해 {qr_machine}호기가 자동 선택되었습니다."
        )
    else:
        st.info(
            "QR 접속이 아닙니다. 호기를 직접 선택해 주세요."
        )

    machine = st.selectbox(
        "호기 선택",
        MACHINES,
        index=default_machine_index,
    )

    product = st.text_input(
        "제품명",
        key=f"{machine}_product",
    )

    heat_temp_text = number_text_input(
        "열교환기 온도",
        "heat_temp",
        machine,
    )

    oil_set_temp_text = number_text_input(
        "유탕온도 설정",
        "oil_set_temp",
        machine,
    )

    oil_now_temp_text = number_text_input(
        "유탕온도 현재",
        "oil_now_temp",
        machine,
    )

    steam_usage_text = number_text_input(
        "증기 사용량",
        "steam_usage",
        machine,
    )

    ct_water_text = number_text_input(
        "C/T수",
        "ct_water",
        machine,
    )

    if machine in FLOW_PRESSURE_MACHINES:
        st.markdown("#### 증기 입구 / 중간 / 출구")

        inlet_flow_text = number_text_input(
            "입구 유량",
            "inlet_flow",
            machine,
        )

        inlet_pressure_text = number_text_input(
            "입구 압력",
            "inlet_pressure",
            machine,
        )

        middle_flow_text = number_text_input(
            "중간 유량",
            "middle_flow",
            machine,
        )

        middle_pressure_text = number_text_input(
            "중간 압력",
            "middle_pressure",
            machine,
        )

        outlet_flow_text = number_text_input(
            "출구 유량",
            "outlet_flow",
            machine,
        )

        outlet_pressure_text = number_text_input(
            "출구 압력",
            "outlet_pressure",
            machine,
        )

    else:
        inlet_flow_text = ""
        middle_flow_text = ""
        outlet_flow_text = ""

        inlet_pressure_text = number_text_input(
            "입구 압력",
            "inlet_pressure",
            machine,
        )

        middle_pressure_text = number_text_input(
            "중간 압력",
            "middle_pressure",
            machine,
        )

        outlet_pressure_text = number_text_input(
            "출구 압력",
            "outlet_pressure",
            machine,
        )

    memo = st.text_area(
        "비고",
        key=f"{machine}_memo",
    )

    submitted = st.button(
        "저장",
        use_container_width=True,
    )

    if submitted:

        try:
            save_time = datetime.now(KST)

            check_date = save_time.date()
            check_time = save_time.strftime("%H:%M")
            current_time = save_time.strftime("%Y-%m-%d %H:%M:%S")

            heat_temp = parse_number(
                heat_temp_text,
                "열교환기 온도",
            )

            oil_set_temp = parse_number(
                oil_set_temp_text,
                "유탕온도 설정",
            )

            oil_now_temp = parse_number(
                oil_now_temp_text,
                "유탕온도 현재",
            )

            steam_usage = parse_number(
                steam_usage_text,
                "증기 사용량",
            )

            ct_water = parse_number(
                ct_water_text,
                "C/T수",
            )

            inlet_flow = parse_number(
                inlet_flow_text,
                "입구 유량",
            )

            inlet_pressure = parse_number(
                inlet_pressure_text,
                "입구 압력",
            )

            middle_flow = parse_number(
                middle_flow_text,
                "중간 유량",
            )

            middle_pressure = parse_number(
                middle_pressure_text,
                "중간 압력",
            )

            outlet_flow = parse_number(
                outlet_flow_text,
                "출구 유량",
            )

            outlet_pressure = parse_number(
                outlet_pressure_text,
                "출구 압력",
            )

            record = (
                check_date.isoformat(),
                check_time,
                machine,
                product.strip(),
                heat_temp,
                oil_set_temp,
                oil_now_temp,
                steam_usage,
                ct_water,
                inlet_flow,
                inlet_pressure,
                middle_flow,
                middle_pressure,
                outlet_flow,
                outlet_pressure,
                memo.strip(),
                current_time,
                current_time,
            )

            result = save_or_update_record(record)

            if result == "inserted":
                st.success(
                    f"{machine}호기 점검 기록이 새로 저장되었습니다. "
                    f"저장시간: {current_time}"
                )

            elif result == "updated":
                st.success(
                    f"{machine}호기 기존 기록이 수정되었습니다. "
                    f"수정시간: {current_time}"
                )

        except ValueError as error:
            st.error(str(error))

        except sqlite3.Error as error:
            st.error(
                f"저장 또는 수정 중 오류가 발생했습니다: {error}"
            )


# ==================================================
# 10. 저장 기록 탭
# ==================================================

with tab_records:

    st.subheader("날짜별 저장 기록")

    search_date = st.date_input(
        "조회 날짜",
        key="search_date",
    )

    records = load_records(search_date)

    if not records:
        st.info(
            f"{search_date.isoformat()}에 저장된 기록이 없습니다."
        )

    else:
        display_records = []

        for record in records:
            display_records.append(
                {
                    "번호": record["id"],
                    "날짜": record["check_date"],
                    "시간": record["check_time"],
                    "호기": record["machine"],
                    "제품명": record["product"],
                    "열교환기온도": record["heat_temp"],
                    "유탕온도 설정": record["oil_set_temp"],
                    "유탕온도 현재": record["oil_now_temp"],
                    "증기사용량": record["steam_usage"],
                    "C/T수": record["ct_water"],
                    "입구유량": record["inlet_flow"],
                    "입구압력": record["inlet_pressure"],
                    "중간유량": record["middle_flow"],
                    "중간압력": record["middle_pressure"],
                    "출구유량": record["outlet_flow"],
                    "출구압력": record["outlet_pressure"],
                    "비고": record["memo"],
                    "최초 등록": record["created_at"],
                    "마지막 수정": record["updated_at"],
                }
            )

        st.write(f"총 {len(display_records)}건")

        st.dataframe(
            display_records,
            use_container_width=True,
            hide_index=True,
        )


# ==================================================
# 11. 기록 관리 탭
# ==================================================

with tab_manage:

    st.subheader("저장 기록 삭제")

    manage_date = st.date_input(
        "관리할 날짜",
        key="manage_date",
    )

    manage_records = load_records(manage_date)

    if not manage_records:
        st.info(
            f"{manage_date.isoformat()}에 저장된 기록이 없습니다."
        )

    else:
        record_options = {}

        for record in manage_records:
            label = (
                f'{record["id"]}번 | '
                f'{record["machine"]}호기 | '
                f'{record["check_time"]} | '
                f'{record["product"]}'
            )

            record_options[label] = record["id"]

        selected_label = st.selectbox(
            "삭제할 기록",
            list(record_options.keys()),
        )

        selected_record_id = record_options[selected_label]

        delete_confirmed = st.checkbox(
            "선택한 기록을 삭제하는 것에 동의합니다."
        )

        if st.button(
            "선택 기록 삭제",
            type="primary",
            use_container_width=True,
        ):

            if not delete_confirmed:
                st.warning(
                    "삭제 확인 체크박스를 선택해 주세요."
                )

            else:
                try:
                    deleted_count = delete_record(
                        selected_record_id
                    )

                    if deleted_count == 1:
                        st.success(
                            "선택한 기록이 삭제되었습니다."
                        )

                        st.rerun()

                    else:
                        st.warning(
                            "삭제할 기록을 찾을 수 없습니다."
                        )

                except sqlite3.Error as error:
                    st.error(
                        f"삭제 중 오류가 발생했습니다: {error}"
                    )


# ==================================================
# 12. Excel 다운로드 탭
# ==================================================

with tab_excel:

    st.subheader("기존 증숙기 점검일지 생성")

    excel_date = st.date_input(
        "점검일지 생성 날짜",
        key="excel_date",
    )

    records = load_records(excel_date)

    if not records:
        st.warning(
            f"{excel_date.isoformat()}에 저장된 기록이 없습니다."
        )

    else:
        try:
            excel_file, machine_count = make_template_excel(
                excel_date
            )

            file_name = (
                f"증숙기_점검일지_"
                f"{excel_date.isoformat()}.xlsx"
            )

            st.write(
                f"{machine_count}개 호기의 기록이 "
                "Excel 양식에 입력됩니다."
            )

            st.download_button(
                label="기존 양식 Excel 다운로드",
                data=excel_file,
                file_name=file_name,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except FileNotFoundError as error:
            st.error(str(error))

        except Exception as error:
            st.error(
                f"Excel 생성 중 오류가 발생했습니다: {error}"
            )
