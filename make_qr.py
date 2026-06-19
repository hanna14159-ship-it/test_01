from pathlib import Path

import qrcode


# 여기에 실제 Streamlit 앱 주소 입력
APP_URL = "https://your-app-name.streamlit.app"

MACHINES = [
    "11", "12", "13", "14",
    "21", "22", "23", "24",
    "51", "52", "53", "54",
]

OUTPUT_FOLDER = Path("qr_codes")

OUTPUT_FOLDER.mkdir(exist_ok=True)


for machine in MACHINES:

    url = f"{APP_URL}/?machine={machine}"

    qr_image = qrcode.make(url)

    file_path = OUTPUT_FOLDER / f"steamer_{machine}.png"

    qr_image.save(file_path)

    print(f"{machine}호기 QR 생성 완료: {file_path}")
