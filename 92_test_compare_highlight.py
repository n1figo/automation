import pytesseract
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
import cv2
import pdfplumber
import re
import os

# ... (이전의 extract_table_from_image, extract_table_from_pdf, preprocess_dataframe, compare_dataframes 함수들은 그대로 유지) ...

def highlight_changes_on_image(image_path, changes, output_path):
    print("변경 사항을 이미지에 표시하는 중...")
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 20)  # 폰트 경로 확인 필요

    y_position = 10
    for change in changes:
        # 빨간 테두리와 노란 배경의 네모 그리기
        draw.rectangle([10, y_position, image.width - 10, y_position + 80], 
                       outline="red", fill=(255, 255, 0, 64), width=2)
        # 변경 내용 텍스트 추가
        draw.text((15, y_position + 5), change[:50] + "...", fill="black", font=font)
        y_position += 85

    image.save(output_path)
    print(f"변경 사항이 표시된 이미지가 저장되었습니다: {output_path}")

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. ㅇKB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_image_path = os.path.join(output_dir, "highlighted_changes.png")

    df_image = extract_table_from_image(image_path)
    df_pdf = extract_table_from_pdf(pdf_path)

    if df_image is not None and df_pdf is not None:
        changes = compare_dataframes(df_image, df_pdf)

        print("\n감지된 변경 사항:")
        for i, change in enumerate(changes, 1):
            print(f"변경 사항 {i}: {change}")

        highlight_changes_on_image(image_path, changes, output_image_path)
    else:
        print("표 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()