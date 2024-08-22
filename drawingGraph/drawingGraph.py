import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

if __name__ == '__main__':  

    # 엑셀 파일에서 데이터 읽기 (파일 경로와 시트 이름을 수정하세요)
    file_path = 'Book1.xlsx'
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)

    # 2번째 행: 넘버 (0-based index로 1번째 행)
    number_row = df.iloc[2, 2:]  # 4번째 열부터 끝까지

    # 3번째 행: 저항 값 (0-based index로 2번째 행)
    resistance_row = df.iloc[3, 2:]  # 4번째 열부터 끝까지

    # 4번째 행: 게이지 값 (0-based index로 3번째 행)
    gauge_row_1 = df.iloc[4, 2:]  # 4번째 열부터 끝까지

    gauge_row_2 = df.iloc[5, 2:] 

    # 그래프 그리기
    plt.figure(figsize=(10, 6))

    plt.plot(resistance_row, gauge_row_1, marker='o', linestyle='-', color='blue', label='Spec Gauge vs Resistance')

    plt.plot(resistance_row, gauge_row_2, marker='s', linestyle='--', color='green', label='Actual Gauge  vs Resistance')


    plt.title('Gauge vs Resistance Graph')
    plt.xlabel('Resistance')
    plt.ylabel('Gauge')
    plt.legend()

    # 그래프를 이미지 파일로 저장 (선택 사항)
    #plt.savefig('gauge_vs_resistance_graph.png')

    graph_image_path = 'gauge_vs_resistance_graph.png'
    plt.savefig(graph_image_path)


    wb = load_workbook(file_path)
    ws = wb.active

    img = Image(graph_image_path)

    ws.add_image(img, 'B25')

    wb.save(file_path)

    # 그래프 화면에 표시
    #plt.show()
