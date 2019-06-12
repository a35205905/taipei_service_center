import matplotlib.pyplot  as plt

from openpyxl import load_workbook 
from matplotlib.font_manager import _rebuild

# 重新建立字型索引列表
_rebuild()
# 設定中文字型
plt.rcParams['font.sans-serif']=[u'simhei']
# 圖片大小(寬, 高)
plt.rcParams['figure.figsize'] = (10.0, 8.0)

def main():
    wb = load_workbook('107年臺北市旅遊服務中心基礎統計表(OPENDATA).xlsx', data_only=True)
    # 預設讀取第一張工作表
    ws = wb.active

    month = []
    # 取月份
    column = list(ws.columns)[0]
    for cell in list(column)[4:16]:
        month.append(cell.value)

    # 1-12月資料
    for column in list(ws.columns)[1:13]:
        location = column[2].value.replace('\n', '')
        service = []
        for cell in list(column)[4:16]:
            service.append(cell.value)

        plt.plot(month, service, label=location)

    plt.xlabel('Month')
    plt.ylabel('Service Times')
    plt.title('107 Taipei Service Center')
    # 圖例
    plt.legend()
    # 設定圖例位置
    plt.legend(loc='upper left', borderaxespad=0.)
    # 格線
    plt.grid(True)
    # 產生圖檔
    plt.savefig('taipei_service_center.png', dpi=300, format='png')
    # 預覽畫面
    plt.show()

if __name__ == "__main__":
    main()