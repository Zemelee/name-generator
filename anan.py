import openpyxl
from flask import Flask, request, render_template_string

workbook = openpyxl.load_workbook('C:\\Users\\Lenovo\\Desktop\\五格三才.xlsx')
sheet = workbook['选字']
data_list = []

start = 3
end_row = 38
# 遍历每行
for row in sheet.iter_rows(min_row=start, max_row=end_row, min_col=4, max_col=12):
    current_row_data = []  # 创建一个空列表用于存储当前行的数据
    # 遍历每行中的每列
    for cell in row:
        current_row_data.append(cell.value)  # 将单元格的值添加到行数据列表中
    data_list.append(current_row_data)  # 将行数据列表添加到总数据列表中

c6 = []
c10 = []
c12 = []
c8 = []
c15 = []
c16 = []
c17 = []
c9 = []
c7 = []

for (i, item) in enumerate(data_list):
    for (j, sub_item) in enumerate(item):
        if j == 0 and sub_item:
            c6.append(sub_item)
        elif j == 1 and sub_item:
            c7.append(sub_item)
        elif j == 2 and sub_item:
            c8.append(sub_item)
        elif j == 3 and sub_item:
            c10.append(sub_item)
        elif j == 4 and sub_item:
            c12.append(sub_item)
        elif j == 6 and sub_item:
            c15.append(sub_item)
        elif j == 7 and sub_item:
            c16.append(sub_item)
        elif j == 8 and sub_item:
            c17.append(sub_item)


def get(row, col):
    res = [f"{row}+{col}]: "]
    row_values = globals().get(f"c{row}", "")
    col_values = globals().get(f"c{col}", "")
    for item1 in row_values:
        for item2 in col_values:
            res.append(f"{item1}{item2}")
    return res



GENERATOR_TEMPLATE = '''
<!doctype html>
<html>
<head>
  <title>Hello, Anan</title>
</head>
<body>
  <h2>Computer-Aided Generator</h2>
  <form method="post">
    <input type="submit" value="生成列表" />
  </form>

  {% if result %}
    <h3>生成结果：</h3>
    <pre>
      {% for line in result %}
        {{ line }}<br>  {# 使用 <br> 标签来换行 #}
      {% endfor %}
    </pre>
  {% endif %}
</body>
</html>

'''
app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def generator():
    res = None
    if request.method == 'POST':
        res = [
            get(6, 10), "\n",
            get(6, 12), "\n",
            get(8, 10), "\n",
            get(8, 15), "\n",
            get(8, 16), "\n",
            get(8, 17), "\n",
            get(9, 15), "\n",
            get(9, 16), "\n",
            get(16, 7), "\n"
        ]
        # 将每个结果转换为字符串，并添加换行符
        res = [str(item) + "\n" for item in res]
    return render_template_string(GENERATOR_TEMPLATE, result=res)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7000)