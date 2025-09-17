import FreeSimpleGUI as sg
from datetime import datetime
from docxtpl import DocxTemplate
import os
import platform
import subprocess
import xlsxwriter

# Get current year and month
current_year = datetime.now().year
current_month = datetime.now().month

# Determine the default festival based on the current month
festivals = ["元旦节", "春节", "清明节", "五一劳动节", "端午节", "国庆节", "中秋节"]
festival_months = [1, 2, 4, 5, 6, 10, 9]
default_festival = festivals[
    next((i for i, m in enumerate(festival_months) if m >= current_month), 0)
]


def generate_docx_with_template(template_path, output_path, context):
    """Generate a Word document using a template and context."""
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)


def generate_excel_with_template(output_excel, names):
    # 创建Workbook对象
    workbook = xlsxwriter.Workbook(output_excel)
    worksheet = workbook.add_worksheet("慰问品领用表")

    # 定义格式
    title_format = workbook.add_format(
        {
            "bold": True,
            "font_size": 20,
            "align": "center",
            "valign": "vcenter",
            "border": 0,
            "bg_color": "#C6E2FF",
        }
    )

    header_format = workbook.add_format(
        {
            "bold": True,
            "font_size": 10,
            "font_color": "white",
            "bg_color": "#4F81BD",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )

    data_format = workbook.add_format(
        {"border": 1, "font_size": 10, "align": "center", "valign": "vcenter"}
    )

    currency_format = workbook.add_format(
        {
            "font_size": 10,
            "num_format": "¥#,##0.00",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        }
    )
    # 合计的格式
    sum_format = workbook.add_format(
        {
            "font_size": 10,
            "num_format": "¥#,##0.0",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        }
    )
    # 合并标题行
    worksheet.merge_range(
        "A1:E1", f"中山部{year}年{festival}慰问品领用签字表", title_format
    )  # type: ignore

    # 写入表头
    headers = ["序号", "姓名", "物品明细", "金额", "签收"]
    worksheet.write_row("A2", headers, header_format)

    # 写入数据
    data = []

    for i, name in enumerate(names, start=1):
        data.append([i, name, details, per_people_amount, ""])
    # 生成数据行

    for row_num, row_data in enumerate(data, start=3):  # 从第3行开始
        worksheet.write_row(f"A{row_num}", row_data[:1], data_format)  # 产品名称
        for col_num, value in enumerate(row_data[1:], start=1):  # 从B列开始
            worksheet.write(row_num - 1, col_num, value, currency_format)
    # 设置C列自动换行
    # worksheet.set_column("C:C", 24)
    # worksheet.set_column("C:C", None, workbook.add_format({"text_wrap": True}))

    # 合并并添加总计行
    # 合并单元格，并显示“合计”
    worksheet.merge_range(f"A{len(data) + 3}:C{len(data) + 3}", "合计", header_format)  # type: ignore
    # 合并单元格，并显示总金额
    worksheet.merge_range(
        f"D{len(data) + 3}:E{len(data) + 3}", sum(row[3] for row in data), sum_format
    )  # type: ignore
    # # 设置列宽
    worksheet.set_column("A:A", 6)  # 序号
    worksheet.set_column("B:B", 8)  # 姓名
    worksheet.set_column("C:C", 24)  # 物品明细
    worksheet.set_column("D:D", 4)  # 金额
    worksheet.set_column("E:D", 20)  # 签收

    # 设置行高
    worksheet.set_row(0, 30)  # 设置第0行，标题行高
    worksheet.set_row(1, 20)  # 设置第1行，表头行高
    for row_num in range(2, len(data) + 2):  # 设置从第3行开始的数据行高
        worksheet.set_row(row_num, 35)
    # 保存文件
    workbook.close()
    print("Excel文件已生成: sales_report_with_merged.xlsx")


def open_file_with_default_program(file_path):
    """Open a file with the default program based on the OS."""
    system_name = platform.system()
    try:
        if system_name == "Windows":
            subprocess.run(["start", file_path], shell=True)
        elif system_name == "Darwin":  # macOS
            subprocess.run(["open", file_path])
        elif system_name == "Linux":
            subprocess.run(["xdg-open", file_path])
        else:
            sg.popup("无法识别的操作系统", title="错误")
    except Exception as e:
        sg.popup(f"无法打开文件: {e}", title="错误")


# Layout for the GUI
layout = [
    [
        sg.Text(
            "工会慰问品领用方案和领用签字表生成器",
            justification="center",
            expand_x=True,
        )
    ],
    [
        [
            sg.Text("年份:"),
            sg.Combo(list(range(2005, 2031)), default_value=current_year, key="year"),
            sg.Text("节日:"),
            sg.Combo(
                festivals, default_value=default_festival, key="festival", size=(10, 1)
            ),
            sg.Text("人数:"),
            sg.Combo([12, 16], default_value=12, key="people"),
        ],
    ],
    [
        sg.Text("慰问品总金额:"),
        sg.Input(key="total_amount", size=(10, 1)),
        sg.Button("查看人均金额"),
        sg.Text("人均金额："),
        sg.Text("0", key="avg_price"),
    ],
    [sg.Button("生成方案"), sg.Button("生成领用表"), sg.Button("退出")],
    [
        sg.Multiline(
            size=(50, 10),
            key="details",
            disabled=False,
            default_text="""1.明治鲜牛奶950ml*2盒
2.盒马椰子水250ml*6盒
3.简爱0添加酸奶135g*4盒
4.冰鲜原切澳洲谷饲牛腱1kg
5.进口牛肋条800g
6.黄金香葡萄550g
7.陕西冰糖冬枣400g
8.佳农进口香蕉800g""",
        )
    ],
]

# Create the window
window = sg.Window("工会工作方案和领用表生成器", layout, font=("微软雅黑", 16))

NAME_12 = [
    "张林",
    "高雅妮",
    "郑晓臻",
    "李艳丽",
    "陆琼",
    "廖方智",
    "王伟",
    "钟超",
    "黄浩文",
    "刘锦城",
    "邹颖",
    "禹枭",
]
NAME_16 = NAME_12 + [
    "叶若芳",
    "何照月",
    "何俊璋",
    "唐晓亮",
]
# Event loop
while True:
    event, values = window.read()  # type: ignore
    if event == sg.WINDOW_CLOSED or event == "退出":
        break
    elif event == "查看人均金额":
        if values["total_amount"] == "":
            sg.popup("请输入总金额", title="错误", keep_on_top=True)
            continue
        avg = round(float(values["total_amount"]) / int(values["people"]), 2)
        if avg >= 300:
            sg.popup("人均金额超标", title="错误", keep_on_top=True)
            continue
        window["avg_price"].update(str(round(avg, 2)))
        # print(values["avg_price"])
    elif event == "生成方案":
        if values["total_amount"] == "":
            sg.popup("请输入总金额", title="错误", keep_on_top=True)
            continue
        try:
            output_path_doxc = f"{values['year']}年{values['festival']}慰问品方案.docx"

            people_str = ""
            if int(values["people"]) == 12:
                people_str = f"员工{values['people']}人"
            elif int(values["people"]) == 16:
                people_str = f"员工12人,客户经理4人"
            context = {
                "年份": values["year"],
                "节日名称": values["festival"].strip(),
                "人数": people_str,
                "总份数": values["people"],
                "慰问品内容": values["details"],
                "总金额": values["total_amount"],
                "签字日期": f"日期：{datetime.now().strftime('%Y-%m-%d')}",
                "人均金额": round(
                    float(values["total_amount"]) / int(values["people"]), 2
                ),
            }
            generate_docx_with_template(
                "工会慰问品派发方案模板.docx", output_path_doxc, context
            )
            sg.popup("方案生成成功", title="提示")
            open_file_with_default_program(output_path_doxc)
        except Exception as e:
            sg.popup(f"生成方案失败: {e}", title="错误")
    elif event == "生成领用表":
        try:
            if values["total_amount"] == "":
                sg.popup("请输入总金额", title="错误")
                continue
            output_path_excel = (
                f"{values['year']}年{values['festival']}慰问品领用签字表.xlsx"
            )
            year = values["year"]
            festival = values["festival"].strip()
            people = int(values["people"])
            details = values["details"]
            total_amount = float(values["total_amount"])
            sign_data = (f"{datetime.now().strftime('%Y-%m-%d')}",)
            per_people_amount = round((total_amount / people), 2)
            if people == 12:
                names = NAME_12
            elif people == 16:
                names = NAME_16
            else:
                sg.popup("人数不在范围内", title="错误")
                continue
            generate_excel_with_template(output_path_excel, names)
            sg.popup("领用表生成成功", title="提示")
            open_file_with_default_program(output_path_excel)
        except Exception as e:
            sg.popup(f"生成领用表失败: {e}", title="错误")

window.close()
