# 將事件編號、事件名稱、處理情形放入表格當中
import  os
import pandas as pd
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR
# from pptx.dml.color import RGBColor
# from pptx.enum.shapes import MSO_SHAPE
# from pptx.enum.dml import MSO_THEME_COLOR

# ================================更改的參數=====================================
n = int(input("請輸入多少筆資料為一頁："))
# 定義dataframe中的資料
# 事件名稱
# 正則表達式匹配，[\u4e00-\u9fa5]+\d+\uff1a，匹配 資安事件1： 等字
case_name = ['事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', '事件名稱XXX', ]
# ================================更改的參數=====================================
# 事件編號，從case_name的索引取得
case_number = []
for i, j in enumerate(case_name):
    case_number.append(i+1)
# 處理結果
case_result = "處理完成"
df = pd.DataFrame(data={'事件編號': case_number,
                        '事件名稱': case_name,
                        '處理結果': case_result})
print(df)
print(df.columns)
print(df.shape)


print('開始製作PPTX')
prs = Presentation()
prs.slide_height = Cm(19.05)  # 設定ppt的高度 
prs.slide_width = Cm(25.4)  # 設定ppt的寬度

# 插入投影片
# 總計需插入？頁
# 1.如果是10的倍數筆資料(如10、20、30...)，則插入 len(case_name) // 10 頁
# 2.如果不是，則插入 len(case_name) // 10 + 1 頁
blank_page = len(case_name) // n if len(case_name) % n == 0 else len(case_name) // n + 1
for z in range(blank_page):
    blank_slide_layout = prs.slide_layouts[6]  # 用內置模板(0-10)添加一個全空的ppt頁面
    slide = prs.slides.add_slide(blank_slide_layout)
    left, top, width, height = Cm(1.76), Cm(0.49), Cm(21.89), Cm(1.28)
    tBox = slide.shapes.add_textbox(left, top, width, height)
    tf = tBox.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "事件處理統計列表"
    font = run.font
    font.name = '微軟正黑體'
    font.size = Pt(39)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
  
    # 插入表格
    rows, cols, left, top, width, height = n+1, 3, Cm(0.5), Cm(4.5), Cm(1), Cm(1.23)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    # 調整行高、列寬
    for i in range(rows):
        if i == 0:
            table.rows[i].height = Cm(1.1) 
        else:
            table.rows[i].height = Cm(1.26)
    table.columns[0].width = Cm(2.8)
    table.columns[1].width = Cm(13.6)
    table.columns[2].width = Cm(7.86)

    # 寫入表頭
    header = df.columns
    # print(header)
    for i, h in enumerate(header):
        cell = table.cell(0, i)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE  
        tf = cell.text_frame
        para = tf.paragraphs[0]
        para.text = h
        para.font.size = Pt(16)
        para.font.name = '微軟正黑體'
        para.font.bold = True
        para.alignment = PP_ALIGN.CENTER  # 居中
        
    # 按行寫入數據
    r, c = df.shape
    # print(df.shape)
    # 如果資料為n的倍數筆資料，處理方式
    if len(case_name) % n == 0:
        for i in range(0+n*z, n+n*z):
            # print('i=', i)
            for j in range(c):
                # print('j=', j)
                cell = table.cell(i+1-n*z, j)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                para = tf.paragraphs[0]
                # print(df.iloc[i,j])
                para.text = str(df.iloc[i, j])
                para.font.size = Pt(16)
                para.font.name = '微軟正黑體'        
                para.alignment = PP_ALIGN.CENTER  # 水平置中對齊
    # 如果資料不為10的倍數筆資料，處理方式
    else:
        # 最後一頁時，資料的處理方式
        if z+1 == blank_page:
            for i in range(0+n*z, len(case_name)):
                for j in range(c):
                    cell = table.cell(i+1-n*z, j)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    tf = cell.text_frame
                    para = tf.paragraphs[0]
                    para.text = str(df.iloc[i, j])
                    para.font.size = Pt(16)
                    para.font.name = '微軟正黑體'
                    para.alignment = PP_ALIGN.CENTER
        else:
            # 前面的頁數時，與一般情況相同，因此處理方式相同
            for i in range(0+n*z, n+n*z):
            # print('i=', i)
                for j in range(c):
                    # print('j=', j)
                    cell = table.cell(i+1-n*z, j)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    tf = cell.text_frame
                    para = tf.paragraphs[0]
                    # print(df.iloc[i,j])
                    para.text = str(df.iloc[i, j])
                    para.font.size = Pt(16)
                    para.font.name = '微軟正黑體'        
                    para.alignment = PP_ALIGN.CENTER  # 水平置中對齊
        # print(z)
        
prs.save('pptx_table_list.pptx')
print('PPTX製作完成')