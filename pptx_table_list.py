# -*- coding:utf-8 -*-
# 將事件編號、事件名稱、處理情形放入表格當中
import math
from typing import Any, Dict, List, Text

import arrow
import my_case_name_list
import polars as pl
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Pt

# from pptx.dml.color import RGBColor
# from pptx.enum.shapes import MSO_SHAPE
# from pptx.enum.dml import MSO_THEME_COLOR


class PTTXReport(object):
    def __init__(self) -> None:
        self.prs = Presentation()
        self.prs.slide_height, self.prs.slide_width = Cm(19.05), Cm(25.4)
        self.title_list = ["事件編號", "事件名稱", "處理結果"]
        self.title_num = len(self.title_list)
        self.update_date = arrow.now().format("YYYYMMDD")

    def data_to_table(
        self,
        table,
        data_list: List[Dict[Text, Any]],
    ):
        """將資料寫入table中

        Args:
            table (_type_): 每一頁要操作的table
            data_list (Union[List[Text], List[Dict[Text, Any]]]): 有可能是header的資料或是事件的相關資料
        """
        data_list.insert(0, {title: title for title in self.title_list})
        for row_index, data in enumerate(data_list):
            for column_index, (title, value) in enumerate(data.items()):
                cell = table.cell(row_index, column_index)
                tf = cell.text_frame
                para = tf.paragraphs[0]
                para.text = value
                para.font.size = Pt(16)
                para.font.name = "微軟正黑體"
                para.font.bold = True if row_index == 0 else False
                para.alignment = PP_ALIGN.CENTER  # 水平置中對齊

    def run_all(self):
        n = int(input("請輸入？筆資料為一頁："))
        case_list = (
            pl.LazyFrame(
                {
                    self.title_list[1]: my_case_name_list.case_name_list,
                }
            )
            .with_row_index(self.title_list[0], offset=1)
            .with_columns(pl.lit("處理完成").alias(self.title_list[2]))
            .cast(pl.String)
            .collect()
            .to_dicts()
        )
        # 總計需插入？頁，使用無條件進位處理
        blank_page = int(math.ceil(len(case_list) / n))
        for page in range(blank_page):
            # 用內置模板(0-10)添加一個全空的ppt頁面
            blank_slide_layout = self.prs.slide_layouts[6]
            slide = self.prs.slides.add_slide(blank_slide_layout)
            left, top, width, height = Cm(1.76), Cm(0.49), Cm(21.89), Cm(1.28)
            tBox = slide.shapes.add_textbox(left, top, width, height)
            tf = tBox.text_frame
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = "事件處理統計列表"
            font = run.font
            font.name = "微軟正黑體"
            font.size = Pt(39)
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

            # 插入表格
            rows = n + 1
            cols = self.title_num
            left = Cm(0.5)
            top = Cm(4.5)
            width = Cm(1)
            height = Cm(1.23)
            table = slide.shapes.add_table(rows, cols, left, top, width, height).table
            # 調整行高、列寬
            for index in range(rows):
                table.rows[index].height = (
                    Cm(12.6 / n * (1.1 / 1.26)) if index == 0 else Cm(12.6 / n)
                )
            table.columns[0].width = Cm(2.8)
            table.columns[1].width = Cm(13.6)
            table.columns[2].width = Cm(7.86)
            # 最後一頁的case_list的範圍稍微不同
            self.data_to_table(
                table,
                case_list[0 + page * n : (page + 1) * n],
            ) if (page + 1 < blank_page) else self.data_to_table(
                table,
                case_list[0 + page * n :],
            )
        self.prs.save(f"report_{self.update_date}.pptx")

    def main(self):
        current_except = None
        try:
            self.run_all()
        except Exception as e:
            print(e)
            current_except = e
        else:
            print("PPTX製作完成")
        finally:
            if current_except:
                raise current_except


if __name__ == "__main__":
    PTTXReport().main()
