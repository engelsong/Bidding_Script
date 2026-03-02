from __future__ import annotations

import re
from datetime import date, datetime
from typing import Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


class Quotation:
    """通过 project 实例创建报价表（不依赖模板）。"""

    def __init__(self, project) -> None:
        self.project = project

        self.title_font = Font(name="宋体", size=16, bold=True)
        self.header_font = Font(name="宋体", size=12, bold=True)
        self.normal_font = Font(name="宋体", size=11)
        self.center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        self.right = Alignment(horizontal="right", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="000000")
        self.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.import_fill = PatternFill("solid", fgColor="FFFF00")
        self.header_fill = PatternFill("solid", fgColor="808080")

    @staticmethod
    def _safe_name(name: str) -> str:  # 用于生成文件名，过滤掉无法使用的字符
        invalid = '<>:"/\\|?*'
        result = name or "project"
        for ch in invalid:
            result = result.replace(ch, "_")
        return result.strip()

    @staticmethod
    def _parse_date(raw) -> date:  # 生成日期
        if isinstance(raw, datetime):
            return raw.date()
        if isinstance(raw, date):
            return raw
        if isinstance(raw, str):
            m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", raw)
            if m:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return date.today()

    @staticmethod
    def _parse_quantity(raw) -> Optional[int]:  # 用于统计物资数量，将字符串转化为int
        if raw is None:
            return 1
        if isinstance(raw, int):
            return raw
        if isinstance(raw, float):
            return int(raw)
        text = str(raw).strip().replace(",", "")
        m = re.search(r"-?\d+", text)
        if not m:
            return None
        return int(m.group(0))

    def _set_columns(self, ws, widths: Dict[str, float]) -> None:  #设置当前worksheet的列宽
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

    def _style_row(self, ws, row: int, col_start: int, col_end: int, header: bool = False) -> None:
        #  用于修改指定worksheet行号的从几列到几列的样式
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row, col)
            cell.font = self.header_font if header else self.normal_font
            cell.alignment = self.center
            cell.border = self.border
            if header:
                cell.fill = self.header_fill

    def _build_all_suppliers(self, ws, items: List[Dict]) -> None:
        ws.title = "全部厂家备用"
        self._set_columns(
            ws,
            {
                "A": 7.6,
                "B": 13.1,
                "C": 12.2,
                "D": 10.6,
                "E": 9.8,
                "F": 42.0,
                "G": 12.2,
                "H": 10.0,
                "I": 14.1,
                "J": 13.1,
                "K": 16.2,
                "L": 8.2,
                "M": 7.7,
                "N": 9.6,
                "O": 12.8,
                "P": 12.1,
                "Q": 8.1,
                "R": 11.2,
                "S": 10.0,
                "T": 13.0,
                "U": 13.0,
                "V": 13.0,
                "W": 13.0,
                "X": 13.0,
                "Y": 13.0,
                "Z": 13.0,
                "AA": 13.0,
                "AB": 6.0,
                "AC": 13.0,
                "AD": 15.3,
                "AE": 13.0,
                "AF": 15.6,
                "AG": 14.2,
                "AH": 13.6,
                "AI": 16.7,
                "AJ": 13.5,
            },
        )
        headers = [
            "序号",
            "品名",
            "HS编码",
            "数量",
            "单位",
            "规格",
            "检验标准",
            "品牌",
            "型号",
            "单价",
            "总价",
            "生产厂商",
            "供货商",
            "偏离情况",
            "强制性认证取证情况",
            "生产供货地",
            "交货期",
            "质量体系",
            "有效期",
            "环境体系",
            "有效期",
            "职业健康",
            "有效期",
            "节能",
            "有效期",
            "环境标志",
            "有效期",
            "备注",
            "",
            "排名",
            "性价比",
            "得分",
            "性能",
            "质量保证",
            "三体系",
            "售后",
        ]
        ws.append(headers)
        self._style_row(ws, 1, 1, len(headers), header=True)

        for idx, item in enumerate(items, start=2):
            ws[f"A{idx}"] = item["serial"]
            ws[f"B{idx}"] = item["name"]
            ws[f"C{idx}"] = item["hs"]
            ws[f"D{idx}"] = item["qty"]
            ws[f"E{idx}"] = item["unit"]
            ws[f"F{idx}"] = item["spec"]
            ws[f"G{idx}"] = item["standard"]
            ws[f"H{idx}"] = ""
            ws[f"I{idx}"] = ""
            ws[f"J{idx}"] = 0
            ws[f"K{idx}"] = f"=D{idx}*J{idx}"
            ws[f"L{idx}"] = ""
            ws[f"M{idx}"] = ""
            ws[f"N{idx}"] = "无"
            ws[f"O{idx}"] = "无"
            ws[f"P{idx}"] = self.project.destination
            ws[f"Q{idx}"] = self.project.trans_time
            ws[f"R{idx}"] = ""
            ws[f"S{idx}"] = ""
            ws[f"T{idx}"] = ""
            ws[f"U{idx}"] = ""
            ws[f"V{idx}"] = ""
            ws[f"W{idx}"] = ""
            ws[f"X{idx}"] = ""
            ws[f"Y{idx}"] = ""
            ws[f"Z{idx}"] = ""
            ws[f"AA{idx}"] = ""
            ws[f"AB{idx}"] = ""
            ws[f"AC{idx}"] = ""
            ws[f"AD{idx}"] = 1
            ws[f"AE{idx}"] = ""
            ws[f"AF{idx}"] = ""
            ws[f"AG{idx}"] = ""
            ws[f"AH{idx}"] = ""
            ws[f"AI{idx}"] = ""
            ws[f"AJ{idx}"] = ""
            self._style_row(ws, idx, 1, len(headers), header=False)
            ws[f"F{idx}"].alignment = self.left
            ws.row_dimensions[idx].height = max(24, min(120, (str(item["spec"]).count("\n") + 1) * 16))

    def _build_selector(self, ws, items: List[Dict]) -> None:
        ws.title = "物资选择"
        self._set_columns(ws, {"A": 8, "B": 28, "C": 10})
        for i, item in enumerate(items, start=1):
            ws[f"A{i}"] = item["serial"]
            ws[f"B{i}"] = item["name"]
            ws[f"C{i}"] = i + 1  # 对应“全部厂家备用”的行号
            self._style_row(ws, i, 1, 3, header=False)

    def _build_fee_input(self, ws) -> None:
        ws.title = "费用输入"
        self._set_columns(ws, {"A": 24, "B": 18})
        rows = [
            ("项目", "金额"),
            ("物资检验费", 0),
            ("运输保险费", 0),
            ("国外运费", 0),
            ("安防费", 0),
            ("税金费率", 0.0003),
            ("对外金额", float(self.project.totalsum)),
            ("资金占用时间(月)", 0),
            ("年化利率", 0.0),
            ("资金占用费", "=ROUND(B7*B8/12*B9,2)"),
        ]
        for r, (name, val) in enumerate(rows, start=1):
            ws[f"A{r}"] = name
            ws[f"B{r}"] = val
            self._style_row(ws, r, 1, 2, header=(r == 1))

    def _build_inner_quote(self, ws, item_count: int) -> int:
        ws.title = "2.物资对内分项报价表"
        self._set_columns(
            ws,
            {
                "A": 12,
                "B": 30,
                "C": 14,
                "D": 10,
                "E": 14,
                "F": 12,
                "G": 10,
                "H": 10,
                "I": 12,
                "J": 12,
                "K": 12,
                "L": 12,
                "M": 12,
                "N": 14,
            },
        )

        ws.merge_cells("A1:N1")
        ws["A1"] = "二.物资对内分项报价表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:N2")
        ws["A2"] = "报价单位：人民币元（保留小数点后两位）"
        ws["A2"].font = self.normal_font
        ws["A2"].alignment = self.left

        headers = [
            "物资",
            "品名",
            "商品购买单价",
            "数量",
            "商品购买总价",
            "国内运杂费",
            "包装费",
            "保管费",
            "物资检验费",
            "运输保险费",
            "国外运费",
            "实施服务费",
            "税金",
            "合计",
        ]
        for c, h in enumerate(headers, start=1):
            ws.cell(3, c, h)
        self._style_row(ws, 3, 1, 14, header=True)

        item_start = 4
        item_end = item_start + item_count - 1
        total_row = item_end + 1
        summary_row = total_row + 2

        ws.merge_cells(f"A{item_start}:A{item_end}")
        ws[f"A{item_start}"] = "供货清单（一）"
        ws[f"A{item_start}"].alignment = self.center
        ws[f"A{item_start}"].font = self.normal_font

        for i in range(1, item_count + 1):
            row = item_start + i - 1
            ws[f"B{row}"] = f"=物资选择!A{i}&\".\"&物资选择!B{i}"
            ws[f"C{row}"] = f"=INDEX('全部厂家备用'!J:J,物资选择!C{i})"
            ws[f"D{row}"] = f"=INDEX('全部厂家备用'!D:D,物资选择!C{i})"
            ws[f"E{row}"] = f"=C{row}*D{row}"
            ws[f"F{row}"] = 0
            ws[f"G{row}"] = 0
            ws[f"H{row}"] = 0
            ws[f"I{row}"] = f"=ROUND(E{row}/E${total_row}*I${summary_row},2)"
            ws[f"J{row}"] = f"=ROUND(E{row}/E${total_row}*J${summary_row},2)"
            ws[f"K{row}"] = f"=ROUND(E{row}/E${total_row}*K${summary_row},2)"
            ws[f"L{row}"] = f"=ROUND(E{row}/E${total_row}*L${summary_row},2)"
            ws[f"M{row}"] = f"=ROUND(E{row}/E${total_row}*M${summary_row},2)"
            ws[f"N{row}"] = f"=SUM(E{row}:M{row})"
            self._style_row(ws, row, 2, 14, header=False)

        ws[f"A{total_row}"] = "合计"
        ws.merge_cells(f"A{total_row}:B{total_row}")
        ws[f"A{total_row}"].font = self.header_font
        ws[f"A{total_row}"].alignment = self.center
        ws[f"A{total_row}"].border = self.border
        ws[f"B{total_row}"].border = self.border
        for col in "EFGHIJKLMN":
            ws[f"{col}{total_row}"] = f"=SUM({col}{item_start}:{col}{item_end})"
        self._style_row(ws, total_row, 3, 14, header=False)

        ws[f"H{summary_row}"] = "分摊基数"
        ws[f"I{summary_row}"] = "=费用输入!B2"
        ws[f"J{summary_row}"] = "=费用输入!B3"
        ws[f"K{summary_row}"] = "=费用输入!B4"
        ws[f"L{summary_row}"] = "='4.技术服务费报价表'!H14"
        ws[f"M{summary_row}"] = f"=ROUND((SUM(E{total_row}:L{total_row}))*费用输入!B6,2)"
        ws[f"N{summary_row}"] = f"=SUM(E{total_row}:M{total_row})"
        self._style_row(ws, summary_row, 8, 14, header=False)
        return total_row

    def _build_tax_sheet(self, ws, item_count: int, inner_total_row: int) -> int:
        ws.title = "3.各项物资退抵税额表"
        self._set_columns(
            ws,
            {"A": 8, "B": 26, "C": 20, "D": 16, "E": 16, "F": 16, "G": 16, "H": 18},
        )

        ws.merge_cells("A1:H1")
        ws["A1"] = "三.增值税和消费税退抵税额表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center
        ws.merge_cells("A2:H2")
        ws["A2"] = "报价单位：人民币元（保留小数点后两位）"
        ws["A2"].font = self.normal_font
        ws["A2"].alignment = self.left

        headers = [
            "序号",
            "品名",
            "购买价款",
            "退抵增值税率(%)",
            "退抵增值税额",
            "退抵消费税率(%)",
            "退抵消费税额",
            "退抵税总额",
        ]
        for c, h in enumerate(headers, start=1):
            ws.cell(3, c, h)
        self._style_row(ws, 3, 1, 8, header=True)

        item_start = 4
        item_end = item_start + item_count - 1
        trans_row = item_end + 1
        insurance_row = item_end + 2
        inspect_row = item_end + 3
        total_row = item_end + 4

        for i in range(1, item_count + 1):
            row = item_start + i - 1
            ws[f"A{row}"] = f"=物资选择!A{i}"
            ws[f"B{row}"] = f"=物资选择!B{i}"
            ws[f"C{row}"] = f"='2.物资对内分项报价表'!E{3+i}"
            ws[f"D{row}"] = 13
            ws[f"E{row}"] = f"=ROUND(C{row}/(1+D{row}/100)*D{row}/100,2)"
            ws[f"F{row}"] = 0
            ws[f"G{row}"] = f"=ROUND(C{row}/(1+F{row}/100)*F{row}/100,2)"
            ws[f"H{row}"] = f"=E{row}+G{row}"
            self._style_row(ws, row, 1, 8, header=False)

        ws[f"B{trans_row}"] = "运输"
        ws[f"C{trans_row}"] = f"='2.物资对内分项报价表'!K{inner_total_row}"
        ws[f"D{trans_row}"] = 0
        ws[f"E{trans_row}"] = f"=ROUND(C{trans_row}/(1+D{trans_row}/100)*D{trans_row}/100,2)"
        ws[f"F{trans_row}"] = 0
        ws[f"G{trans_row}"] = f"=ROUND(C{trans_row}/(1+F{trans_row}/100)*F{trans_row}/100,2)"
        ws[f"H{trans_row}"] = f"=E{trans_row}+G{trans_row}"

        ws[f"B{insurance_row}"] = "保险"
        ws[f"C{insurance_row}"] = f"='2.物资对内分项报价表'!J{inner_total_row}"
        ws[f"D{insurance_row}"] = 0
        ws[f"E{insurance_row}"] = (
            f"=ROUND(C{insurance_row}/(1+D{insurance_row}/100)*D{insurance_row}/100,2)"
        )
        ws[f"F{insurance_row}"] = 0
        ws[f"G{insurance_row}"] = (
            f"=ROUND(C{insurance_row}/(1+F{insurance_row}/100)*F{insurance_row}/100,2)"
        )
        ws[f"H{insurance_row}"] = f"=E{insurance_row}+G{insurance_row}"

        ws[f"B{inspect_row}"] = "第三方检验"
        ws[f"C{inspect_row}"] = f"='2.物资对内分项报价表'!I{inner_total_row}"
        ws[f"D{inspect_row}"] = 0
        ws[f"E{inspect_row}"] = (
            f"=ROUND(C{inspect_row}/(1+D{inspect_row}/100)*D{inspect_row}/100,2)"
        )
        ws[f"F{inspect_row}"] = 0
        ws[f"G{inspect_row}"] = (
            f"=ROUND(C{inspect_row}/(1+F{inspect_row}/100)*F{inspect_row}/100,2)"
        )
        ws[f"H{inspect_row}"] = f"=E{inspect_row}+G{inspect_row}"
        self._style_row(ws, trans_row, 1, 8, header=False)
        self._style_row(ws, insurance_row, 1, 8, header=False)
        self._style_row(ws, inspect_row, 1, 8, header=False)

        ws[f"A{total_row}"] = "共计"
        ws.merge_cells(f"A{total_row}:B{total_row}")
        ws[f"A{total_row}"].font = self.header_font
        ws[f"A{total_row}"].alignment = self.center
        ws[f"A{total_row}"].border = self.border
        ws[f"B{total_row}"].border = self.border
        ws[f"C{total_row}"] = f"=SUM(C{item_start}:C{inspect_row})"
        ws[f"E{total_row}"] = f"=SUM(E{item_start}:E{inspect_row})"
        ws[f"G{total_row}"] = f"=SUM(G{item_start}:G{inspect_row})"
        ws[f"H{total_row}"] = f"=SUM(H{item_start}:H{inspect_row})"
        self._style_row(ws, total_row, 3, 8, header=False)
        return total_row

    def _build_tech_sheet(self, ws) -> None:
        ws.title = "4.技术服务费报价表"
        self._set_columns(ws, {"A": 6, "B": 22, "C": 12, "D": 12, "E": 8, "F": 10, "G": 14, "H": 14})

        ws.merge_cells("A1:H1")
        ws["A1"] = "四.技术服务费报价表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center

        headers = ["序号", "费用名称", "美元单价", "人民币单价", "人数", "天/次数", "美元合计", "人民币合计"]
        for c, h in enumerate(headers, start=1):
            ws.cell(2, c, h)
        self._style_row(ws, 2, 1, 8, header=True)

        people = self.project.techinfo[0] if self.project.is_tech and len(self.project.techinfo) >= 2 else 0
        days = self.project.techinfo[1] if self.project.is_tech and len(self.project.techinfo) >= 2 else 0

        items = [
            "护照和签证手续费",
            "防疫免疫费",
            "技术服务人员保险费",
            "国内交通费",
            "国际交通费",
            "住宿费",
            "伙食费",
            "津贴补贴",
            "当地雇工费",
            "当地设备工具材料购置或租用费",
            "其它确需发生的费用",
        ]
        for i, name in enumerate(items, start=3):
            ws[f"A{i}"] = i - 2
            ws[f"B{i}"] = name
            ws[f"D{i}"] = 0
            ws[f"E{i}"] = people
            ws[f"F{i}"] = days if i in (8, 9, 10) else "-"
            ws[f"G{i}"] = f"=C{i}*E{i}*IF(F{i}=\"-\",1,F{i})"
            ws[f"H{i}"] = f"=D{i}*E{i}*IF(F{i}=\"-\",1,F{i})"
            self._style_row(ws, i, 1, 8, header=False)
        ws["A14"] = "共计："
        ws.merge_cells("A14:F14")
        ws["A14"].font = self.header_font
        ws["A14"].alignment = self.center
        for col in range(1, 7):
            ws.cell(14, col).border = self.border
        ws["G14"] = "=SUM(G3:G13)"
        ws["H14"] = "=SUM(H3:H13)"
        self._style_row(ws, 14, 7, 8, header=False)

    def _build_train_sheet(self, ws) -> None:
        ws.title = "5.来华培训费报价表"
        self._set_columns(ws, {"A": 6, "B": 20, "C": 12, "D": 14, "E": 8, "F": 10, "G": 16})

        ws.merge_cells("A1:G1")
        ws["A1"] = "五.来华培训费报价表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center

        headers = ["序号", "费用名称", "标准", "费用计算方式", "人数", "天(次)数", "人民币(元)"]
        for c, h in enumerate(headers, start=1):
            ws.cell(2, c, h)
        self._style_row(ws, 2, 1, 7, header=True)

        num = self.project.training_num if self.project.is_cc else 0
        days = self.project.training_days if self.project.is_cc else 0
        stay_days = max(days - 1, 0)

        rows = [
            ("一", "培训费", "360元/人*天", 360, num, days),
            ("二-1", "日常伙食费", "190元/人*天", 190, num, days),
            ("二-2", "住宿费", "350元/人*天", 350, num, stay_days),
            ("二-3", "宴请费", "200元/人*次", 200, num, 1 if num else 0),
            ("二-4", "零用费", "150元/人*天", 150, num, days),
            ("二-5", "小礼品费", "200元/人", 200, num, 1 if num else 0),
            ("二-6", "人身意外伤害保险", "150元/人", 150, num, "-"),
            ("三", "国际旅费", "5000元/人", 5000, num, "-"),
            ("四-1", "承办管理费", "6%", None, None, None),
            ("四-2", "管理人员伙食费", "190元/人*天", 190, 2 if num else 0, days if num else 0),
            ("四-3", "管理人员住宿费", "350元/人*天", 350, 2 if num else 0, days if num else 0),
        ]

        start = 3
        for idx, row in enumerate(rows, start=start):
            code, name, std, unit_price, people, day = row
            ws[f"A{idx}"] = code
            ws[f"B{idx}"] = name
            ws[f"C{idx}"] = std
            if unit_price is not None:
                ws[f"D{idx}"] = unit_price
                ws[f"E{idx}"] = people
                ws[f"F{idx}"] = day
                if isinstance(day, str):
                    ws[f"G{idx}"] = f"=D{idx}*E{idx}"
                else:
                    ws[f"G{idx}"] = f"=D{idx}*E{idx}*F{idx}"
            else:
                ws[f"G{idx}"] = f"=ROUND((SUM(G{start}:G{start+7}))*0.06,2)"
            self._style_row(ws, idx, 1, 7, header=False)

        total_row = start + len(rows)
        ws[f"A{total_row}"] = "五"
        ws[f"B{total_row}"] = "合计"
        ws[f"G{total_row}"] = f"=SUM(G{start}:G{total_row-1})"
        self._style_row(ws, total_row, 1, 7, header=False)

    def _build_total_sheet(
        self, ws, inner_total_row: int, tax_total_row: int, bid_date: date
    ) -> None:
        ws.title = "1.投标报价总表"
        self._set_columns(ws, {"A": 8, "B": 40, "C": 20, "D": 24})

        ws.merge_cells("A1:D1")
        ws["A1"] = "一.投标报价总表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center

        ws.merge_cells("A2:D2")
        ws["A2"] = "报价单位：人民币元（保留小数点后两位）"
        ws["A2"].font = self.normal_font
        ws["A2"].alignment = self.left

        headers = ["序号", "费用项目", "合计金额", "备注"]
        for c, h in enumerate(headers, start=1):
            ws.cell(3, c, h)
        self._style_row(ws, 3, 1, 4, header=True)

        rows = [
            ("一", "全部物资价格", f"='2.物资对内分项报价表'!N{inner_total_row}", f"{self.project.trans}{self.project.destination}价"),
            ("二", "技术服务费", "='4.技术服务费报价表'!H14", ""),
            ("三", "来华培训和接待费", "='5.来华培训费报价表'!G14", ""),
            ("四", "其他费用-安防费用", "=费用输入!B5", ""),
            ("五", "增值税和消费税退抵税额", f"='3.各项物资退抵税额表'!H{tax_total_row}", ""),
        ]
        for i, (no, item, amount, note) in enumerate(rows, start=4):
            ws[f"A{i}"] = no
            ws[f"B{i}"] = item
            ws[f"C{i}"] = amount
            ws[f"D{i}"] = note
            self._style_row(ws, i, 1, 4, header=False)
            ws[f"B{i}"].alignment = self.left

        ws["B9"] = "共计"
        ws["C9"] = "=SUM(C4:C7)-C8"
        self._style_row(ws, 9, 2, 3, header=False)
        ws["B9"].font = self.header_font
        ws["C9"].font = self.header_font

        ws["C13"] = "投标人盖章："
        ws["C14"] = "日期："
        ws["D14"] = bid_date
        ws["C13"].font = self.normal_font
        ws["C14"].font = self.normal_font
        ws["D14"].font = self.normal_font

    def _build_opening_sheet(self, ws, bid_date: date) -> None:
        ws.title = "3.开标一览表"
        self._set_columns(ws, {"A": 20, "B": 24, "C": 48, "D": 24})

        ws.merge_cells("A1:D1")
        ws["A1"] = "三.开标一览表"
        ws["A1"].font = self.title_font
        ws["A1"].alignment = self.center

        ws["A2"] = "招标编号："
        ws["B2"] = self.project.code
        ws["C2"] = f"项目名称：{self.project.name}"
        for c in ("A2", "B2", "C2"):
            ws[c].font = self.normal_font

        headers = ["投标人名称", "投标报价", "启运或运抵时间", "备注"]
        for col, name in zip(("A", "B", "C", "D"), headers):
            ws[f"{col}3"] = name
        self._style_row(ws, 3, 1, 4, header=True)

        ws["A4"] = "中国海外经济合作有限公司"
        ws["B4"] = "='1.投标报价总表'!C9"
        ws["C4"] = self.project.trans_time
        ws["D4"] = ""
        self._style_row(ws, 4, 1, 4, header=False)
        ws["A4"].alignment = self.left
        ws["C4"].alignment = self.left

        ws["C8"] = "投标人盖章："
        ws["C9"] = "日期："
        ws["D9"] = bid_date
        ws["C8"].font = self.normal_font
        ws["C9"].font = self.normal_font
        ws["D9"].font = self.normal_font

    def generate(self, filename: Optional[str] = None) -> str:
        items = self.project.commodities
        bid_date = self._parse_date(self.project.date)

        wb = Workbook()
        ws_all = wb.active
        # ws_fee = wb.create_sheet("费用输入")
        # ws_open = wb.create_sheet("3.开标一览表")
        # ws_total = wb.create_sheet("1.投标报价总表")
        # ws_inner = wb.create_sheet("2.物资对内分项报价表")
        # ws_tax = wb.create_sheet("3.各项物资退抵税额表")
        # ws_tech = wb.create_sheet("4.技术服务费报价表")
        # ws_train = wb.create_sheet("5.来华培训费报价表")
        ws_pick = wb.create_sheet("物资选择")

        self._build_all_suppliers(ws_all, items)
        self._build_selector(ws_pick, items)
        self._build_fee_input(ws_fee)
        inner_total_row = self._build_inner_quote(ws_inner, len(items))
        tax_total_row = self._build_tax_sheet(ws_tax, len(items), inner_total_row)
        self._build_tech_sheet(ws_tech)
        self._build_train_sheet(ws_train)
        self._build_total_sheet(ws_total, inner_total_row, tax_total_row, bid_date)
        self._build_opening_sheet(ws_open, bid_date)

        wb.calculation.fullCalcOnLoad = True
        if not filename:
            filename = f"投标报价表-{self._safe_name(self.project.name)}.xlsx"
        wb.save(filename)
        return filename
