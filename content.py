from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.worksheet.page import PageMargins


class Content:
    """
    通过project实例创建投标文件目录1
    """
    
    # 设置公用样式
    title_font = Font(name='宋体', size=24, bold=True)
    header_font = Font(name='仿宋_GB2312', size=14, bold=True)
    normal_font = Font(name='仿宋_GB2312', size=14)
    header_border = Border(bottom=Side(style='medium'))
    normal_border = Border(bottom=Side(style='thin', color='80969696'))
    ctr_alignment = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True)
    left_alignment = Alignment(
        horizontal='left',
        vertical='center',
        wrap_text=True,
        indent=1)
    margin = PageMargins()


    def __init__(self, project) -> None:
        self.project = project