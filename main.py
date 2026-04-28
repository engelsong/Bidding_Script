from __future__ import annotations

import re
from os import listdir
from os.path import exists
from typing import Iterable, List, Optional

from content import Content
from cover import Cover
from directory import Directory
from project import Project
from quotation import Quotation
from separate import Separate


MENU_TEXT = """
请按照序号选择你需要的功能：
1、生成报价表
2、生成目录
3、生成封面
4、生成空白本文件夹结构
5、拆分报价表
>>> """

VALID_OPTIONS = "12345"


def find_project_doc() -> str:
    doc_pattern = re.compile(r"^project.*\.docx$")
    doc_name = None
    for doc in listdir():
        if re.match(doc_pattern, doc):
            doc_name = doc
    if doc_name is None:
        raise FileNotFoundError("当前目录未找到 project*.docx 文件。")
    return doc_name


def parse_menu_selection(raw: str) -> List[str]:
    selected = []
    for ch in raw.replace(" ", ""):
        if ch not in VALID_OPTIONS:
            raise ValueError("您输入了非法指令，请重新输入。")
        if ch not in selected:
            selected.append(ch)
    if not selected:
        raise ValueError("您输入的指令为空，请重新输入。")
    return selected


def prompt_menu_selection() -> List[str]:
    while True:
        try:
            return parse_menu_selection(input(MENU_TEXT))
        except ValueError as exc:
            print(f"<<< {exc} >>>")


def prompt_yes_no(message: str) -> bool:
    while True:
        answer = input(message).strip()
        if answer in {"Y", "y"}:
            return True
        if answer in {"N", "n"}:
            return False
        print("<<< 请输入 Y 或 N >>>")


def build_project_context(selected: Iterable[str]) -> Optional[Project]:
    if not any(option in {"1", "2", "3", "4"} for option in selected):
        return None
    return Project(find_project_doc())
1

def quotation_filename(project: Project) -> str:
    safe_name = Quotation._safe_name(project.name)
    return f"投标报价表-{safe_name}.xlsx"


def content_filename(project: Project) -> str:
    safe_name = Content._safe_name(project.name)
    return f"目录-{safe_name}.xlsx"


def cover_filename(project: Project) -> str:
    return f"封面-{project.name}.docx"


def directory_root(project: Project) -> str:
    return f"投标文件-{project.name}"


def run_selected_actions(selected: List[str]) -> None:
    project = build_project_context(selected)
    generated_quotation: Optional[str] = None

    quotation = Quotation(project) if project is not None else None
    content = Content(project) if project is not None else None
    cover = Cover(project) if project is not None else None
    directory = Directory(project) if project is not None else None
    separate = Separate()

    for option in selected:
        if option == "1":
            target = quotation_filename(project)
            if exists(target) and not prompt_yes_no(f"!!! {target} 已存在，是否覆盖（Y/N）>>> "):
                print(f"<<< 已跳过：{target} >>>")
                continue
            generated_quotation = quotation.generate()
            print(f"<<< 已生成报价表：{generated_quotation} >>>")

        elif option == "2":
            target = content_filename(project)
            if exists(target) and not prompt_yes_no(f"!!! {target} 已存在，是否覆盖（Y/N）>>> "):
                print(f"<<< 已跳过：{target} >>>")
                continue
            output = content.generate_content()
            print(f"<<< 已生成目录：{output} >>>")

        elif option == "3":
            target = cover_filename(project)
            if exists(target) and not prompt_yes_no(f"!!! {target} 已存在，是否覆盖（Y/N）>>> "):
                print(f"<<< 已跳过：{target} >>>")
                continue
            cover.generate()
            print(f"<<< 已生成封面：{target} >>>")

        elif option == "4":
            target = directory_root(project)
            if exists(target) and not prompt_yes_no(f"!!! {target} 已存在，是否继续创建目录（Y/N）>>> "):
                print(f"<<< 已跳过：{target} >>>")
                continue
            try:
                directory.make_dir()
                print(f"<<< 已创建目录结构：{target} >>>")
            except FileExistsError:
                print(f"<<< 目录结构已存在，未重复创建：{target} >>>")

        elif option == "5":
            workbook_name = generated_quotation
            if workbook_name is None:
                workbook_name = separate._find_workbook_name()
            output_files = separate.generate(workbook_name)
            if output_files:
                print(f"<<< 已拆分报价表，共生成 {len(output_files)} 个文件 >>>")
            else:
                print("<<< 未找到可拆分的报价表工作表 >>>")


def main() -> None:
    try:
        selected = prompt_menu_selection()
        run_selected_actions(selected)
    except Exception as exc:
        print(f"<<< 出现异常：{exc} >>>")
    finally:
        input("<<< 程序已经运行完成，按回车键退出 >>>")


if __name__ == "__main__":
    main()
