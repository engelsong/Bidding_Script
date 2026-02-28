from project import Project
from directory import Directory
from cover import Cover
from content import Content
from os import listdir
import re


doc_pattern = re.compile('^project.*\.docx$')
doc_name = None
for doc in listdir():
    if re.match(doc_pattern, doc):
        doc_name = doc
if doc_name is None:
    raise FileNotFoundError('No project*.docx file found in current directory.')

project = Project(doc_name)
# project.show_info()
directory = Directory(project)
try:
    directory.make_dir()
except FileExistsError:
    pass

content = Content(project)
content.generate_content()

# cover = Cover(project)
# cover.generate()
