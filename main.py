from project import Project
from directory import Directory
from cover import Cover
from os import listdir
import re


doc_pattern = re.compile('^project.*\.docx$')
doc_name = None
for doc in listdir():
    if re.match(doc_pattern, doc):
        doc_name = doc
project = Project(doc_name)
# project.show_info()
directory = Directory(project)
directory.make_dir()

# cover = Cover(project)
# cover.generate()
