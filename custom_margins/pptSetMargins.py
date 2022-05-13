#!/usr/bin/env python3

import argparse, os, shutil
import xml.etree.cElementTree as ET
from re import match
from pathlib import Path

parser = argparse.ArgumentParser(
    allow_abbrev=False, 
    description="File name has to be 'todo.pptx' located in the Desktop"
    )

parser.add_argument("-l", "--left", required=False, nargs="?", default=0, type=float, help="left margin")
parser.add_argument("-r", "--right", required=False, nargs="?", default=0, type=float, help="right margin")
parser.add_argument("-t", "--top", required=False, nargs="?", default=0, type=float, help="top margin")
parser.add_argument("-b", "--bottom", required=False, nargs="?", default=0, type=float, help="bottom margin")
args = vars(parser.parse_args())

main_path = Path.home() / Path('Desktop')
end_path = main_path / Path('margins')
todo_pptx = main_path / Path('todo.pptx')
todo_zip = main_path / Path('todo.zip')
xml_path = end_path / Path('ppt', 'slideLayouts')
xml_file = end_path / Path('ppt', 'slideLayouts', 'slideLayout2.xml')
custom = main_path / Path('custommargins')
custom_zip = main_path / Path('custommargins.zip')
custom_pptx = main_path / Path('custommargins.pptx')
tag1 = '{http://schemas.openxmlformats.org/presentationml/2006/main}cSld'
tag2 = '{http://schemas.openxmlformats.org/presentationml/2006/main}spTree'
tag3 = '{http://schemas.openxmlformats.org/presentationml/2006/main}sp'
tag4 = '{http://schemas.openxmlformats.org/presentationml/2006/main}txBody'
tag5 = '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr'

os.rename(todo_pptx, todo_zip)
shutil.unpack_archive(todo_zip, end_path, 'zip')

def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        ET.register_namespace(ns, namespaces[ns])
register_all_namespaces(xml_file)

for file in os.listdir(xml_path):
    if not file.endswith('.xml'): continue
    fullname = xml_path / Path(file)
    tree = ET.parse(fullname)
    root = tree.getroot()
    
    all_bodyPr = root.findall(f'{tag1}/{tag2}/{tag3}/{tag4}/{tag5}')
    
    for elem in all_bodyPr:
        elem.set('lIns', str(args["left"]))
        elem.set('tIns', str(args["top"]))
        elem.set('rIns', str(args["right"]))
        elem.set('bIns', str(args["bottom"]))
    
    tree.write(fullname,
            xml_declaration=True,
            encoding='utf-8',
            method='xml')

shutil.make_archive(custom, 'zip', end_path)
os.rename(custom_zip, custom_pptx)

print('Tadam!')