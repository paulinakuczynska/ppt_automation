#!/usr/bin/env python3

import argparse, os, shutil
from xml.etree.ElementTree import Element, SubElement
import xml.etree.cElementTree as ET
from re import match

# params

parser = argparse.ArgumentParser(
    allow_abbrev=False, 
    description="""
    You can add from one to fifty custom colors
    Value -v is a hex code of a color e.g. 20f00f
    Name -n is a color name which is displayed in PPT e.g. Blue
    Number of values and names must be equal, min 1, max 50
    """
    )

parser.add_argument("-v", "--value", required=True, nargs="+", type=str, help="hex values")
parser.add_argument("-n", "--name", required=True, nargs="+", type=str, help="name to display in PPT")
args = vars(parser.parse_args())

filtered_values = list(filter(lambda v: match('[a-zA-Z0-9]{6,6}$', v), args["value"]))

if not filtered_values or len(filtered_values) != len(args["name"]) or len(filtered_values) >= 51:
    print("""
    Inappropriate format:
    numbers of values and names must be equal
    values must be hex format e.g. 20f00f
    max 50 colors can be added
    """)
    quit()

# xml modification

main_path = '/home/alona/ProcessAutomation/testfolder/'
end_path = '/home/alona/ProcessAutomation/testfolder/new/'
xml_file = '/home/alona/ProcessAutomation/testfolder/new/ppt/theme/theme1.xml'
tag_color_list = '{http://schemas.openxmlformats.org/drawingml/2006/main}custClrLst'

os.rename(f'{main_path}template.pptx', f'{main_path}template.zip')
shutil.unpack_archive(f'{main_path}template.zip', end_path, 'zip')

def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        ET.register_namespace(ns, namespaces[ns])

register_all_namespaces(xml_file)

# def remove_namespace(doc, namespace):
#     ns = u'{%s}' % namespace
#     nsl = len(ns)
#     for elem in tree.iter(doc):
#         if elem.tag.startswith(ns):
#             elem.tag = elem.tag[nsl:]

tree = ET.parse(xml_file)
root = tree.getroot()

# remove_namespace(root, u'http://schemas.microsoft.com/office/thememl/2012/main')

# result = []
for elem in root:
    if elem.tag == tag_color_list:
        elem.clear()
        # for subelem in elem:
        #     result.append(elem)

if root.find(tag_color_list) is None:
    color_main_element = ET.Element(tag_color_list)
    root.insert(3, color_main_element)

color_list = root.find(tag_color_list)
for n, v in zip(args["name"], args["value"]):
    color_name = ET.SubElement(color_list, 'a:custClr')
    color_value = ET.SubElement(color_name, 'a:srgbClr')
    color_name.set('name', n)
    color_value.set('val', v)

tree.write(xml_file,
            xml_declaration=True,
            encoding='utf-8',
            method='xml')

shutil.make_archive(f'{main_path}zrobiony', 'zip', end_path)
os.rename(f'{main_path}zrobiony.zip', f'{main_path}zrobiony.pptx')

print('Tadam!')