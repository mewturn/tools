from simple_idml import idml
from lxml import etree
import os

def unzip(filename, directory=""):
    import zipfile
    if directory == "":
        directory = filename.split(".")[0]
    with zipfile.ZipFile(filename, "r") as ref:
        print(f"Extracting contents to directory {directory}")
        ref.extractall(directory)
        
    return directory

def parse_xml(filename, last_count=0):
    # print(f"Processing {filename}")
    import xml.etree.ElementTree as ET
    tree = ET.parse(filename)
    root = tree.getroot()
    
    prefix = 'aaa'
    
    contents = root.findall(".//Content")
    for content in contents:
        if content.text is not None:
            print(content.text)
            content.set("updated", "yes")
            content.text = f"{prefix}[{last_count}]"
            last_count += 1

    tree.write(filename)
    return last_count

def parse_text_from_idml(idml_filename):
    directory = unzip(idml_filename)
    idml_main = idml.IDMLPackage(idml_filename)
    last_count = 0

    for story_filename in idml_main.stories:
        story_filename = f"{directory}/{story_filename}"
        last_count = parse_xml(story_filename, last_count)
