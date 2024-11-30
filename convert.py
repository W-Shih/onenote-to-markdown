import re
import os
import shutil
import sys

import fitz
import win32com.client as win32
import pywintypes
import re
import traceback
from xml.etree import ElementTree

OUTPUT_DIR = os.path.join(os.path.expanduser('~'), "Desktop", "OneNoteExport")
ASSETS_DIR = "assets"
PROCESS_RECYCLE_BIN = False
LOGFILE = 'onenote_to_markdown.log' # Set to None to disable logging
# For debugging purposes, set this variable to limit which pages are exported:
LIMIT_EXPORT = r'' # example: YourNotebook\Notes limits it to the Notes tab/page

def log(message):
    print(message)
    if LOGFILE is not None:
        with open(LOGFILE, 'a', encoding='UTF-8') as lf:
            lf.write(f'{message}\n')

def safe_str(name):
    return re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", name)

def should_handle(path):
    return path.startswith(LIMIT_EXPORT)

def extract_pdf_pictures(pdf_path, assets_path, page_name):
    os.makedirs(assets_path, exist_ok=True)
    image_names = []
    try:
        doc = fitz.open(pdf_path)
    except:
        return []
    img_num = 0
    for i in range(len(doc)):
        for img in doc.get_page_images(i):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            png_name = "%s_%s.png" % (page_name, str(img_num).zfill(3))
            png_name = png_name.replace(' ', '_')
            png_path = os.path.join(assets_path, png_name)
            log("Writing png: %s" % png_path)
            if pix.n < 5:
                pix.save(png_path)
            else:
                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                pix1.save(png_path)
                pix1 = None
            pix = None
            image_names.append(png_name)
            img_num += 1
    return image_names

def fix_image_names(md_path, image_names):
    tmp_path = md_path + '.tmp'
    i = 0
    with open(md_path, 'r', encoding='utf-8') as f_md:
        with open(tmp_path, 'w', encoding='utf-8') as f_tmp:
            body_md = f_md.read()
            for i,name in enumerate(image_names):
                body_md = re.sub("media\/image" + str(i+1) + "\.[a-zA-ZА-Яа-яЁё]+", ASSETS_DIR + "/" + name, body_md)
            f_tmp.write(body_md)
    shutil.move(tmp_path, md_path)

def remove_html_image_dimensions(md_path):
    tmp_path = md_path + '.tmp'
    i = 0
    with open(md_path, 'r', encoding='utf-8') as f_md:
        with open(tmp_path, 'w', encoding='utf-8') as f_tmp:
            body_md = f_md.read()
            body_md = re.sub("style=\"width:[0-9.]+in;height:[0-9.]+in\"", "", body_md)
            f_tmp.write(body_md)
    shutil.move(tmp_path, md_path)

def handle_page(onenote, elem, path, i):
    safe_name = safe_str("%s_%s" % (str(i).zfill(3), elem.attrib['name']))
    if not should_handle(os.path.join(path, safe_name)):
        return

    full_path = os.path.join(OUTPUT_DIR, path)
    os.makedirs(full_path, exist_ok=True)
    path_assets = os.path.join(full_path, ASSETS_DIR)
    safe_path = os.path.join(full_path, safe_name)
    path_docx = safe_path + '.docx'
    path_pdf = safe_path + '.pdf'
    path_md = safe_path + '.md'
    # Remove temp files if exist
    if os.path.exists(path_docx):
        os.remove(path_docx)
    if os.path.exists(path_pdf):
        os.remove(path_pdf)
    try:
        # Create docx
        onenote.Publish(elem.attrib['ID'], path_docx, win32.constants.pfWord, "")
        # Convert docx to markdown
        log("Generating markdown: %s" % path_md)
        os.system('pandoc.exe -i "%s" -o "%s" -f docx -t gfm --wrap=none --markdown-headings=atx' % (path_docx, path_md))
        # Create pdf (for the picture assets)
        onenote.Publish(elem.attrib['ID'], path_pdf, 3, "")
        # Output picture assets to folder
        image_names = extract_pdf_pictures(path_pdf, path_assets, safe_name)
        # Replace image names in markdown file
        fix_image_names(path_md, image_names)
        remove_html_image_dimensions(path_md)
    except pywintypes.com_error as e:
        log("!!WARNING!! Page Failed: %s" % path_md)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    # Clean up docx, html
    if os.path.exists(path_docx):
        os.remove(path_docx)
    if os.path.exists(path_pdf):
        os.remove(path_pdf)

def handle_element(onenote, elem, path='', i=0, target_notebook=None, target_page=None):
    if elem.tag.endswith('Notebook'):
        notebook_name = elem.attrib['name']
        if target_notebook and notebook_name != target_notebook:
            return
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsChildren, "")
        for i,c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i, target_notebook, target_page)
    elif elem.tag.endswith('Section'):
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsPages, "")
        for i,c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i, target_notebook, target_page)
    elif elem.tag.endswith('SectionGroup') and (not elem.attrib['name'].startswith('OneNote_RecycleBin') or PROCESS_RECYCLE_BIN):
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsSections, "")
        for i,c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i, target_notebook, target_page)
    elif elem.tag.endswith('Page'):
        if target_page and elem.attrib['name'] != target_page:
            return
        try:
            handle_page(onenote, elem, path, i)
        except:
            print("Page failed unexpectedly: %s" % path, file=sys.stderr)

if __name__ == "__main__":
    try:
        target_notebook, target_page = None, None
        if len(sys.argv) > 1:
            target_notebook = sys.argv[1]
        if len(sys.argv) > 2:
            target_page = sys.argv[2]

        onenote = win32.gencache.EnsureDispatch("OneNote.Application.12")
        hier = onenote.GetHierarchy("", win32.constants.hsNotebooks, "")
        root = ElementTree.fromstring(hier)
        for child in root:
            handle_element(onenote, child, target_notebook=target_notebook, target_page=target_page)
    except pywintypes.com_error as e:
        traceback.print_exc()
        log("!!!Error!!! Hint: Make sure OneNote is open first.")
