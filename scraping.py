import requests
from lxml import html
from urllib.parse import urlparse
import re
import os
import zipfile
import shutil
from glob import glob
import logging
import textract
import win32com.client
import tempfile
import json
import pyunpack
import traceback

import PIL
import pytesseract 
import pdf2image

import sys
import pathlib

repo_dir = pathlib.Path(__file__).parent.absolute()

SETTINGS_PATH = os.path.join(repo_dir, 'settings.json')
QUERIES_TO_LOTS_FILENAME = 'queries_to_lots.json'
LOTS_TO_NAMES_FILENAME = 'lots_to_names.json'
ZIP_SUBDIR = 'zip'
UNZIPPED_SUBDIR = 'unzipped'
TXT_SUBDIR = 'txt'
QUERY_SUBDIR = 'query'

def read_object(path):
    with open(path) as f:
        return json.load(f)
    
def write_object(obj, path):
    with open(path, 'w') as f:
        json.dump(obj, f, indent=4)

def write_content(response, filepath):
    with open(filepath, 'wb') as output_file:
        output_file.write(response.content)
        
def get_filename(response):
    cd = response.headers.get('content-disposition')
    return re.findall('filename\*?=(UTF-8\'\')?(?P<filename>.+)', cd)[0][1].strip('"')

def rm_empty_dirs(path):
    if not os.path.isdir(path):
        return
    for f in os.listdir(path):
        fullpath = os.path.join(path, f)
        if os.path.isdir(fullpath):
            rm_empty_dirs(fullpath)
    if len(os.listdir(path)) == 0:
        os.rmdir(path)

def unzip(filepath, rm_archive):
    unzip_folder = os.path.splitext(filepath)[0]
    with zipfile.ZipFile(filepath, 'r') as z:
        for name in z.namelist():
            cur_path = z.extract(name, unzip_folder)
            try:
                better_name = os.path.join(unzip_folder, name.encode('cp437').decode('cp866'))
                os.rename(cur_path, os.path.join(unzip_folder, better_name))
            except:
                pass
    if rm_archive:
        os.remove(filepath)
    return unzip_folder

def unrar(filepath, rm_archive):
    unzip_folder = os.path.splitext(filepath)[0]
    try_makedirs(unzip_folder)
    pyunpack.Archive(filepath).extractall(unzip_folder)
    if rm_archive:
        os.remove(filepath)
    return unzip_folder
        
def unzip_recursive(root, rm_archive):
    created_folders = []
    ext = os.path.splitext(root)[1]
    if ext == '.zip':
        created_folders.append(unzip(root, rm_archive=rm_archive))
    elif ext == '.rar':
        created_folders.append(unrar(root, rm_archive=rm_archive))
    elif os.path.isdir(root):
        folders = [i for i in glob(os.path.join(root, '*')) if os.path.isdir(i)]
        for i in folders:
            created_folders += unzip_recursive(i, rm_archive=rm_archive)            
        zip_files = glob(os.path.join(root, '*.zip')) + glob(os.path.join(root, '*.rar'))
        for i in zip_files:
            created_folders += unzip_recursive(i, rm_archive=rm_archive)
    for i in created_folders:
        unzip_recursive(i, rm_archive=rm_archive)
    rm_empty_dirs(root)
    return created_folders

def convert_to_docx(src, dst):
    """Library textract does not support .doc correctly, so convert them"""
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(src)
    doc.Activate()
    word.ActiveDocument.SaveAs(dst, FileFormat=win32com.client.constants.wdFormatXMLDocument)
    doc.Close(False)
    word.Quit()

def doc_to_text(path):
    with tempfile.TemporaryDirectory(prefix='advanced_search_') as tmp_dir:
        dst = os.path.join(tmp_dir, 'converted.docx')
        convert_to_docx(path, dst)
        content = textract.process(dst)
    return content

def pdf_to_text(pdf_path):
    try:
        return pdf_to_text_ocr(pdf_path)
    except pytesseract.TesseractError:
        return textract.process(pdf_path)

def fix_line_endings(text):
    text = text.replace('\r\n', '\n')
    while ('\n\n' in text):
        text = text.replace('\n\n', '\n')
    return text

def pdf_to_text_ocr(pdf_path):
    pages = pdf2image.convert_from_path(pdf_path, 300)
    logging.info(u'Обработка {} ({} страниц) с использованием оптического распознавания символов'.format(pdf_path, len(pages)))
    texts = []
    with tempfile.TemporaryDirectory(prefix='advanced_search_') as tmp_dir:
        for page in pages:
            filepath = os.path.join(tmp_dir, "page.jpg")
            degrees = pytesseract.image_to_osd(page, output_type=pytesseract.Output.DICT)['orientation']
            if degrees != 0:
                page = page.rotate(degrees, expand=True)
            page.save(filepath, 'JPEG')
            text = pytesseract.image_to_string(PIL.Image.open(filepath), lang='rus')
            texts.append(text)
    return '\n'.join(texts)

def any_file_to_str(path):
    ext = os.path.splitext(path)[-1]
    converters = {'.pdf': pdf_to_text,
                  '.doc': doc_to_text}
    converter = converters[ext] if ext in converters else textract.process
    result = converter(path)
    if isinstance(result, bytes):
        result = result.decode('utf-8')
    result = fix_line_endings(result)
    return result
    
def try_makedirs(path):
    try:
        os.makedirs(path)
    except OSError:
        pass


def convert_to_txt_recursively(root_src, root_dst, extentions):
    logging.info(u'Конвертация {} в текстовый формат и сохранение в {}'.format(root_src, root_dst))
    problem_files = []
    converted_files = []
    ignored_files = []
    for root, dirs, files in os.walk(root_src):
        for f in files:
            cur_path = os.path.join(root, f)
            if os.path.splitext(cur_path)[-1] not in extentions:
                ignored_files.append(cur_path)
                break
            new_path = cur_path.replace(root_src, root_dst, 1) + '.txt'
            try:
                content = any_file_to_str(cur_path)
                converted_files.append(cur_path)
            except KeyboardInterrupt:
                sys.exit('KeyboardInterrupt')
            except:
                logging.debug(traceback.format_exc())
                problem_files.append(cur_path)
                break
            new_dir = os.path.dirname(new_path)
            try_makedirs(new_dir)
            with open(new_path, 'w', encoding='utf-8') as f:
                f.write(content)
    return dict(converted=converted_files, igonored=ignored_files, problem=problem_files)

def convert_to_txt_wrapper(root_src, root_dst, extentions=('.xls', '.xlsx', '.doc', '.docx', '.pdf', '.txt')):
    if not os.path.isdir(root_dst) or not os.listdir(root_dst):
        categories = convert_to_txt_recursively(root_src, root_dst, extentions)
        if categories['problem']:
            logging.warning(u'Проблема при конвертации в текст:\n{}'.format(categories['problem']))
    else:
        logging.info(u'Пропускаем конвертацию в текстовый формат для {}:\nДиректория {} уже не пуста'.format(root_src, root_dst))
        

def get_url_root(url):
    return '{uri.scheme}://{uri.netloc}/'.format(uri=urlparse(url))

def get_tree(url):
    response = requests.get(url)
    tree = html.fromstring(response.text)
    tree.make_links_absolute(get_url_root(url))
    return tree

def get_list_of_lots(query_url):
    logging.info(u'Выполнение запроса:\n{}'.format(query_url))
    tree = get_tree(query_url)
    title_nodes = tree.xpath('//a[@class="section-procurement__item-title" and @href]')
    lots = [node.attrib['href'] for node in title_nodes]
    return lots

def get_list_of_lots_cached(query_url, cache_file):
    query_to_lots_mapping = dict()
    if os.path.isfile(cache_file):
        query_to_lots_mapping = read_object(cache_file)
        if query_url in query_to_lots_mapping:
            logging.info(u'Необходимая информация найдена в кэше, запрос {} не будет выполнен'.format(query_url))
            return query_to_lots_mapping[query_url]
    list_of_lots = get_list_of_lots(query_url)
    query_to_lots_mapping[query_url] = list_of_lots
    write_object(query_to_lots_mapping, cache_file)    
    return list_of_lots
    
def download_file(link, directory):
    logging.info(u'Скачивание файла по ссылке {}'.format(link))
    response = requests.get(link)
    fname = get_filename(response)
    path = os.path.join(directory, fname)
    write_content(response, path)
    return path

def download_files(lot_url, storage_dir, one_by_one, force):
    f = download_files_one_by_one if one_by_one else download_files_in_zip
    if force or not os.path.isdir(storage_dir) or not os.listdir(storage_dir):
        try_makedirs(storage_dir)
        tree = get_tree(lot_url)
        f(tree, storage_dir)
    else:
        logging.info(u'Пропускаем загрузку файлов для {}:\nДиректория {} уже не пуста'.format(lot_url, storage_dir))
        
def download_files_in_zip(tree, storage_dir):
    download_link = tree.xpath('//a[@class="downloadDocument btn procedure__lot-button" and @href]')[0]
    link = download_link.attrib['href']
    download_file(link, storage_dir)

def download_files_one_by_one(tree, storage_dir):
    links = tree.xpath('//div[@class="item-name"]/a[@href]')
    links = [i.attrib['href'] for i in links]
    for i in links:
        download_file(i, storage_dir)

def get_lot_id(lot_url):
    return lot_url.split('/')[-1]

def get_lot_name(lot_tree):
    links = lot_tree.xpath('//span[@class="procedure__item-name"]')
    return links[0].text

def get_lot_name_cached(lot_url, workdir):
    url_to_name = dict()
    path = os.path.join(workdir, LOTS_TO_NAMES_FILENAME)
    if os.path.isfile(path):
        url_to_name = read_object(path)
    if lot_url not in url_to_name:
        lot_name = get_lot_name(get_tree(lot_url))
        url_to_name[lot_url] = lot_name
        write_object(url_to_name, path)
    return url_to_name[lot_url]

def unzip_recursive_wrapper(src, dst):
    if os.path.isdir(dst) and os.listdir(dst):
        logging.info(u'Пропускаем разархивирование для {}:\nДиректория {} уже не пуста'.format(src, dst))
        return
    logging.info(u'Разархивирование {} в {}'.format(src, dst))
    shutil.copytree(src, dst, dirs_exist_ok=True)
    unzip_recursive(dst, rm_archive=True)
    
def create_filename_suffix(lot_name):
    invalid = r'<>:"/\|?*'
    for i in invalid:
        lot_name = lot_name.replace(i, '')
    lot_name = lot_name.replace('\n', '')
    lot_name = re.sub(' +', ' ', lot_name)
    return lot_name[:50].strip()

def get_subdirs(lot_url, root, suffix=None):
    if suffix is None:
        suffix = get_lot_id(lot_url)
    lot_zipped_subdir = os.path.join(root, ZIP_SUBDIR, suffix)
    lot_unzipped_subdir = os.path.join(root, UNZIPPED_SUBDIR, suffix)
    lot_txt_subdir = os.path.join(root, TXT_SUBDIR, suffix)
    return lot_zipped_subdir, lot_unzipped_subdir, lot_txt_subdir

def get_symlinks(lot_url, workdir, root):
    lot_id = get_lot_id(lot_url)
    lot_name = get_lot_name_cached(lot_url, workdir)
    suffix = '{} {}'.format(lot_id, create_filename_suffix(lot_name))
    return get_subdirs(lot_url, root=root, suffix=suffix)

def create_query_subdir(query_url, workdir):
    query_subdir = os.path.join(workdir, QUERY_SUBDIR)
    logging.info(u'Копирование релевантных лотов для запроса {} в папку {}'.format(query_url, query_subdir))
    list_of_lots = get_list_of_lots_cached(query_url, os.path.join(workdir, QUERIES_TO_LOTS_FILENAME))
    if os.path.isdir(query_subdir):
        try:
            shutil.rmtree(query_subdir)
        except:
            print(traceback.format_exc())
            print(u'Не удалось очистить директорию {}\nЗакройте файлы, хранящиеся в ней, или удалите её вручную, после чего перезапустите скрипт'.format(query_subdir))
            sys.exit(1)
    for url in list_of_lots:
        subdirs = get_subdirs(url, workdir)
        symlinks = get_symlinks(url, workdir, query_subdir)
        for target, link in zip(subdirs, symlinks):
            try_makedirs(os.path.dirname(link))
            #  os.symlink(link, target, True)  # Admin rights needed, so just copy:
            if os.path.isdir(target):
                shutil.copytree(target, link)
    return query_subdir
            

def process_query(query_url, workdir):
    try_makedirs(workdir)
    queries_to_lots_filepath = os.path.join(workdir, QUERIES_TO_LOTS_FILENAME)
    list_of_lots = get_list_of_lots_cached(query_url, queries_to_lots_filepath)
    logging.info(u'Найдено {} релевантных данному запросу лотов'.format(len(list_of_lots)))
    for url in list_of_lots:
        try:
            lot_zipped_subdir, lot_unzipped_subdir, lot_txt_subdir = get_subdirs(url, workdir)
            get_lot_name_cached(url, workdir)
            download_files(url, lot_zipped_subdir, one_by_one=False, force=False)
            unzip_recursive_wrapper(lot_zipped_subdir, lot_unzipped_subdir)
            convert_to_txt_wrapper(lot_unzipped_subdir, lot_txt_subdir)
        except:
            logging.warning(u'Возникла проблема при обработке {}'.format(url))
            logging.debug(traceback.format_exc())
    return create_query_subdir(query_url, workdir)

_example_of_valid_url = 'https://www.tektorg.ru/procedures?q=%D0%A3%D0%B7%D0%B5%D0%BB+%D1%83%D1%87%D0%B5%D1%82%D0%B0+%D0%BD%D0%B5%D1%84%D1%82%D0%B8'

def input_url_is_valid(url):
    if not url.startswith('https://www.tektorg.ru/procedures?q='):
        print('Invalid url. Example of valid url:\n{}'.format(_example_of_valid_url))
        return False
    return True



if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Downloads and processes data from tektorg.ru')
    parser.add_argument('--logging', default='INFO', choices=('DEBUG', 'INFO', 'WARNING', 'ERROR'))
    parser.add_argument('url', help='URL of query. Copy and paste from here:\nhttps://www.tektorg.ru/procedures')
    args = parser.parse_args()

    settings = read_object(SETTINGS_PATH)
    workdir = os.path.expanduser(settings['workdir'])
    tesseract_path = os.path.expanduser(settings['tesseract_path'])
    tesseract_exe = os.path.join(tesseract_path, 'tesseract.exe')
    if not os.path.isdir(tesseract_path) or not os.path.isfile(tesseract_exe):
        sys.exit('tesseract_path is incorrect. Change it in settings.json')
    pytesseract.pytesseract.tesseract_cmd = tesseract_exe
    sys.path.insert(0, tesseract_path)
    logging.basicConfig(level=getattr(logging, args.logging), format='%(message)s')
    if input_url_is_valid(args.url):
        query_subdir = process_query(args.url, workdir)
        print('\nDone. Output:\n{}\n'.format(query_subdir))
