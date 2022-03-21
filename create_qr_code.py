import pandas as pd
import os
import re
import pyqrcode
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pytesseract as pt
from PyPDF2 import PdfFileReader, PdfFileWriter
# from PIL import Image
from pdf2image import convert_from_path
import shutil
from docx2pdf import convert

master_list = pd.read_excel('D:/Shweta/data_digitization/reference_docs/2022_03_19_patient_master_list_sk.xlsx')
categorised_excel = pd.read_excel('D:/Shweta/data_digitization/reference_docs/2010_file_categorization_excel.xlsx')

def change_sep(string, old_sep, new_sep):
    """ change the separator between the string or within the string
    :param string: string
    :param old_sep: string separators (' ', '_', '/')
    :param new_sep: string separators (' ', '_', '/')
    :return: string with changed separators
    """
    changed_sep = re.sub(old_sep, new_sep, str(string))
    return changed_sep

id_cols = ['mr_number', 'patient_name', 'dob']

def get_id_data(master_list, file_number, id_cols):
    """
    get id values from input id names and single row of master list
    :param master_list: pd.DataFrame
    :param file_number: file_number separated by '_'
    :param id_cols: col-names which stores the patients identifing info(file_number, mr_number, name, dob)
    :return: list of id data for single row of master list
    """
    id_data = master_list[master_list['file_number'] == file_number]
    id_data = id_data[id_cols]
    return id_data

folder_col_heads = ['report_name', 'subfolder_name']

def get_folder_subfolder(categorized_excel, index):
    """
    it will give the folder name and sub-folder name for the qr code
    :param categorized_excel: pd.DataFrame categorized excel
    :param index: integer
    :return:
    """
    folder_dat = []
    for col_name in folder_col_heads:
        folder_info = categorized_excel[col_name][index]
        folder_dat.append(folder_info)
    return folder_dat

report_types_dic = {'1': 'Patient Information', '2': 'Clinical Examination',
                    '3': 'Radiology', '4': 'Metastatic Examination',
                    '5': 'Biopsy Pathology', '6': 'Neo-Adjuvant Chemotherapy',
                    '7':'Surgical Procedures', '8': 'Patient Images',
                    '9':'Surgery Media', '10': 'Surgery Pathology',
                    '11': 'Chemotherapy', '12': 'Radiotherapy',
                    '13': 'Follow-up Notes', '14': 'Genetics',
                    '15': 'Miscellaneous', '16': 'Patient File Data',
                    '17': 'PROMS'}

def get_data_for_file_number(file_number, categorized_excel):
    grouped_data = categorized_excel[categorised_excel['file_number']==file_number]
    return grouped_data

id_dat = get_id_data(master_list,'38_10', id_cols)

def make_qr_code(master_list, categorized_excel, qr_destination_path):
    for i in range(len(categorized_excel)):
        file_number = categorized_excel['file_number'][i]
        file_number_str = change_sep(file_number, '_', '/')
        folder_name = categorized_excel['report_name'][i]
        subfolder = categorized_excel['subfolder_name'][i]
        id_dat = get_id_data(master_list, file_number, id_cols)
        mr_number = id_dat['mr_number'][0]
        if subfolder is not None:
            qr_code = file_number_str + '_' + str(mr_number) + '_' + str(folder_name) + '_' + str(subfolder)
            qr = pyqrcode.create(qr_code)
            report_type_for_name = change_sep(str(folder_name), ' ', '_')
            subfolder_for_name = change_sep(str(subfolder), ' ', '_')
            qr_img_name = file_number + '_' + str(mr_number) + '_' + report_type_for_name + '_' + str(subfolder_for_name) + '.png'
            qr_path = os.path.join(qr_destination_path, qr_img_name)
            qr.png(qr_path, scale=4)
            print('QR code created for ' + file_number + ' ' + folder_name + ' ')
        else:
            qr_code = file_number_str + '_' + str(mr_number) + '_' + str(folder_name)
            qr = pyqrcode.create(qr_code)
            report_type_for_name = re.sub(' ', '_', str(folder_name))
            qr_img_name = file_number + '_' + str(mr_number) + '_' + report_type_for_name + '.png'
            qr_path = os.path.join(qr_destination_path, qr_img_name)
            qr.png(qr_path, scale=4)
            print('QR code created for ' + file_number + ' ' + folder_name + ' ')

def format_word_doc(doc, id_value):
    text = doc.add_paragraph()
    report_type_name = text.add_run(str(id_value))
    report_type_name.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    report_type_name.bold = True
    report_type_name.font.size = Pt(28)
    report_type_name.font.name = 'Arial Black'

def add_qr_code_in_word_document(qr_code_path, master_list, categorised_excel):
    doc = Document()
    doc.add_picture(qr_code_path)
    qr_alignment = doc.paragraphs[-1]
    qr_alignment.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for row, report_type in enumerate(master_list['report_name']):
        format_word_doc(doc, report_type)

