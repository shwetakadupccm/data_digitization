# from numpy import source
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
from file_functions import HelperFunctions


class QrCode():

    def __init__(self, root, master_list_name, categorized_excel):
        self.root = root
        self.master_list_name = master_list_name
        self.categorized_excel = categorized_excel
        self.hf = HelperFunctions(
            self.root, self.master_list_name, self.categorized_excel)

    # def function_params(self, param_name):
        # param_dict = dict(master_list=pd.read_excel(os.path.join(self.root, 'reference_docs', self.master_list_name)),
        #   categorized_excel=pd.read_excel(os.path.join(
        #   self.root, 'reference_docs', self.categorized_excel)),
        #   id_cols=['mr_number',
        #    'patient_name', 'date_of_birth'],
        #   qr_data_cols=['File Number', 'MR Number',
        # 'Patient Name', 'Date of Birth'],
        #   folder_col_heads=['report_name', 'subfolder_name'],
        #   report_types_dict=['Patient Information', 'Clinical Examination',
        #  'Radiology', 'Metastatic Examination',
        #  'Biopsy Pathology', 'Neo-Adjuvant Chemotherapy',
        #  'Surgical Procedures', 'Patient Images',
        #  'Surgery Media', 'Surgery Pathology',
        #  'Chemotherapy', 'Radiotherapy',
        #  'Follow-up Notes', 'Genetics',
        #  'Miscellaneous', 'Patient File Data',
        #  'PROMS'])
        # return param_dict.get(param_name)

    def add_qr_code_in_word_document(self, category_row):
        qr_code_dat = self.make_qr_code(category_row)  # add input params
        doc_dat = qr_code_dat.split("_")
        # file_number, mr_number, patient_name, date_of_birth = doc_dat[0:4]
        doc = Document()
        doc.add_picture(os.path.join(self.root, 'tmp\qr_codes\qr_img.png'))
        qr_alignment = doc.paragraphs[-1]
        qr_alignment.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        folders = doc_dat[4:]
        # id_texts = [', '.join(doc_dat[0:2]), patient_name, date_of_birth]
        # should work? - working
        folder = [folder for folder in folders if folder is not None]
        report_type = ' '.join(id for id in folder)
        doc = self.hf.format_word_doc(report_type, doc)
        blank_para = doc.add_paragraph()
        run = blank_para.add_run()
        run.add_break()
        # todo for loop - done
        prefixes = self.hf.function_params('qr_data_cols')
        for idx, prefix in enumerate(prefixes):
            doc = self.hf.format_word_doc(
                str(prefix) + ': ' + str(doc_dat[idx]), doc)
        # doc_name = qr_code_dat.replace("/", "_")
        # doc_name = doc_name.replace(" ", '_') + '.docx'
        doc_name = 'coded_file.docx'
        # coded_data_dir = self.hf.create_folder_for_data_type(data_type=)
        doc_folder = doc_name.removesuffix('.docx')
        doc_path = self.hf.create_folder_for_data_type(
            source_folder='coded_data', data_type=doc_name)
        doc.save(doc_path)
        pdf_name = doc_name.replace('.docx', '.pdf')
        pdf_path = self.hf.create_folder_for_data_type(
            data_type=pdf_name, source_folder='coded_data')
        convert(doc_path, pdf_path)
        return pdf_path

    @ classmethod
    def make_patient_full_name(self, master_list):
        first_name_col_idx = master_list.columns.get_loc('first_name')
        last_name_col_idx = master_list.columns.get_loc('last_name')
        master_list['patient_name'] = master_list[master_list.columns[first_name_col_idx:last_name_col_idx]].apply(
            lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        master_list['patient_name'] = master_list['patient_name'].str.title()
        return master_list

    def get_id_data(self, file_number):
        """
        get id values from input id names and single row of master list
        :param master_list: pd.DataFrame
        :param file_number: file_number separated by '_'
        :param id_cols: col-names which stores the patients identifing info(file_number, mr_number, name, dob)
        :return: list of id data for single row of master list
        """
        master_list = self.hf.function_params('master_list')
        master_list = self.make_patient_full_name(master_list)
        id_cols = self.hf.function_params('id_cols')
        msater_list = self.hf.function_params('master_list')
        id_data = master_list['file_number'] == file_number][id_cols]  # type: ignore
        # id_data =id_data

        id_data_lst = id_data.values.tolist()
        id_data = [id_text for sublist in id_data_lst for id_text in sublist]
        return id_data

        # def change_sep(string, old_sep, new_sep):
        # """ change the separator between the string or within the string
        # :param string: string
        # :param old_sep: string separators (' ', '_', '/')
        # :param new_sep: string separators (' ', '_', '/')
        # :return: string with changed separators
        # """
        # changed_sep = re.sub(old_sep, new_sep, str(string))
        # return changed_sep

    # def get_folder_subfolder(categorized_excel, index):
    #     """
    #     it will give the folder name and sub-folder name for the qr code
    #     :param categorized_excel: pd.DataFrame categorized excel
    #     :param index: integer
    #     :return:
    #     """
    #     folder_dat = []
    #     for col_name in folder_col_heads:
    #         folder_info = categorized_excel[col_name][index]
    #         folder_dat.append(folder_info)
    #     return folder_dat

    # def get_data_for_file_number(file_number, categorized_excel):
    #     grouped_data = categorized_excel[categorised_excel['file_number']==file_number]
    #     return grouped_data
    #
    # id_dat = get_id_data(master_list,'38_10', id_cols)

# todo send only 1 line of categorized excel to this function and 1 line of master list
    # todo make required folders in root dir
    def make_qr_code(self, category_row):
        """
        create qr code from input data with proper formating. returns the data used
        to create the qr code
        :param master_list_row:
        :param category_row:
        :param qr_destination_path:
        :return: qr_code_dat
        """
        folders = list(
            category_row[self.hf.function_params('folder_col_heads')])
        file_number=category_row[self.hf.function_params(
            'file_number')].get('file_number')
        # id data = mr number, patient_name, date_of_birth
        id_dat=self.get_id_data(file_number)
        file_number_str=file_number.replace("_", "/")
        folder=[folder for folder in folders if folder is not None]
        id_data=[file_number_str, str(id_dat[0])] + folder
        qr_code_dat='_'.join(id_data)
        # qr_code_dat = str(file_number_str) + '_' + '_'.join(str(id_text) for id_text in (id_dat[0:1] + folder))
        qr_dat_lst=[file_number_str] + id_dat + folder
        qr_dat='_'.join(str(id_text) for id_text in qr_dat_lst)
        qr_img=pyqrcode.create(qr_code_dat)
        qr_dir=self.hf.create_folder_for_data_type(  # type: ignore
            source_folder = 'tmp', data_type = 'qr_code')
        qr_img.png(os.path.join(qr_dir, 'qr_img.png'),   # type: ignore
            scale = 4)

        return qr_dat

    # @staticmethod
    # def format_word_doc(id_text, doc):
        # doc_text = doc.add_paragraph()
        # doc_text = doc_text.add_run(id_text)
        # doc_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # doc_text.bold = True
        # doc_text.font.size = Pt(12)
        # doc_text.font.name = 'Arial Black'
        # return doc

    # def split_id_data_to_string(id_data):
    #     mr_number = id_data['mr_number'][0]
    #     patient_name = id_data['patient_name'][0]
    #     dob = id_data['dob'][0]
    #     return mr_number, patient_name, dob


# tmp_qr_code_folder = 'D:/Shweta/data_digitization/sample_output/2022_03_21/qr_codes'
# tmp_coded_data = 'D:/Shweta/data_digitization/sample_output/2022_03_21/coded_data'
#
# add_qr_code_in_word_document(tmp_qr_code_folder, master_list, categorised_excel, tmp_coded_data)
if __name__ == '__main__':
    qr=QrCode('D:/Shweta/data_digitization',
                'patient_master_list_aj_jj.xlsx',
                '2010_file_categorization_excel.xlsx')
    # master_list = qr.function_params('master_list')
    category_excel=qr.hf.function_params('categorized_excel')
    qr.add_qr_code_in_word_document(category_excel.iloc[0])  # type: ignore
    print('qr code created')
