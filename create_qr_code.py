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

class QrCode():

    def __init__ (self, root, master_list_name, categorized_excel):
        self.root = root
        self.master_list_name = master_list_name
        self.categorized_excel = categorized_excel

    def function_params(self, param_name):
        param_dict = dict(master_list= pd.read_excel(os.path.join(self.root, 'reference_docs', self.master_list_name)),
                          categorized_excel = pd.read_excel(os.path.join(self.root, 'reference_docs', self.categorized_excel)),
                          file_number = ['file_number'],
                          id_cols=['mr_number', 'patient_name', 'date_of_birth'],
                          folder_col_heads=['report_name', 'subfolder_name'],
                          report_types_dict=['Patient Information', 'Clinical Examination',
                                            'Radiology', 'Metastatic Examination',
                                            'Biopsy Pathology', 'Neo-Adjuvant Chemotherapy',
                                            'Surgical Procedures', 'Patient Images',
                                            'Surgery Media', 'Surgery Pathology',
                                            'Chemotherapy', 'Radiotherapy',
                                            'Follow-up Notes', 'Genetics',
                                            'Miscellaneous', 'Patient File Data',
                                            'PROMS'])
        return param_dict.get(param_name)


    def add_qr_code_in_word_document(self, id_data):
        qr_code_dat = self.make_qr_code() #add input params
        doc_dat = qr_code_dat.split("_")
        patient_name, date_of_birth = id_data[2:]
        # master_list = self.function_param.get('master_list')
        # patient_name = list(master_list[master_list['file_number'] == doc_dat[0] & master_list['mr_number'] == doc_dat[1]]['patient_name'])[0]
        # for row, file_number in enumerate(self.categorised_excel['file_number']):
        #     print(row, file_number)
        doc = Document()
        # id_data = get_id_data(self.master_list, file_number, self.function_params('id_cols'))
        # id_name = '_'.join(str(id) for id in id_data)
        #     doc_dat = str.split(qr_code_dat, '_')
        #     mr_number, patient_name, dob = split_id_data_to_string(id_data)
        #     report_type = categorised_excel['report_name'][row]
        #     report_type_str = change_sep(report_type, ' ', '_')
        #     report_type_no = categorised_excel['report_type_number'][row]
        #     subfolder = categorised_excel['subfolder_name'][row]
        #     subfolder_str = change_sep(subfolder, ' ', '_')
        #     qr_code_name = str(file_number) + '_' + str(mr_number) + '_' + str(report_type_str) + '_' + str(
        #         subfolder_str) + '.png'
        #     qr_code_path = os.path.join(tmp_qr_code_folder, qr_code_name)
        doc.add_picture(os.path.join(self.root, 'tmp\qr_code\qr_code.png'))
        qr_alignment = doc.paragraphs[-1]
        qr_alignment.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        doc = doc.add_paragraph()
        folders = doc_dat[2:],
        id_texts = [', '.join(doc_dat[0:2]), patient_name, date_of_birth]
        # should work?
        doc = [self.format_word_doc(folder, doc) for folder in folders]
        blank_para = doc.add_paragraph()
        run = blank_para.add_run()
        run.add_break()
        doc = [self.format_word_doc(id_text, doc) for id_text in id_texts]
            # format_word_doc(id_text, doc)
            # report_type_name = text.add_run(str(report_type_no) + '. ' + str(report_type))
            # format_word_doc(report_type_name)
            # blank_para = doc.add_paragraph()
            # run = blank_para.add_run()
            # run.add_break()
            # id_text = 'File Number: ' + str(file_number)
            # format_word_doc(id)
            # id = doc.add_paragraph().add_run('MR Number: ' + str(mr_number))
            # format_word_doc(id)
            # id = doc.add_paragraph().add_run('Patient Name: ' + str(patient_name))
            # format_word_doc(id)
            # id = doc.add_paragraph().add_run(dob)
            # format_word_doc(id)
        doc_name = qr_code_dat.replace("/| ", "_") + '.docx'
        doc_path = os.path.join(os.path.join(self.root, 'tmp/coded_data', doc_name))
        doc.save(doc_path)
        pdf_path = doc_path.replace('.docx', '.pdf')
        # pdf_path = os.path.join(doc_path, pdf_name)
        convert(doc_path, pdf_path)

    def get_id_data(self, file_number):
        """
        get id values from input id names and single row of master list
        :param master_list: pd.DataFrame
        :param file_number: file_number separated by '_'
        :param id_cols: col-names which stores the patients identifing info(file_number, mr_number, name, dob)
        :return: list of id data for single row of master list
        """
        master_list = self.function_params('master_list')
        id_cols = self.function_params('id_cols')
        id_data = master_list[master_list['file_number'] == file_number][id_cols]
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
    def make_qr_code(self, file_number, category_row):
        """
        create qr code from input data with proper formating. returns the data used
        to create the qr code
        :param master_list_row:
        :param category_row:
        :param qr_destination_path:
        :return: qr_code_dat
        """
        folders = list(category_row[self.function_params('folder_col_heads')])
        file_number = str(category_row[self.function_params('file_number')])
        id_dat = self.get_id_data(file_number)
        file_number.replace("_", "/")
        folder = [folder for folder in folders if folder is not None]
        id_dat = id_dat + folders
        qr_code_dat = str(file_number) + '_'.join(id_dat + folder)
        qr_img = pyqrcode.create(qr_code_dat)
        qr_img.png(os.path.join(self.root, 'tmp/qr_code/qr_img.png'), scale=4)
        return qr_code_dat

    def format_word_doc(id_text, doc):
        doc_text = doc.add_paragraph().add_run(id_text)
        doc_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc_text.bold = True
        doc_text.font.size = Pt(28)
        doc_text.font.name = 'Arial Black'
        blank_para = doc.add_paragraph()
        blank_para.add_run()
        return doc

    # def split_id_data_to_string(id_data):
    #     mr_number = id_data['mr_number'][0]
    #     patient_name = id_data['patient_name'][0]
    #     dob = id_data['dob'][0]
    #     return mr_number, patient_name, dob


# tmp_qr_code_folder = 'D:/Shweta/data_digitization/sample_output/2022_03_21/qr_codes'
# tmp_coded_data = 'D:/Shweta/data_digitization/sample_output/2022_03_21/coded_data'
#
# add_qr_code_in_word_document(tmp_qr_code_folder, master_list, categorised_excel, tmp_coded_data)
