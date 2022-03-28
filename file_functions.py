'''
functions that are required to move/rename files create folders etc
define paramters
'''
import os
import re
import pandas as pd
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as WD_ALIGN_PARAGRAPH
from docx.shared import Pt

class HelperFunctions:
    '''
    functions that are required to move/rename files create folders etc
    define paramters
    '''

    def __init__(self, root, master_list_name, categorized_excel):
        self.root = root
        self.master_list_name = master_list_name
        self.categorized_excel = categorized_excel
     # ignore type unexpected-keyword-arg)

    def function_params(self, param_name):
        '''
        define paths and names and data dfs used for all data_digitization functions
        '''
        param_dict = dict(master_list=pd.read_excel(os.path.join(self.root, 'reference_docs', self.master_list_name)),
                          categorized_excel=pd.read_excel(os.path.join(
                              self.root, 'reference_docs', self.categorized_excel)),
                          id_cols=['mr_number',
                                   'patient_name', 'date_of_birth'],
                          file_number=['file_number'],
                          qr_data_cols=['File Number', 'MR Number',
                                        'Patient Name', 'Date of Birth'],
                          folder_col_heads=['report_name', 'subfolder_name'],
                          report_types_list=['Patient Information',
                                             'Clinical Examination',
                                             'Radiology',
                                             'Metastatic Examination',
                                             'Biopsy Pathology',
                                             'Neo-Adjuvant Chemotherapy',
                                             'Surgical Procedures',
                                             'Patient Images',
                                             'Surgery Media',
                                             'Surgery Pathology',
                                             'Chemotherapy',
                                             'Radiotherapy',
                                             'Follow-up Notes',
                                             'Genetics',
                                             'Miscellaneous',
                                             'Patient File Data',
                                             'PROMS'])
        dat = param_dict.get(param_name)
        return dat

    # def create_folder_for_data_type(self, data_type, source_folder):
    #     '''
    #     create folders and subfolders for data types in source folders for eg. tmp
    #     '''
    #     data_type_dir = False
    #     if not os.path.isdir(os.path.join(self.root, source_folder)):
    #         os.mkdir(os.path.join(self.root, source_folder))
    #     dat = []
    #     for folder in data_type:
    #         dat.append(folder)
    #         folders = "/".join(dat)
    #         data_type_dir = os.path.join(self.root, source_folder, folders)
    #         if not os.path.isdir(data_type_dir):
    #             os.mkdir(data_type_dir)
    #     return data_type_dir

    # data type = list?
    def create_folder_for_data_type(self, data_type, source_folder):
        source_path = os.path.join(self.root, source_folder)
        if not os.path.isdir(source_path):
            os.mkdir(source_path)
        data_type_dir = os.path.join(source_path, data_type)
        if not os.path.isdir(data_type_dir):
            os.mkdir(data_type_dir)
        return data_type_dir

    @staticmethod
    def change_sep(string, old_sep, new_sep):
        """
        change the separator between the string or within
        :param string: string
        :param old_sep: string separators (' ', '_', '/')
        :param new_sep: string separators (' ', '_', '/')
        :return: string with changed separators
        """
        changed_sep = re.sub(old_sep, new_sep, str(string))
        return changed_sep

    # @staticmethod
    # def format_word_doc(id_text, doc):
    #
    #     doc_text = doc.add_paragraph()
    #     doc_text = doc_text.add_run(id_text)
    #     doc_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    #     doc_text.bold = True
    #     doc_text.font.size = Pt(12)
    #     doc_text.font.name = 'Arial Black'
    #     return doc

    @staticmethod
    def format_word_doc(id_text, doc):
        """
        format text strings in a word document
        :param id_text: id_value(eg = 'file_number', 'mr_number', 'patient_name', 'dob')
        :param doc: docx.Document
        :return: formatted document
        """
        doc_text = doc.add_paragraph()
        doc_text = doc_text.add_run(id_text)
        doc_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # type: ignore
        doc_text.bold = True
        doc_text.font.size = Pt(12)
        doc_text.font.name = 'Arial Black'
        return doc
