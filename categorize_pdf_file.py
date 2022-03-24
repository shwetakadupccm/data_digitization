import shutil
import math
import os
import re
import numpy as np
from PyPDF2 import PdfFileReader, PdfFileWriter
from create_qr_code import QrCode

class CategorizeFile():

    def __init__(self, root, master_list_name, categorized_excel, scanned_files_folder_path):
        self.root = root
        self.scanned_files_folder_path = scanned_files_folder_path
        self.master_list_name = master_list_name
        self.categorized_excel = categorized_excel

    def split_pdf_by_pages(self, file_number):
        qr_code = QrCode(root=self.root, master_list_name=self.master_list_name, categorized_excel=self.categorized_excel)
        tmp_folder = qr_code.create_tmp_folder_for_data_type(data_type='splitted_files')
        splitted_file_path = os.path.join(tmp_folder, str(file_number))
        if not os.path.isdir(splitted_file_path):
            os.mkdir(splitted_file_path)
        scanned_file_name = str(file_number) + '.pdf'
        scanned_file = PdfFileReader(os.path.join(self.scanned_files_folder_path, scanned_file_name))
        page_range = scanned_file.getNumPages()
        for i in range(page_range):
            page = scanned_file.getPage(i)
            page_no = i + 1
            splitted_file = str(file_number) + '_' + str(page_no) + '.pdf'
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(page)
            with open(os.path.join(splitted_file_path, splitted_file), 'wb') as out:
                pdf_writer.write(out)
        print("file number: ", file_number + " splitted")
        return splitted_file_path

    @staticmethod
    def get_image_no(file_number, file_images_lst):
        file_images_no_lst = []
        for file_image in file_images_lst:
            file_image_no = re.sub(file_number, '', str(file_image))
            file_image_no = re.sub('.pdf', '', file_image_no)
            file_image_no = re.sub('_', '', file_image_no)
            file_image_no = file_image_no.strip()
            file_images_no_lst.append(file_image_no)
        return file_images_no_lst

    @staticmethod
    def split_report_page_no(report_page_no):
        if isinstance(report_page_no, float):
            page_no_lst = []
            if not math.isnan(report_page_no):
                integer = int(report_page_no)
                page_no_lst.append(str(integer))
            return page_no_lst
        elif isinstance(report_page_no, int):
            page_no_lst = []
            page_no_lst.append(str(report_page_no))
            return page_no_lst
        elif ';' in report_page_no:
            page_no_lst = []
            report_page_no_splitted = report_page_no.split(';')
            for page_no in report_page_no_splitted:
                if '|' in page_no:
                    partitions = page_no.partition('|')
                    start = int(partitions[0])
                    end = int(partitions[2]) + 1
                    page_nos = np.arange(start, end)
                    page_nos_lst = page_nos.tolist()
                    for no in page_nos_lst:
                        page_no_lst.append(str(no))
                else:
                    page_no_lst.append(str(page_no))
            return page_no_lst
        elif '|' in report_page_no:
            page_no_lst = []
            partitions = report_page_no.partition('|')
            start = int(partitions[0])
            end = int(partitions[2]) + 1
            page_nos = np.arange(start, end)
            page_nos_lst = page_nos.tolist()
            for no in page_nos_lst:
                page_no_lst.append(str(no))
            return page_no_lst
        elif type(report_page_no) in (float, int):
            page_no_lst = []
            report_page_no = int(report_page_no)
            page_no_lst.append(str(report_page_no))
            return page_no_lst
        else:
            page_no_lst = []
            page_no_lst.append(str(report_page_no))
            return page_no_lst

    @classmethod
    def classify_file_images_by_report_types(self, splitted_file_path_file_no, report_page_nums, file_number,
                                             report_type, destination_path):
        splitted_scanned_files = os.listdir(splitted_file_path_file_no)
        img_no_lst = self.get_image_no(file_number, splitted_scanned_files)
        report_page_no_splitted = self.split_report_page_no(report_page_nums)
        file_no_dir = os.path.join(destination_path, str(file_number))
        if not os.path.isdir(file_no_dir):
            os.mkdir(file_no_dir)
        report_dir = os.path.join(file_no_dir, report_type)
        if not os.path.isdir(report_dir):
            os.mkdir(report_dir)
        for page_no in report_page_no_splitted:
            if page_no in img_no_lst:
                report_page_name = str(file_number) + \
                                   '_' + str(page_no) + '.pdf'
                source_path = os.path.join(
                    splitted_file_path_file_no, report_page_name)
                dest_path = os.path.join(report_dir, report_page_name)
                shutil.copy(source_path, dest_path)

    @staticmethod
    def make_pdf_name_using_page_no(page_no_lst, file_number):
        for page_no in page_no_lst:
            report_page_name = str(file_number) + '_' + str(page_no) + '.pdf'
        return report_page_name

    @staticmethod
    def get_report_page_rename(page_no_lst, report_page_lst):

        for page_no in page_no_lst:
            report_name =


    @staticmethod
    def rename_images(pdf_doc_path, dir_path, file_no, report_type, destination_path):
        report_dir = os.path.join(dir_path, str(report_type))
        img_list = os.listdir(report_dir)
        for index, img in enumerate(img_list):
            old_file_path = os.path.join(report_dir, img)
            img_no = index + 1
            new_name = str(file_no) + '_' + str(img_no) + '.pdf'
            file_dir = os.path.join(destination_path, str(file_no))
            if not os.path.isdir(file_dir):
                os.mkdir(file_dir)
            new_file_path = os.path.join(file_dir, report_type)
            if not os.path.isdir(new_file_path):
                os.mkdir(new_file_path)
            dest_path = os.path.join(new_file_path, new_name)
            shutil.copy(old_file_path, dest_path)
            coded_file_name = 'code_' + \
                              str(file_no) + '_' + str(report_type) + '.pdf'
            shutil.copy(pdf_doc_path, os.path.join(
                new_file_path, coded_file_name))




    def categorize_file_by_report_types(self):
        qr_code = QrCode(root=self.root, master_list_name=self.master_list_name,
                         categorized_excel=self.categorized_excel)
        categorized_excel = qr_code.function_params('categorized_excel')
        for i in range(len(categorized_excel)):
            # file_number = self.categorized_files_df['file_number'][i]
            # mr_number = self.categorized_files_df['mr_number'][i]
            # patient_name = self.categorized_files_df['patient_name'][i]
            # dob = self.categorized_files_df['date_of_birth'][i]
            category_row = categorized_excel.iloc[i]
            file_number = category_row[qr_code.function_params('file_number')].get('file_number')
            coded_pdf_path = qr_code.add_qr_code_in_word_document(category_row)
            splitted_file_path = self.split_pdf_by_pages(file_number)
            report_page_nums = category_row['page_numbers']
            page_no_lst = self.split_report_page_no(report_page_nums)
            report_no_lst = self.get_image_no(file_number, os.listdir(splitted_file_path))




            # classified_files_path = os.path.join(
            #     self.tmp_folder_path, 'classfied_files')
            # if not os.path.isdir(classified_files_path):
            #     os.mkdir(classified_files_path)
            # self.classify_file_images_by_report_types(splitted_file_dir, str(
            #     report_page_nums), file_number, report_type_str, classified_files_path)
            # renamed_files_path = os.path.join(
            #     classified_files_path, str(file_number))
            # self.rename_images(coded_pdf_path, renamed_files_path, str(
            #     file_number), report_type_str, self.destination_path)
            # print("file: ", file_number,
            #       ' classified by report types and arranged by sequence')
