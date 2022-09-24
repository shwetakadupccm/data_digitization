import shutil
import os
import re
from PyPDF2 import PdfFileReader, PdfFileWriter
from create_qr_code import QrCode
from file_functions import HelperFunctions


class CategorizeFile():

    def __init__(self, root, master_list_name, categorized_excel):
        self.root = root
        self.master_list_name = master_list_name
        self.categorized_excel = categorized_excel
        self.hf = HelperFunctions(root=self.root, master_list_name=self.master_list_name,
                                  categorized_excel=self.categorized_excel)
        self.qr_code = QrCode(root=self.root, master_list_name=self.master_list_name,
                              categorized_excel=self.categorized_excel)

    def split_pdf_by_pages(self, file_number):
        tmp_folder = self.qr_code.create_tmp_folder_for_data_type(
            data_type='splitted_files')
        splitted_file_path = os.path.join(tmp_folder, str(file_number))
        if not os.path.isdir(splitted_file_path):
            os.mkdir(splitted_file_path)
        scanned_file_name = str(file_number) + '.pdf'
        scanned_file = PdfFileReader(os.path.join(
            'scanned_files', scanned_file_name))
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

    # @classmethod
    def classify_file_images_by_report_types(self, splitted_file_path_file_no, report_page_nums, file_number,
                                             report_type):
        splitted_scanned_files = os.listdir(splitted_file_path_file_no)
        img_no_lst = self.get_image_no(file_number, splitted_scanned_files)
        report_page_no_splitted = self.hf.split_report_page_no(
            report_page_nums)
        file_no_dir = os.path.join('coded_data', str(file_number))
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

    # @classmethod
    def copy_rename_page(self, file_number, page_no_list, splitted_file_path, folder_dir):
        report_no_lst = self.get_image_no(
            file_number, os.listdir(splitted_file_path))
        for idx, page_no in enumerate(page_no_list):
            if page_no in report_no_lst:
                report_page_name = self.hf.make_pdf_name_using_page_no(
                    page_no, file_number)
                source_path = os.path.join(
                    splitted_file_path, report_page_name)
                new_file_name = self.hf.make_pdf_name_using_page_no(
                    str(idx + 1), file_number)
                dest_path = os.path.join(folder_dir, new_file_name)
                shutil.copy(source_path, dest_path)

    @staticmethod
    def rename_images(pdf_doc_path, dir_path, file_no, report_type):
        report_dir = os.path.join(dir_path, str(report_type))
        img_list = os.listdir(report_dir)
        for index, img in enumerate(img_list):
            old_file_path = os.path.join(report_dir, img)
            img_no = index + 1
            new_name = str(file_no) + '_' + str(img_no) + '.pdf'
            file_dir = os.path.join('coded_files', str(file_no))
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

    def make_folder_for_report_type(self, file_number, folder_subfolder_lst):
        file_number_folder = os.path.join(
            'coded_files', str(file_number))
        if not os.path.isdir(file_number_folder):
            os.mkdir(file_number_folder)
        parent_folder = os.path.join(
            file_number_folder, folder_subfolder_lst[0])
        if not os.path.isdir(parent_folder):
            os.mkdir(parent_folder)
        subdir = os.path.join('coded_files', str(file_number),
                              '/'.join(str(folder) for folder in folder_subfolder_lst if folder != 'nan'))
        if not os.path.isdir(subdir):
            os.mkdir(subdir)
        return subdir

    # def categorize_scanned_file(self, file_number, doc_data):
    #
    #     splitted_file_path = self.split_pdf_by_pages(file_number) # splitting the scanned pdf using file number
    #     folder_dir = self.make_folder_for_report_type(file_number, doc_data[4:]) ## it creates a directories for report type and return dir path
    #     report_page_nums = category_row['page_numbers'] # page numbers for report type
    #     page_no_lst = self.split_report_page_no(report_page_nums) # list of page numbers
    #     shutil.move(coded_pdf_path, folder_dir) # moving the coded pdf to report type folder
    #     self.copy_rename_page(file_number, page_no_lst, splitted_file_path, folder_dir) # copy and renaming the report pages in report type folder
    #

    def categorize_file_by_report_types(self):
        qr_code = QrCode(root=self.root, master_list_name=self.master_list_name,
                         categorized_excel=self.categorized_excel)
        categorized_data = self.hf.function_params(
            'categorized_excel')  # reading categorized excel
        for i in range(len(categorized_data)):  # type: ignore
            file_number = category_row[self.hf.function_params('file_number')].get(
                'file_number')  # file number from that row
            # creating qr code, adding # creating qr code, adding  it into doc, converting  it into pdf and returns doc_data list and pdf path
            doc_data, coded_pdf_path = qr_code.add_qr_code_in_word_document(
                category_row)
            # doc_dat = [file_number, mr_number, patient_name, date_of_birth, folder, subfolder]
            # splitting the scanned pdf using file number
            splitted_file_path = self.split_pdf_by_pages(file_number)
            # it creates a directories for report type and return dir path
            folder_dir = self.make_folder_for_report_type(
                file_number, doc_data[4:])
            # page numbers for report type
            report_page_nums = category_row['page_numbers']
            page_no_lst = self.hf.split_report_page_no(
                report_page_nums)  # list of page numbers
            # moving the coded pdf to report type folder
            shutil.move(coded_pdf_path, folder_dir)
            # copy and renaming the report pages in report type folder
            self.copy_rename_page(file_number, page_no_lst,
                                  splitted_file_path, folder_dir)
