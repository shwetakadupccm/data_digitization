import ftplib
from genericpath import isdir, isfile
from logging import root
import os
from threading import local
import pandas as pd

REF_DIR_LIST = ['master_list', 'final_data',
                'status_sheets', 'categorization_data', 'scanned_files']
FTP_ROOT_FOLDER = 'CLINICAL DATA BASE/Digitized Files/file_categorization'
COMMON_PATH = 'final_data'


# ROOT_FOLDER = os.path.join('D', 'repos', 'data_digitization')


def open_connection():
    ftp = ftplib.FTP('192.168.1.5')
    ftp.login('Devaki', 'Secure@2023')
    ftp.cwd(FTP_ROOT_FOLDER)
    return ftp


def get_ref_data(ref_dir):
    ftp = open_connection()
    try:
        ftp.cwd(ref_dir)
    except ftplib.error_perm:
        ftp.mkd(ref_dir)
    files = ftp.nlst()
    log_list = None
    for file_name in files:
        if file_name.endswith('.xlsx'):
            if os.path.exists(os.path.join(ref_dir, file_name)):
                print('Previous version of ', file_name, ' deleted')
                os.remove(os.path.join(ref_dir, file_name))
            try:
                ftp.retrbinary(
                    "RETR " + file_name, open(os.path.join(ref_dir, file_name), 'wb').write)
                log = file_name + ' retrieved'
            except ftplib.error_perm:
                log = ' error ' + file_name
            if log_list is None:
                log_list = []
            log_list.append(log)
    ftp.close()
    return(log_list)


def get_all_ref_data():
    create_ref_dir()
    for ref_dir in REF_DIR_LIST[:-1]:
        log_list = get_ref_data(ref_dir)
        print(ref_dir, log_list)


def create_ref_dir():
    [os.mkdir(ref) for ref in REF_DIR_LIST if not isdir(ref)]
    ref_dirs = dir()
    print(ref_dirs)


def get_scanned_file(file_number, ref_dir='scanned_files'):
    ftp = open_connection()
    ftp.cwd(ref_dir)
    file_name = file_number + '.pdf'
    try:
        ftp.retrbinary(
            "RETR " + file_name, open(os.path.join(ref_dir, file_name), 'wb').write)
        log = file_name + ' retrieved'
    except ftplib.error_perm:
        log = ' error ' + file_name
    ftp.close()
    return(log)


def placeFiles(ftp, ftp_path):
    subdir = None
    # while len(os.listdir(path)) != len(ftp.nlst()):
    for name in os.listdir(os.path.join(COMMON_PATH, ftp_path)):
        if os.path.isfile(name):
            print("STOR", name, ftp_path)
            ftp.storbinary('STOR ' + name, open(ftp_path, 'rb'))
            print('Uploaded '+name + ' to ' + ftp_path)
        elif os.path.isdir(name):
            if subdir is None:
                subdir = []
            subdir.append(name)
            print("MKD", name)
            try:
                ftp.mkd(name)
                # ignore("directory already exists")
            except ftplib.error_perm as e:
                if not e.args[0].startswith('550'):
                    raise
    return subdir


def recurse_dir(localpath):
    print(localpath)
    ftp = open_connection()
    paths = localpath.split(COMMON_PATH)
    ftp_path = paths[:1][0]
    # ftp_path = path.replace(ROOT_FOLDER, '')
    print(ftp_path)
    ftp.cwd(os.path.join(COMMON_PATH, ftp_path))
    sub = placeFiles(ftp, ftp_path)
    ftp.close()
    return sub


def copy_dir_struct_files(ref_dir):
    subdir_1 = recurse_dir(ref_dir)
    if subdir_1 is not None:
        subdir_2 = None
        for sub in subdir_1:
            subdir_2 = recurse_dir(os.path.abspath(os.path.join(ref_dir, sub)))
            # ftp.close()
            if subdir_2 is not None:
                for sub2 in subdir_2:
                    subdir_3 = recurse_dir(os.path.join(ref_dir, sub2))
