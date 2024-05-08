import sys
import os
from docx import Document
from tkinter import filedialog
import zipfile
import win32com.client


def get_scripts_directory_path():
    return getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))


def get_project_directory_path():
    return getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(get_scripts_directory_path())))


def get_assets_directory_path():
    return os.path.join(get_project_directory_path(), 'assets')


def get_file_path_empty_word_document():
    return os.path.join(get_assets_directory_path(), "XXX-000-X_LPT_LKS_Rohdatei_IBE.docx")


def get_word_document(file_path):
    return Document(file_path)


def get_directory_filedialog(title):
    return os.path.abspath(filedialog.askdirectory(title=title))


def save_word_file(document, file_path):
    document.save(file_path)


def get_all_docx_files_in_folder(folder_path):
    docx_files = []
    for file in os.listdir(folder_path):
        if file.endswith(".docx"):
            docx_files.append(file)
    return docx_files


def copy_headers_and_footers(old_docx_path, new_docx_path):
    backup_path = new_docx_path + ".bak"
    os.rename(new_docx_path, backup_path)

    # Open the old and backup documents
    with zipfile.ZipFile(old_docx_path, 'r') as old_zip, zipfile.ZipFile(backup_path, 'r') as backup_zip:
        # Create a new archive for the new document
        with zipfile.ZipFile(new_docx_path, 'w') as new_zip:
            print(backup_zip.namelist())
            # Copy files from the backup, except for headers and footers
            for name in backup_zip.namelist():
                if not name.startswith('word/media/image2.jpeg'):
                    content = backup_zip.read(name)
                    new_zip.writestr(name, content)

            # Copy headers and footers from the old document
            for name in old_zip.namelist():
                if name.startswith('word/media/image1.jpeg'):
                    content = old_zip.read(name)
                    new_zip.writestr('word/media/image2.jpeg', content)

    # Remove the backup file
    os.remove(backup_path)


def copy_content_of_docx_file_into_empty_docx_file(empty_document, full_document):

    for element in full_document.element.body:
        empty_document.element.body.append(element)

    return empty_document
