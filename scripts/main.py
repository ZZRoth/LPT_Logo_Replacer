import file_handling
import os
import zipfile


def main(*args):
    article_folder_path = file_handling.get_directory_filedialog("Ordner auswählen mit Artikeldatenblättern")

    article_file_names = file_handling.get_all_docx_files_in_folder(article_folder_path)
    file_path_empty = file_handling.get_file_path_empty_word_document()

    for article_file_name in article_file_names:
        file_path = os.path.join(article_folder_path, article_file_name)
        file_handling.copy_headers_and_footers(file_path_empty, file_path)


main()
