"""
Creating an xlsx file containing the names of wav files
"""
import xlrd
import os


def file_exists(this_file: str) -> bool:
    """Check whether this_file exists"""
    if not os.path.isfile(this_file):
        print("{} does not exist.".format(this_file))
        print("Please make sure you typed the name of the xlsx file correctly and that you are in the correct"
              " directory.\n")
        return False
    else:
        return True


def no_forbidden_characters(this_new_item: str) -> bool:
    """Check whether this_new_item (file or folder) contains no forbidden characters"""
    forbidden_characters = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    folder_name_wrong = any([char in forbidden_characters for char in this_new_item])
    if folder_name_wrong:
        print("You cannot use the following characters in your folder name or excel file entries:\n")
        print(', '.join(forbidden_characters) + '\n')
        return False
    else:
        return True


def folder_already_exists(this_new_folder: str)-> bool:
    """Check whether a folder called this_folder already exists"""
    this_new_folder_already_exists = os.path.isdir(this_new_folder)
    if this_new_folder_already_exists:
        print("The folder '{}' already exists. Please choose another name or delete the existing folder.\n".format(
            this_new_folder))
        return True
    else:
        return False


def generate_wav_files(excel_file_name: str, new_folder_wav_files: str) -> None:
    """
    Look at the entries in excel_file_name and generate the wav files in new_folder_wav_files
    """

    #  Get a column name and use its ASCII value to (later) determine the column number
    col_name = input("Enter the column containing the Basaa words: ")
    asc_val = ord(col_name)

    col_name_is_not_an_alph = asc_val < 65 or 90 < asc_val < 97 or asc_val > 122
    if col_name_is_not_an_alph:
        print("Make sure the column name is in A...Z or a...z")

    if (file_exists(excel_file_name) and
            no_forbidden_characters(new_folder_wav_files) and
            not folder_already_exists(new_folder_wav_files)) and not col_name_is_not_an_alph:
        # Open workbook
        workbook = xlrd.open_workbook(excel_file_name)

        # Open first sheet. We can change that later.
        worksheet = workbook.sheet_by_index(0)

        total_num_rows = worksheet.nrows

        # Create the folder and move into it
        os.mkdir(new_folder_wav_files)
        os.chdir(new_folder_wav_files)

        replacement_character = '_'  # Can pick another one

        if 65 <= asc_val <= 90:
            col_num = asc_val - 65
        else:
            col_num = asc_val - 97
        
        
        for current_row in range(1, total_num_rows):  # Start from second row
            basaa_word = str(worksheet.cell(current_row, col_num).value).replace(' ', replacement_character)

            word_to_use = basaa_word

            if not no_forbidden_characters(word_to_use):  # Stop as soon as we find an invalid entry
                quit()

            # We can increase that dictionary if need be.
            char_to_replace = {'ɓ': 'B', 'ɛ': 'E', 'ŋ': 'NG', 'ɔ': 'O'}

            for key in char_to_replace:
                word_to_use = word_to_use.replace(key, char_to_replace[key])

            if word_to_use != '':
                basaa_word_file_name = word_to_use + "BW.wav"
                french_word_file_name = word_to_use + "FW.wav"
                basaa_example_file_name = word_to_use + "BE.wav"
                french_example_file_name = word_to_use + "FE.wav"

                open(basaa_word_file_name, 'w')
                open(french_word_file_name, 'w')
                open(basaa_example_file_name, 'w')
                open(french_example_file_name, 'w')

        print("Success!\n")


if __name__ == '__main__':
    excel_file_name = input("Enter the name of the xlsx file: ")
    new_folder_wav_files = input("Enter the name of the new folder for the wav files: ")
    generate_wav_files(excel_file_name + ".xlsx", new_folder_wav_files)
