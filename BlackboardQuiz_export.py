# Convert xlsx spreadsheet export from Blackboard quiz to numbered html and PDF files in directory to aid printout for marking
# Created by Dr Tim Wilding,  2022
# Copyright (c) 2022, Dr Tim Wilding
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.
import errno
import os
import pathlib
import shutil
import openpyxl.utils.exceptions
from openpyxl import load_workbook
from pathlib import Path
from tkinter import filedialog
import tkinter
import sys
import pdfkit
import atexit


# Main
if __name__ == '__main__':
    atexit.register(input, "Enter any Key to Close/Exit")
    pdf = 1

    if wkh := shutil.which("wkhtmltopdf"):
        print("Found wkhtmltopdf in path at %s, using that version" % wkh)
        wkh_location = wkh
    else:
        wkh_location = 'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
        print("wkhtmltopdf not found in path, trying at this location %s" % wkh_location)

    try:
        config = pdfkit.configuration(wkhtmltopdf=wkh_location)
        print("Found wkhtmltopdf at %s using pdf output." % wkh_location)
    except IOError:
        print("Couldn't find or open wkhtmltopdf.exe, running without save to pdf.  Html files only.")
        print("Expecting wkhtmltopdf at %s, edit script to change location, or install wkhtmltopdf at"
                " that location.\nDownload wkhtmltopdf from https://wkhtmltopdf.org/ or use package manager." %
                wkh_location)
        pdf = 0

    # start a root tk window and hide it now, so that it goes away before we do the filedialog
    root = tkinter.Tk()
    root.withdraw()
    # root.update()
    filename = filedialog.askopenfilename(title="Spreadsheet results file",
                                          filetypes=(("xlsx", "*.xlsx"), ("all", "*.*")))

    print("Reading spreadsheet file %s" % filename)
    try:
        wb = load_workbook(filename=filename)
    except IOError:
        print("Error: Could not open the spreadsheet file, is it open in Excel?  Close it, then try again, "
              "or check file and file permissions if not open.")
        sys.exit(1)
    except openpyxl.utils.exceptions.InvalidFileException:
        print("Cannot open the file, is it a valid xlsx excel file? Check it opens in Excel, close it, then try again.")
        sys.exit(1)

    fullpath = Path(filename)
    filepath = str(fullpath.parent)
    # remove spaces and . from filename when creating output path to avoid issues with directories with spaces or
    # multiple .
    i_filename = "_".join(fullpath.stem.split(" "))
    i_filename = "_".join(i_filename.split("."))
    # make the path
    out_dir = pathlib.PurePath(filepath, i_filename)

    print("directory: %s name: %s" % (filepath, i_filename))

    sheet = wb.worksheets[0]

    row_count = sheet.max_row
    col_count = sheet.max_column

    print("Spreadsheet has %d rows, %d cols" % (row_count, col_count))

    header = sheet[1]
    if col_count < 9 or header[0].value != "Username" or header[1].value != "Last Name" or \
            header[2].value != "First Name" or header[3].value != "Question ID 1":
        print("This spreadsheet doesn't look like a Blackboard Quiz result download.  I'm stopping here as something"
              " isn't quite right.  Please take a look at the spreadsheet file.")
        sys.exit(1)

    try:
        os.makedirs(out_dir)
    except OSError as exc:
        # if the directory already exists then do nothing, as that is ok.  Otherwise, raise error (which also exits).
        if exc.errno != errno.EEXIST:
            raise
        print("Output directory %s already exists.  Existing files will be overwritten" % out_dir)
        if not input("Are you sure? (y/n): ").lower().strip()[:1] == "y":
            print("You entered something, which is not Yes or Y or y or yes, so I'm exiting now.  Goodbye!")
            sys.exit(1)
        pass

    # first three columns are username, last name, first name
    # next columns are Question ID n,  Question n, Answer n, Possible points n, Auto Score n, Manual Score n
    number_of_questions = int((col_count - 3) / 6)

    print("There are %d questions to parse" % number_of_questions)

    questions_store = []
    answer_array = [[]]
    student_array = [[]]
    count = 1
    # iterate through sheet, starting row 2 as row 1 is the headings row
    for row in sheet.iter_rows(2):
        outfile_html = pathlib.PurePath(out_dir, (str(count) + ".html"))
        outfile_pdf = pathlib.PurePath(out_dir, (str(count) + ".pdf"))
        with open(outfile_html, 'w') as f:
            print("<html><body>", file=f)
            print("Student %s" % row[0].value)
            print("<h1>Student number %d</h1>" % count, file=f)
            # loop by row, starting from row 2 as row one is header
            for qlp in range(number_of_questions):
                # print("Question %d" % (qlp+1))
                # here we get the question and replace \n with <br />
                question = str(row[3 + (qlp*6) + 1].value)
                parsed_question = "<br />".join(question.split("\\n"))
                # here we get the answer and replace \n with <br />
                answer = str(row[3 + (qlp*6) + 2].value)
                parsed_answer = "<br />".join(answer.split("\\n"))
                if parsed_question not in questions_store:
                    questions_store.append(parsed_question)
                    # get length at this point, because array starts at pos 0, but we label from one, so first
                    # one will be 1
                    q_num = len(questions_store)
                    print("Found new question %d" % q_num)
                    question_html = pathlib.PurePath(out_dir, ("Question_" + str(q_num) + ".html"))
                    question_pdf = pathlib.PurePath(out_dir, ("Question_" + str(q_num) + ".pdf"))
                    # this is a new question, so append the answer to the end of the outer array
                    answer_array.append([parsed_answer])
                    student_array.append([count])
                    with open(question_html, 'w') as qf:
                        print("<html><body>", file=qf)
                        print("<h1>Question %d</h1>" % q_num, file=qf)
                        print(parsed_question, file=qf)
                        print("</body></html>", file=qf)
                    if pdf:
                        pdfkit.from_file(str(question_html), str(question_pdf), configuration=config)
                else:
                    # find array position, we count from Q number 1 but arrays start at 0
                    q_num = questions_store.index(parsed_question) + 1
                    print("Found question %d" % q_num)
                # print("Question <br /> %s" % parsed_question, file=f)
                    # build up answer array, use q_num - 1 as q-num starts at 1, whereas array indexes start at zero
                    answer_array[q_num-1].append(parsed_answer)
                    student_array[q_num-1].append(count)
                print("<h2>Answer %d for question %d</h2><br /> %s" % (qlp+1, q_num, parsed_answer), file=f)
            print("</body></html>", file=f)

        if pdf:
            pdfkit.from_file(str(outfile_html), str(outfile_pdf), configuration=config)
        # finished this student, add one to count for the next student, then loop back for next row
        count += 1

 
