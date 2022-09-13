# BlackboardQuizExport

This python scrip opens a file diaglog to select a spreadsheet file.

The spreadsheet file is an export from a blackboard quiz/test.  That file will contain student answers to the test.  The script parses the file and generates a PDF file of the student answers numbered by row of the spreadsheet to facilitate anonymous marking on paper/pdf.  The actual questions are output also but images etc are not exported from blackboard so you just get the question text.

pip install -r requirements.txt before running.

You will also need wkhtmltopdf installed on your system in the path, or standard windows location.  Available from https://wkhtmltopdf.org/



