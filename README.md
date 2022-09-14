# BlackboardQuizExport

Used to generate pdf file exports from blackboard quiz results.  This can be helpful for marking purposes.  Particularly for essay style questions where you may want to markup the answers with marking ticks/comments an feedback inline.  Could be possible to upload the PDFs in to something like turnitin to mark them there with your quickmarks etc.

This python script opens a file diaglog to select a spreadsheet file.

The spreadsheet file is an export from a blackboard quiz or test.  That file will contain student answers to the test, works for essay/short answers and MCQs.  The script parses the file and generates a PDF file of the student answers numbered by row of the spreadsheet to facilitate anonymous marking on paper/pdf.  The actual questions are output also but images etc are not exported from blackboard so you just get the question text.

pip install -r requirements.txt before running.

You will also need wkhtmltopdf installed on your system in the path, or standard windows location.  Available from https://wkhtmltopdf.org/

Don't forget that you need to export the files as a CSV format from Blackboard, select comma and not tabs for format.

