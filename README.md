Word Document Merger with Comments
This Python tool allows you to merge pairs of Word documents (*.docx) from a specified directory. Each pair consists of two files with names ending in "-a.docx" and "-J.docx", and the tool combines them into a single merged document, preserving any comments within the documents.

Prerequisites
Python 3.x installed on your system
Required Python libraries (tkinter, docx)

1) Select DirectoryClick on the "Select Directory" button in the GUI to choose the directory containing the pairs of Word documents.
2) Combine FilesClick on the "Combine Files" button to merge the identified pairs of Word documents. The merged documents will be saved in the selected output directory.

Notes
The tool assumes that each pair of Word documents consists of one file ending in "-a.docx" and another ending in "-J.docx".
Comments from both documents in each pair will be included in the merged document.
Example Directory Structure
css
Copy code
my_documents/
├── document1-a.docx
├── document1-J.docx
├── document2-a.docx
├── document2-J.docx
├── ...

Author
LiyanageCMadusanka
GitHub: Liyanage99
