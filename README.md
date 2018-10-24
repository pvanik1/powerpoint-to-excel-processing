# Powerpoint to Excel processing
Python tool to process my custom ppt template and to output a specified excel table. Uses textract and xlsx writer python modules.

Includes executable to run the program located in "\dist\convertPptToWebExcel". To run it from the .py file, to make modifications or to rebuild, please follow the instructions below.

# Release notes:
23/10/2018
Added fields to be parsed and corrected parsing of fields.

3/10/2018
Added executable version using pyinstaller 3.4. Build instructions included.

3/8/2018
Major improvements to robustness, stability and integrity of results. Added follow-up parsing functions. Tested on 60+ files to deal with individual cases generating errors.

14/6/2018
First working version.

# How to run the main .py file:
Instructions based on https://github.com/deanmalmgren/textract/issues/194. Dependencies can be found in the 'installation and dependencies' folder.

-1. Install Python 3.7+
0. download and install microsoft visual c++ build tools (4GB - do steps 1.-3. in the meantime. You will be asked to reboot.)
1. download swigwin, unzip it and add it to system Path variable
	http://www.swig.org/download.html
2. download ebooklib v. 0.15 because textract module is incompatible with newer versions
	https://github.com/aerkalov/ebooklib/releases
3. Delete line 44 in README.md in the ebooklib unzipped folder
4. Download the .whl file of Lxml according to your Python and Windows (32/64 bit) version
	https://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
5. In the directory wher eyou put the .whl file, in Powershell run "pip install [whl filename]"
6. Go to the ebooklib directory in Powershell and type "pip install ." without the quotes
7. Run pip install textract

# Building the executable ---
8. pip install pyinstaller
9. this *should* be fixed in pyinstaller >3.4, but just in case:
   pip install -U pywin32-ctypes
10. pip install pyinstaller
11. run pyinstaller --hidden-import textract.parsers.pptx_parser convertPptToWebExcel.py
    if asked to overwrite/delete during installation, type 'y' and press Enter
12. Find the .exe in the 'dist' directory that was created.
13. When copying the .exe, make sure to copy the whole 'dist' folder