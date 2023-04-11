# Combine different XLS files into one

The purpuse of this script was to deduplicate all names that are already in the
other Excel files. As well as to bring more standards in the way it looks like, 
e.g. make phone numbers looks the same, make names all in caps etc. 

### How to use it
To use it follow this instruction:
* Create folder called "CF"
* Prepare you files, there should be following columns: *NAME, SURNAME, CONTACT #, EMAIL, SOURCE*
* Make sure that you have all your files saved in the format *.xlsx
* Place all files you want to combine in the CF folder
* python -m venv venv
* Activate environment
* pip install -r requirements.txt

All needed requirements to install you will find here in [requirements](/requirements.txt)

