import xlrd
import xlsxwriter
import re
from pathlib import Path

FILE_DIR = Path(__file__).parent

EXCEL_AUTHORS = (FILE_DIR/"../data/UB_cs_authors.xlsx").resolve()
EXCEL_PAPERS = (FILE_DIR/"../data/UB_cs_papers_scopus.xlsx").resolve()

EXCEL_OUTPUT = (FILE_DIR/"../data/UB_cs_papers_cleaned.xlsx").resolve()

PAPER_TYPES = {"Article", "Conference Paper", "Article in Press", "Review", "Book Chapter"}

author_database = [] # list of all relevant authors
paper_set = set() # set for skipping duplicate papers

def init_database():
    global author_database
    wb = xlrd.open_workbook(EXCEL_AUTHORS, on_demand = True)

    # read and insert all authors from excel file to dict list
    for sheet in wb.sheets():
        for row in range(1, sheet.nrows):
            name = sheet.cell_value(row, 0).title()
            lastname = sheet.cell_value(row, 1).title()
            middlename = sheet.cell_value(row, 2) if (sheet.cell_value(row, 2) != "N/A") else ""

            author = dict(name = name, lastname = lastname, middlename = middlename)
            author_database.append(author)

    # sort database in alphabetical order by lastnames then names and middlenames
    author_database = sorted(author_database, key = lambda author: "{0} {1} {2}".format(author['lastname'],
                                                                                        author['name'],
                                                                                        author['middlename']))
    wb.release_resources()
    del wb

def search_name(name, format):
    left = 0
    right = len(author_database) - 1
    while (left <= right):
        middle = left + ((right - left) // 2)
        author = author_database[middle]

        # modify author to match name format
        format_author = format.format(author['name'],       # First name
                                      author['name'][0],    # First letter of name
                                      author['lastname'],   # Last name
                                      author['middlename']) # Middle name
        format_author = format_author.lower()

        if name == format_author:
            return author

        if (name > format_author):
            left = middle + 1
        else:
            right = middle - 1

    return None

def parse_name(name):
    name = name.lower()
    name = name.replace("š", "s").replace("đ", "dj").replace("č", "c").replace("ć", "c").replace("ž", "z")
    name = re.sub("[àáâǎãäåăā]", "a", name);
    name = re.sub("[èéêëėęěĕē]", "e", name);
    name = re.sub("[ìíîï]", "i", name);
    name = re.sub("[òóôǒõöōŏő]", "o", name);
    name = re.sub("[ùúûüũūŭůűǔ]", "u", name);
    name = re.sub("[śŝşš]", "s", name);
    name = re.sub("[ċĉć]", "c", name);
    name = re.sub("[źżž]", "z", name);

    name = name.replace(",", "") # char ',' interferes with binary search and string comparison

    # check format "name middlename lastname" -> "lastname n.m."
    if re.fullmatch(r"[a-z ]+ [a-z]\.[a-z]\.", name):
        format = "{2} {1}.{3}."
        fullname = search_name(name, format)
        if fullname != None:
            return fullname
        name = name[:-2] # lastname n.m. -> lastname n.

    # check format "name lastname" -> "lastname n."
    if re.fullmatch(r"[a-z ]+ [a-z]\.", name):
        format = "{2} {1}."
        fullname = search_name(name, format)
        if fullname != None:
            return fullname

    # check format "name middlename lastname" -> "lastname name m."
    if re.fullmatch(r"[a-z ]+ [a-z]+ [a-z]\.", name):
        format = "{2} {0} {3}."
        fullname = search_name(name, format)
        if fullname != None:
            return fullname
        name = name[:-3] # lastname name m. -> lastname name

    # check format "name lastname" -> "lastname name"
    if re.fullmatch(r"[a-z ]+ [a-z]+", name):
        format = "{2} {0}"
        fullname = search_name(name, format)
        if fullname != None:
            return fullname

    # check format "name middlename lastname" -> "lastname name m"
    if re.fullmatch(r"[a-z ]+ [a-z]+ [a-z]", name):
        format = "{2} {0} {3}"
        fullname = search_name(name, format)
        if fullname != None:
            return fullname
        name = name[:-2] # lastname name m -> lastname name

    # check format "name lastname" -> "lastname name"
    if re.fullmatch(r"[a-z ]+ [a-z]+", name):
        format = "{2} {0}"
        fullname = search_name(name, format)
        if fullname != None:
            return fullname

    return None

def parse_authors(authors):
    fullnames = []
    names = authors.split(" and ")

    for name in names:
        author = parse_name(name.strip())
        if author != None:
            fullnames.append("{0} {1}".format(author['name'], author['lastname']))
 
    return ", ".join(fullnames)

def parse_line(line):
    global paper_set

    # Check paper type
    ptype = line[0]
    if ptype not in PAPER_TYPES:
        return None

    year = line[2]
    title = line[3]
    docname = line[5]

    # Check duplicate papers - leave duplicates that were published separately
    paper = year + " " + title.lower() + " " + docname.lower()
    if paper in paper_set:
        return None
    else:
        paper_set.add(paper)

    # Parse main author
    main_author = parse_authors(re.sub("N/A", "", line[1]))

    # Parse authors
    authors = parse_authors(line[4])
    if (authors.find(main_author) == -1):
        authors = (main_author + ", " + authors) if (authors != "") else (main_author)

    return [ptype, year, title, authors, docname]

def clean():
    print("Initializing author database: " + str(EXCEL_AUTHORS))
    init_database()

    # Input
    wb = xlrd.open_workbook(EXCEL_PAPERS, on_demand = True)
    sheet = wb.sheet_by_index(0)
    print("Reading papers data file: " + str(EXCEL_PAPERS))

    # Output
    owb = xlsxwriter.Workbook(EXCEL_OUTPUT)
    outsheet = owb.add_worksheet(sheet.name)
    outrow = 1

    # Column header titles
    outsheet.write(0, 0, sheet.cell_value(0, 5)) # Type
    outsheet.write(0, 1, sheet.cell_value(0, 2)) # Year
    outsheet.write(0, 2, sheet.cell_value(0, 1)) # Title
    outsheet.write(0, 3, sheet.cell_value(0, 3)) # Authors
    outsheet.write(0, 4, sheet.cell_value(0, 9)) # Document Name

    print("Cleaning...")
    for row in range(1, sheet.nrows):
        line = [sheet.cell_value(row, 5), # Type
                sheet.cell_value(row, 0), # Main Author
                sheet.cell_value(row, 2), # Year
                sheet.cell_value(row, 1), # Title
                sheet.cell_value(row, 3), # Authors
                sheet.cell_value(row, 9)] # Document Name

        line = parse_line(line)

        if line == None:
            continue
        
        outsheet.write_row(outrow, 0, line)
        outrow += 1
    print("Cleaned papers data written to: " + str(EXCEL_OUTPUT))

    outsheet.set_column(0, 0, 22) # Type column width 22 chars
    outsheet.set_column(1, 1, 8)  # Year column width 8 chars
    outsheet.set_column(2, 2, 90) # Title column width 90 chars
    outsheet.set_column(3, 3, 50) # Authors column width 50 chars
    outsheet.set_column(4, 4, 80) # Document name column width 80 chars

    wb.release_resources()
    del wb
    owb.close()
    del owb

clean()