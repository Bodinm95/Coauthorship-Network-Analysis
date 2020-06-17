import xlrd
import xlsxwriter
import networkx as nx
from pathlib import Path

FILE_DIR = Path(__file__).parent

EXCEL_PAPERS = (FILE_DIR/"../data/UB_cs_papers_cleaned.xlsx").resolve()
GRAPH_INPUT = (FILE_DIR/"../results/authors_graph.gexf").resolve()
ETF_INPUT = (FILE_DIR/"../results/authors_graph_etf.gexf").resolve()
MATF_INPUT = (FILE_DIR/"../results/authors_graph_matf.gexf").resolve()
FON_INPUT = (FILE_DIR/"../results/authors_graph_fon.gexf").resolve()

EXCEL_AUTHORS = (FILE_DIR/"../results/analysis_authors.xlsx").resolve()
EXCEL_MODULES = (FILE_DIR/"../results/analysis_modules.xlsx").resolve()

ARTICLE_TYPES = {"Article", "Article in Press"}
CONFERENCE_TYPES = {"Conference Paper"}

"""
Coauthorship network graphs analysis.

For each graph the following information about the nodes is listed:
- Author's full name
- Faculty and department where author is employed
- Total number of scientific papers published
- Total number of co authorships
- Average number of coauthors per paper

From this the average number of coauthors per paper and author for the whole graph is calculated.
Results are written to the separate worksheets for each graph inside the Excel file.
"""
def authors_analysis():
    print("\nInitializing co-authorship network graphs analysis")
    wb = xlsxwriter.Workbook(EXCEL_AUTHORS)

    # column arguments - header , width, format
    column_args = [
        ["Autor", 26, None],
        ["Modul", 10, None],
        ["Broj radova", 12, {'align' : "center"}],
        ["Broj koautora", 14, {'align' : "center"}],
        ["Prosečan broj koautora po radu", 28, {'align' : "center"}],
        ["Prosečan broj koautora po radu za {}", 36, {'align' : "center"}],
        ["Prosečan broj koautora po autoru za {}", 38, {'align' : "center"}]
    ]
    sheet_names = ["UB", "ETF", "MATF", "FON"]
    graph_files = [GRAPH_INPUT, ETF_INPUT, MATF_INPUT, FON_INPUT]

    for i in range(len(graph_files)):
        print("Reading graph file: " + str(graph_files[i]))
        graph = nx.read_gexf(graph_files[i])
        sheet = wb.add_worksheet(sheet_names[i])

        # set column headers, widths and formats
        for col in range (0, len(column_args)):
            header = column_args[col][0].format(sheet_names[i]) # update last headers with sheet name
            width = column_args[col][1]
            cell_format =  wb.add_format(column_args[col][2])

            sheet.write(0, col, header)
            sheet.set_column(col, col, width, cell_format) # set width and format

        # sort authors by paper count
        nodes = sorted(graph.nodes(data = True),
                       key = lambda node : node[1]['count'], reverse = True)

        avg_coauths_paper = 0
        avg_coauths_author = 0
        for j in range(0, len(nodes)):
            node = nodes[j]
            author = node[0]
            module = node[1]['module']
            papers = node[1]['count']
            coauthors = graph.degree(author, weight = 'weight')
            avg_coauths = (coauthors / papers) if (papers > 0) else 0

            avg_coauths_paper += avg_coauths
            avg_coauths_author += coauthors

            row = [author, module, papers, coauthors, "{:.2f}".format(avg_coauths)]
            sheet.write_row(j + 1, 0, row)

        # calculate average number of coauthors per paper in the graph
        avg_coauths_paper /= len(nodes)
        sheet.write(1, len(column_args) - 2, avg_coauths_paper)

        # calculate average number of coauthors per author in the graph
        avg_coauths_author /= len(nodes)
        sheet.write(1, len(column_args) - 1, avg_coauths_author)

        print("Node information succesfully written to worksheet: " + sheet_names[i])

    print("Analysis results written to Excel file: " + str(EXCEL_AUTHORS))
    wb.close()
    del wb

"""
Analysis of scientific production per module - department and faculty.
Scientific production is analyzed separately for different paper types.

Results are writtent to the following tables in Excel file:
- Tables with total number of scientific papers employees published per module.
  Columns for all types, only articles and only conference papers.

- Tables with number of papers published per year for each module.
"""
def modules_analysis():
    print("\nInitializing modules scientific production analysis")

    # Initialize module/faculty databases
    modules = ['ETF_RTI', 'MATF_RTI', 'FON_SI', 'FON_IT', 'FON_IS']
    database_m = dict()
    for module in modules:
        database_m[module] = {
            "papers" : 0,
            "articles" : 0,
            "conferences" : 0,
            "years" : dict()
        }

    faculties = ["ETF", "MATF", "FON"]
    database_f = dict()
    for faculty in faculties:
        database_f[faculty] = {
            "papers" : 0,
            "articles" : 0,
            "conferences" : 0,
            "years" : dict()
        }

    # Input
    graph = nx.read_gexf(GRAPH_INPUT)
    wb = xlrd.open_workbook(EXCEL_PAPERS, on_demand = True)
    sheet = wb.sheet_by_index(0)

    print("Reading papers data file: " + str(EXCEL_PAPERS))
    for row in range(1, sheet.nrows):
        ptype = sheet.cell_value(row, 0)
        year = sheet.cell_value(row, 1)
        authors = sheet.cell_value(row, 3).split(", ")

        # get authors modules/faculties
        module_set = set()
        faculty_set = set()
        for author in authors:
            module = graph.nodes[author]['module']
            module_set.add(module)

            faculty = module[:module.find('_')] # trim department suffix from module name
            faculty_set.add(faculty)

        # update module paper counts
        for module in module_set:
            database_m[module]['papers'] += 1

            if ptype in ARTICLE_TYPES:
                database_m[module]['articles'] += 1

            if ptype in CONFERENCE_TYPES:
                database_m[module]['conferences'] += 1

            # add\update module paper count per year
            database_m[module]['years'][year] = database_m[module]['years'].get(year, 0) + 1

        # update faculty paper counts
        for faculty in faculty_set:
            database_f[faculty]['papers'] += 1

            if ptype in ARTICLE_TYPES:
                database_f[faculty]['articles'] += 1

            if ptype in CONFERENCE_TYPES:
                database_f[faculty]['conferences'] += 1

            # add\update faculty paper count per year
            database_f[faculty]['years'][year] = database_f[faculty]['years'].get(year, 0) + 1

    wb.release_resources()
    del wb

    # Output
    wb = xlsxwriter.Workbook(EXCEL_MODULES)

    # Module tables
    sheet = wb.add_worksheet("Naučna produkcija po modulima")
    column_args = [
        ["{}", 10, None],
        ["Broj radova", 12, {'align' : "center"}],
        ["Časopisi", 12, {'align' : "center"}],
        ["Konferencije", 14, {'align' : "center"}]
    ]
    table_db = [database_m, database_f]
    table_type = ["Katedra", "Fakultet"]
    table_pos = [0, 5] # deptartment/faculty table position in Excel file

    for i in range(0, len(table_db)):
        # set column headers, widths and formats
        for j in range (0, len(column_args)):
            col = j + table_pos[i]
            header = column_args[j][0].format(table_type[i])
            width = column_args[j][1]
            cell_format = wb.add_format(column_args[j][2])

            sheet.write(0, col, header)
            sheet.set_column(col, col, width, cell_format)

        # sort database by paper count
        items = sorted(table_db[i].items(),
                       key = lambda item : item[1]['papers'], reverse = True)

        row = 1 # skip header row
        for data in items:
            module = data[0]
            papers = data[1]['papers']
            articles = data[1]['articles']
            conferences = data[1]['conferences']

            sheet.write_row(row, table_pos[i], [module, papers, articles, conferences])
            row += 1
    print("Module tables created.")

    # Year tables
    sheet = wb.add_worksheet("Naučna produkcija po godinama")
    column_args = [
        ["Godina", 11, {'align' : "right"}],
        ["Broj radova", 13, {'align' : "left"}]
    ]
    table_db = [database_m, database_f]
    table_pos = [0, 11] # deptartment/faculty table position

    for i in range(0, len(table_db)):
        sub_pos = table_pos[i] # year/count subtable position

        for key in table_db[i].keys():
            # set column headers, widths and formats
            for j in range (0, len(column_args)):
                col = j + sub_pos
                header = column_args[j][0]
                width = column_args[j][1]
                cell_format = wb.add_format(column_args[j][2])

                sheet.write(1, col, header)
                sheet.set_column(col, col, width, cell_format)

            # merge header cells (row1, col1, row2, col2)
            sheet.merge_range(0, sub_pos, 0, sub_pos + 1,
                              str(key), wb.add_format({'align' : "center"}))

            # dict containing year info
            years_dict = table_db[i][key]['years']
            # sort years by paper count
            items = sorted(years_dict.items(),
                           key = lambda item : (item[1], item[0]), reverse = True)

            row = 2 # skip two header rows
            for data in items:
                year = data[0]
                count = data[1]
                sheet.write_row(row, sub_pos, [year, count])
                row += 1

            sub_pos += 2 # position of next year/count subtable
    print("Year tables created.")

    print("Analysis results written to Excel file: " + str(EXCEL_MODULES))
    wb.close()
    del wb

authors_analysis()
modules_analysis()
