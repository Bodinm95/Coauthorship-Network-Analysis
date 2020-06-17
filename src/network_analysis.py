import xlrd
import xlsxwriter
import networkx as nx
from pathlib import Path

FILE_DIR = Path(__file__).parent

GRAPH_INPUT = (FILE_DIR/"../results/authors_graph.gexf").resolve()
ETF_INPUT = (FILE_DIR/"../results/authors_graph_etf.gexf").resolve()
MATF_INPUT = (FILE_DIR/"../results/authors_graph_matf.gexf").resolve()
FON_INPUT = (FILE_DIR/"../results/authors_graph_fon.gexf").resolve()

EXCEL_AUTHORS = (FILE_DIR/"../results/analysis_authors.xlsx").resolve()

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

authors_analysis()
