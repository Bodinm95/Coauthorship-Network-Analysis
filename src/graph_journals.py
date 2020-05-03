import networkx as nx
import xlrd
from pathlib import Path

FILE_DIR = Path(__file__).parent

EXCEL_AUTHORS = (FILE_DIR/"../data/UB_cs_authors.xlsx").resolve()
EXCEL_PAPERS = (FILE_DIR/"../data/UB_cs_papers_cleaned.xlsx").resolve()

GRAPH_OUTPUT = (FILE_DIR/"../results/journals_graph.gexf").resolve()

PAPER_TYPES = {"Article", "Article in Press"}

'''
Journals network graph information:

Nodes - specific journals.
Attributes:
    count = number of papers published in the journal

Edges - at least one author that published in both journals.
Attributes:
    weight = number of authors that published in both journals
'''
journal_graph = None

database = dict() # authors with lists of journals in which they published

def init_database():
    global database
    wb = xlrd.open_workbook(EXCEL_AUTHORS, on_demand = True)

    # read and insert all authors from excel file to database
    for sheet in wb.sheets():
        for row in range(1, sheet.nrows):
            name = sheet.cell_value(row, 0).title()
            lastname = sheet.cell_value(row, 1).title()
            fullname = name + " " + lastname

            database[fullname] = set() # init author entry with empty journal set

    wb.release_resources()
    del wb

def init_nodes():
    global journal_graph
    wb = xlrd.open_workbook(EXCEL_PAPERS, on_demand = True)
    sheet = wb.sheet_by_index(0)

    # initialize journal graph nodes from excel file
    for row in range(1, sheet.nrows):
        ptype = sheet.cell_value(row, 0)
        authors = sheet.cell_value(row, 3)
        journal = sheet.cell_value(row, 4).title()

        # check paper type
        if ptype not in PAPER_TYPES:
            continue

        # init/update graph node
        if journal_graph.has_node(journal):
            journal_graph.nodes[journal]['count'] += 1; # update journal paper count
        else:
            journal_graph.add_node(journal)
            journal_graph.nodes[journal]['count'] = 1;

        # update author/journals database
        for author in authors.split(', '):
            database[author].add(journal)

    wb.release_resources()
    del wb

def init_edges(journals):
    global journal_graph

    # add/update an edge for each pair of journals
    for i in range(0, len(journals)):
        for j in range(i + 1, len(journals)):
            if journal_graph.has_edge(journals[i], journals[j]):
                journal_graph[journals[i]][journals[j]]['weight'] += 1 # update existing edge weight
            else:
                journal_graph.add_edge(journals[i], journals[j], weight = 1) # add new edge

def create_graph():
    global database
    global journal_graph
    journal_graph = nx.Graph()

    print("Initializing author/journals database: " + str(EXCEL_AUTHORS))
    init_database()

    print("Initializing journals graph nodes: " + str(EXCEL_PAPERS))
    init_nodes()

    print("Initializing journals graph edges")
    for author in database:
        init_edges(list(database[author]))

    print("Journals graph generated.")
    nx.write_gexf(journal_graph, GRAPH_OUTPUT, prettyprint = True)
    print("Journals graph written to: " + str(GRAPH_OUTPUT))

create_graph()
