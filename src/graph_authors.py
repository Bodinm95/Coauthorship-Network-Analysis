import networkx as nx
import xlrd

EXCEL_AUTHORS = "../data/UB_cs_authors.xlsx"
EXCEL_PAPERS = "../data/UB_cs_papers_cleaned.xlsx"

GRAPH_OUTPUT = "../results/authors_graph.gexf"
ETF_OUTPUT = "../results/authors_graph_etf.gexf"
MATF_OUTPUT = "../results/authors_graph_matf.gexf"
FON_OUTPUT = "../results/authors_graph_fon.gexf"

'''
Coauthorship network graph information:

Nodes - specific authors.
Attributes:
    count = number of papers author published
    module = author's faculty and department label

Edges - at least one co-authored scientific paper
Attributes:
    weight = number of scientific papers written together
'''
author_graph = None

faculty_departments = {
    "matematicki fakultet" : {
        "katedra za racunarstvo i informatiku" : "MATF_RTI"
    },
    "elektrotehnicki fakultet" : {
        "katedra za racunarsku tehniku i informatiku" : "ETF_RTI"
    },
    "fakultet organizacionih nauka" : {
        "Katedra za informacione sisteme" : "FON_IS",
        "Katedra za softversko inzenjerstvo" : "FON_SI",
        "Katedra za informacione tehnologije" : "FON_IT"
    }
}

def init_nodes():
    global author_graph
    wb = xlrd.open_workbook(EXCEL_AUTHORS, on_demand = True)

    # initialize author graph nodes from excel file
    for sheet in wb.sheets():
        for row in range(1, sheet.nrows):
            name = sheet.cell_value(row, 0).title()
            lastname = sheet.cell_value(row, 1).title()
            author = name + " " + lastname

            faculty = sheet.cell_value(row, 4)
            department = sheet.cell_value(row, 3)
            module = faculty_departments[faculty][department]

            author_graph.add_node(author)
            author_graph.nodes[author]['count'] = 0
            author_graph.nodes[author]['module'] = module

    wb.release_resources()
    del wb

def init_edges(authors):
    global author_graph
    authors = authors.split(", ")

    # add/update an edge for each pair of authors
    for i in range(0, len(authors)):
        for j in range(i + 1, len(authors)):
            if author_graph.has_edge(authors[i], authors[j]):
                author_graph[authors[i]][authors[j]]['weight'] += 1 # update existing edge weight
            else:
                author_graph.add_edge(authors[i], authors[j], weight = 1) # add new edge

        # update author paper count
        author_graph.nodes[authors[i]]['count'] += 1

def create_graph():
    global author_graph
    author_graph = nx.Graph()
    paper_set = set()

    print("Initializing authors graph nodes: " + EXCEL_AUTHORS)
    init_nodes()

    wb = xlrd.open_workbook(EXCEL_PAPERS, on_demand = True)
    sheet = wb.sheet_by_index(0)

    print("Initializing authors graph edges: " + EXCEL_PAPERS)
    for row in range(1, sheet.nrows):
        line = [sheet.cell_value(row, 0), # Type
                sheet.cell_value(row, 1), # Year
                sheet.cell_value(row, 2), # Title
                sheet.cell_value(row, 3), # Authors
                sheet.cell_value(row, 4)] # Document Name

        # Check duplicate papers
        if line[2].lower() in paper_set:
            continue
        else:
            paper_set.add(line[2].lower())

        init_edges(line[3])
    print("Authors graph generated.")

    nx.write_gexf(author_graph, GRAPH_OUTPUT, prettyprint = True)
    print("Authors graph written to: " + GRAPH_OUTPUT)

    wb.release_resources()
    del wb

def create_subgraphs():
    global author_graph

    faculties = ["ETF", "MATF", "FON"]
    modules = [{"ETF_RTI"}, {"MATF_RTI"}, {"FON_IT", "FON_IS", "FON_SI"}]
    outputs = [ETF_OUTPUT, MATF_OUTPUT, FON_OUTPUT]

    for i in range(len(faculties)):
        print("\nGenerating " + faculties[i] + " authors subgraph")
        nodes = [node for (node, attr) in author_graph.nodes(data = True)
                                       if attr['module'] in modules[i]]
        subgraph = author_graph.subgraph(nodes)
        nx.write_gexf(subgraph, outputs[i], prettyprint = True)
        print(faculties[i] + " authors subgraph writen to: " + outputs[i])

create_graph()
create_subgraphs()
