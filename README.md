# Coauthorship Network Analysis

This reposirory is a project assignment for **University of Belgrade School of Electrical Engineering** masters course **Social Network Analysis (ASM)**. The full documentation is located in the [*doc/ASM_PZ2_1920.pdf*](doc/ASM_PZ2_1920.pdf) file while the analysis report is located in the [*doc/ASM_report.pdf*](doc/ASM_report.pdf) file.

## Requirements
Project is written in **Python 3.8.2** and uses the following modules:
* [**xlrd**](https://xlrd.readthedocs.io/en/latest/) - for reading Excel files.
* [**XlsxWriter**](https://xlsxwriter.readthedocs.io/) - for writing and creating Excel files.
* [**NetworkX**](https://networkx.github.io/) - for creating and generating graphs.

Project uses [**Gephi**](https://gephi.org/) for graph visualisation and analysis.
## Goal

The project assignment is quantitative and qualitative analysis of scientific production in the field of computer science at **University of Belgrade (UB)** at three most important faculties in that field (**ETF**, **MATF**, **FON**).

The goal of the analysis is to determine the state of computer science at individual faculties, as well as the level of cooperation between employees from the same and different faculties through appropriate bibliometric and scientometric analysis as well as through analysis of collaborative social network.

## Analysis steps

The analysis includes **gathering**, **processing** and **preliminary analysis** of primary raw dataset, **extracting** the necessary data and **modeling** the problem with a network of appropriate type. <br>
The modeled network is then visualized and analyzed by graph and social network processing tool, while the results and answers to research questions are documented in analysis report.

**Primary dataset** - list of relevant authors employed at UB (`UB_cs_authors.xlsx`) and scientific paper database from index database Scopus (`UB_cs_papers_scopus.xlsx`). This is a raw dataset with possible incomplete data and data in unsuitable format.

**Secondary dataset** - Cleaned and formated dataset with only the necessary information from the primary dataset. Cleaning is done by Python script which reads and parses the primary Excel files.

**Social network graphs** - Graphs are generated from the cleaned dataset using [**NetworkX**](https://networkx.github.io/documentation/stable/) Python module. **Coauthor graph** links authors via their shared scientific paper collaborations and **journal graph** links scientific journals with shared authors that published in them.

**Graph visualisation** - Graphs are visualized using [**Gephi**](https://gephi.org/) graph software. Nodes are resized and colored depending on their attributes and clustering, while edges are resized depending on their weight.

**Analysis report** - Contains results and answers to research questions gained by analyzing and visualizing social network graph using [**Gephi**](https://gephi.org/) graph software and analysing the secondary dataset with specialized Python script.

## Code structure

[**cleaner.py**](src/cleaner.py) - Python script that cleans the primary dataset. It does the following:
* removes duplicate and incomplete entries in the data,
* corrects incorrect diacritical marks,
* parses authors of scientific papers - recognizes different formats of their names, singles out only authors of interest for analysis and prints out their full names.

Output of the script is a cleaned secondary data set (`UB_cs_papers_cleaned.xlsx`) that contains only the necessary data from the primary set written to the [*data*](data) folder.

---

[**graph_authors.py**](src/graph_authors.py) - Python script that generates coauthorship network graphs from the secondary dataset for the whole **UB** and separate faculties.  

**Nodes** represent specific authors and have the following attributes:
* `count` - number of papers this author published,
* `module` - author's faculty and department label.  

**Edges** signify at least one co-authored scientific paper and have the following attribute:
* `weight` - number of scientific papers written together.  

Output of the script are the generated graph files in `.gexf` format written to the [*results*](results) folder:
* `authors_graph.gexf` - coauthorship graph for the whole **University of Belgrade**,
* `authors_graph_etf.gexf` - coauthorship graph for **Faculty of Electrical Engineering**,
* `authors_graph_matf.gexf` - coauthorship graph for **Faculty of Mathematics**,
* `authors_graph_fon.gexf` - coauthorship graph for **Faculty of Organisational Sciences**.

---

[**graph_journals.py**](src/graph_journals.py) - Python script that generates journals network graph from the secondary dataset.  

**Nodes** represent specific journals and have the following attribute:
* `count` - number of papers published in the journal.
  
**Edges** signify at least one author that published in both journals and have the following attribute:
* `weight` - number of authors that published in both journals.
  
Output of the script is the generated `journals_graph.gexf` graph file written to the [*results*](results) folder.

---

[**network_analysis.py**](src/network_analysis.py) - Python script which by analyzing the secondary dataset and generated graphs creates several Excel files in which various necessary information and analysis results are summarised. Output of the script are the following Excel files:
* `analysis_authors.xlsx` - contains information about coauthorship graph nodes as well as averages for the whole graphs,
* `analysis_modules.xlsx` - contains information about scientific production per module - department and faculty,
* `analysis_journals.xlsx` - contains information about scientific paper publishing in journals and at conferences.
