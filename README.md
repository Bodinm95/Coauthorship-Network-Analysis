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

**Primary dataset** - list of relevant authors employed at UB (*UB_cs_authors.xlsx*) and scientific paper database from index database Scopus (*UB_cs_papers_scopus.xlsx*). This is a raw dataset with possible incomplete data and data in unsuitable format.

**Secondary dataset** - Cleaned and formated dataset with only the necessary information from the primary dataset. Cleaning is done by Python script which reads and parses the primary Excel files.

**Social network graphs** - Graphs are generated from the cleaned dataset using [**NetworkX**](https://networkx.github.io/documentation/stable/) Python module. **Coauthor graph** links authors via their shared scientific paper collaborations and **journal graph** links scientific journals with shared authors that published in them.

**Graph visualisation** - Graphs are visualized using [**Gephi**](https://gephi.org/) graph software. Nodes are resized and colored depending on their attributes and clustering, while edges are resized depending on their weight.

**Analysis report** - Contains results and answers to research questions gained by analyzing and visualizing social network graph using [**Gephi**](https://gephi.org/) graph software and analysing the secondary dataset with specialized Python script.
