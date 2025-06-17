# VT-MEMS-Data-Analysis-Tool
This project is a Python-based data analysis tool developed for the Virginia Tech MEMS Lab to automate the processing of experimental data from cyclone penetration efficiency (PE) experiments. It streamlines the workflow by replacing manual data handling with a structured and reproducible pipeline that integrates with Excel datasets.

Purpose
The goal of this tool is to:

Automate data extraction, averaging, and interpolation from GRIMM and SEMS instrument outputs.

Calculate penetration efficiency using upstream and downstream particle concentration data.

Visualize PE across particle size bins using semilog plots.

(In Progress) Build a GUI interface for easier data input and user interaction.

Features
Support for GRIMM (0.25–35 µm) and SEMS (0.03–0.75 µm) instrument formats.

Clean integration with Excel files for seamless analysis.

Configurable time windows for upstream/downstream data.

Dynamic correction factor input.

Clear PE output tables and semilog graph generation.

Modular design using Reader and Efficiency classes for reusability.

How It Works
Load Excel files from SEMS and GRIMM.

Select time windows for upstream and downstream measurements.

Enter correction factors as needed.

Automatically compute penetration efficiency.

Display results and plot PE vs. particle diameter.

Requirements
Python

pandas

matplotlib

openpyxl

Future Plans
Implement a GUI interface for non-programmer usability.

Support multiple file formats and batch processing.

Add export options for processed data and plots.