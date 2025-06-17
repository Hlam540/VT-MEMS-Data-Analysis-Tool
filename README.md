# VT-MEMS Data Analysis Tool

A Python-based tool developed for the Virginia Tech MEMS Lab to automate data analysis for cyclone penetration efficiency (PE) experiments. This project streamlines workflows by replacing manual Excel processing with a structured, script-driven pipeline.

---

## Purpose

- Automate data extraction and averaging from GRIMM and SEMS instruments  
- Compute Penetration Efficiency (PE) using upstream/downstream concentration data  
- Visualize results with a semilog plot of particle diameter vs. PE  
- (In Progress) Build a user-friendly GUI to simplify usage for researchers  

---

## Features

- Supports GRIMM (0.25–35 µm) and SEMS (0.03–0.75 µm) Excel datasets  
- Configurable time windows for upstream and downstream intervals  
- Dynamic correction factor input for calibration  
- Interpolates size bins and aligns data between instruments  
- Outputs clean PE tables and semilog graphs  
- Modular design using reader and analyzer classes  

---

## How It Works

1. Load Excel data files from GRIMM and SEMS instruments  
2. Define upstream and downstream measurement periods  
3. Input correction factors (e.g., `1375/800`)  
4. Tool calculates and displays PE  
5. Semilog plot is generated for visualization  

---

## Dependencies

Install the required Python packages:
- Python
- pandas
- matplotlib
- openpyxl