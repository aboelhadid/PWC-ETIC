# PWC-ETIC
**RPA Assessment**
This repository contains the full implementation of the RPA Assessment, developed using UiPath and Python. The solution covers data cleansing, Excel manipulation, UiAutomation, and analytics.

**Project Structure**
RPA_Assessment/
  Data/
    Input/
      Employees_Raw.xlsx               # Provided raw employee data
    Output/
      Employees_Cleaned.xlsx           # Cleaned & standardized data
      Employees_Splitted.xlsx          # Internal vs external emails
      MostActiveStocks.xlsx            # Yahoo Finance scraping output
      Summary_Report.csv               # Python analytics summary
  UiPath/
    Main.xaml                          # Main orchestrator
    Workflows/
      Part1&2_Excel Manapulation       # Excel data cleansing & Manipulation - Mails Consolidation
      Part3_ScrapeYahoo                # Yahoo Finance scraping
  python/
    PythonAnalysis.py                  # Python analytics script

**How to Run**
# UiPath #
  Open PWC-ETIC/UiPatn
    Part I & II Excel Manipulation
    Part3_ScrapeYhoo
  Run Main.xaml.
  Outputs will be generated in Data/Output/.

# Python #
  Navigate to the directory
 Ex: cd "C:\Users\mahmoud.aboelhadid\Documents\UiPath\Excel Manipulation\Excel Manipulation\Output\"
  
  Run your script
  py PythonAnalysis.py

**Time Spent**
Part 1: ~3 hrs (data cleaning & testing)
Part 2: ~3 hrs (email split)
Part 3: ~2.5 hrs (web scraping)
Part 4: ~4 hr (Python analysis)
