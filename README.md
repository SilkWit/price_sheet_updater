# price_sheet_updater
A python project to create a supplier price database that pulls new prices from inserted excel sheets.

# Project Summary
This project was a precursor to my larger roofing estimator tool, and my first real project as I was beginning my Python journey. The goal was to create a small drag-and-drop interface for excel supplier material sheets that would create an SQLite database, adding another file would update the database with the new information. This database could later be used for the roofing estimator tool to pull material info from.

# Problem Addressed
Manual data entry was a time-consuming and tedious process for TPC Roofing (the future client) and prone to errors.

# Solution Developed
With a desktop drag-and-drop tool the supplier data update process was streamlined, reducing manual workload and input errors.

# Teck Stack
Python – Core application logic
Tkinter + TkinterDnD2 – Graphical user interface with drag-and-drop functionality
pandas – Reading and processing Excel data
SQLite – Database storage and updating material prices

# Requirements
pandas==2.3.1
openpyxl==3.1.5
tkinterdnd2==0.4.3
