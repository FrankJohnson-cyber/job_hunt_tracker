# Josh's Job Tracker

A simple GUI application built with Python and `tkinter` to track job applications. It allows users to input job details, save them to an Excel file, and includes a clickable image linking to a webpage.

## Features
- **Fields**: 
  - Status (checkbox: Applied/Not Applied)
  - Date Applied
  - Job Title
  - Company
  - Link (Board)
  - Link (Company)
  - Job Description (multi-line)
  - Link (Resume)
  - Reach Out Person
  - Reach Out Status (checkbox: Contacted/Not Contacted)
  - Interview Questions (multi-line)
- **Output**: Saves data to `josh_rules.xlsx` on your Desktop with a timestamp.
- **Image**: A clickable logo in the bottom-left corner linking to [Skool Cyber Range](https://www.skool.com/cyber-range/about?ref=cc61b1b3cb11431b889d57956597cce5).

## Requirements
- Python 3.11+
- Tcl/Tk 8.6+ (included with Python on most systems)
- Libraries:
  - `openpyxl` (for Excel handling)
  - `Pillow` (for image processing)
  - `webbrowser` (standard library, for opening URLs)

## Installation
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/FrankJohnson-cyber/job_hunt_tracker.git
   cd job-tracker
