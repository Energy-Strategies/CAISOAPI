# CAISO API

A GUI-based Python tool to query and visualize data from the CAISO API. The program efficiently queries 4 different LMP reports from https://oasis.caiso.com. These include: Day Ahead Market (DAM), Hour Ahead Market (HASP), Five Minute Market (FMM), and the Real Time Market (RTM). This program queries the data (for any date range) and formats it appropriately.  

## Features
- Select date ranges
- Download and extract zipped data
- View data visualizations (e.g., top 5% zoom charts)
- Per-node monthly/hourly charts and per-node summary statistics tabs
- Export processed Excel files

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/haileyhendrickson/CAISOAPI.git
   cd CAISOAPI

## Use
Download .exe file to local machine, open it, and enjoy!

Tutorial Video (click to watch): 

[![Watch the video](https://img.youtube.com/vi/L6PMC6PBSPg/hqdefault.jpg)](https://www.youtube.com/watch?v=L6PMC6PBSPg)


## Running on a Local Machine:
- python -m venv venv  # create a new env
- source venv/Scripts/activate  # activate env
- pip install -r requirements.txt  # install dependencies


## Notes for use 
- Separate multiple nodes with a comma (e.g., NODE1,NODE2,NODE3).
- Each report query returns the raw data, monthly averages, hourly averages, and summary statistics.
- With multiple nodes, each node also gets its own monthly/hourly charts and a dedicated summary statistics tab, alongside the combined group view.
- Please be patient; some reports, especially RTM, may take longer to process.
- Data is processed entirely in memory — no temporary files are created on disk.


## Version History -- Last updated: September 2026 (version 4.0) 
- v1.0: Initial release - basic report retrieval
- v2.0: Added summary and analysis pages
- v2.1: Tweaked analysis pages for improved insights
- v2.2: Added data visualizations
- v3.0: Added progress tracking and improved code efficiency — removed partial file saves, allowing partial data to exist in memory
- v4.0: Faster Excel export (workbook assembled in memory and saved once, instead of repeated re-saves); per-node monthly/hourly charts and per-node summary statistics tabs; descriptive output filenames (market + date range + timestamp); bug fixes (Greenhouse Gas column no longer dropped, removed stray index column)

## License
This project is licensed under the [MIT License](LICENSE).
