# Data Processing and Analysis Functions

This repository contains a series of functions designed for data processing and analysis, organized into reusable steps.

## Features

### 1. Data Collection
- Data was collected from the provided URLs using the `BeautifulSoup` library.

### 2. Data Storage
- Collected data was stored in a designated folder for easy access and further processing.

### 3. Text Cleaning
#### First Stage:
- Removed stop words from the data.
- Cleaned results were saved in a separate folder.

#### Second Stage:
- Further cleaned the data by removing punctuation.
- Results were saved in another folder.

### 4. File Naming
- Files were named based on the URLs from which they were sourced.
- These filenames are used in the Excel processing section.

### 5. Scoring
- Calculated scores based on the processed data.
- Results were stored for further analysis.

### 6. Excel Integration
- Used the `OpenPyXL` library to read and write the cleaned data and scores into an Excel file.

### 7. Code Format
- Code was initially developed in a Jupyter Notebook (`.ipynb`) for ease of testing and visualization.
- Final version was converted into a Python script (`.py`).
- **Recommendation:** Use the `.ipynb` file to observe outputs during execution.

## Important Notes

- Ensure the `|` character is removed from stop words, as it can prevent proper data reading during execution.
- Dependencies are included in the project. Please exclude any outdated dependencies before running the project.

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   ```
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the Jupyter Notebook for a step-by-step overview of the process:
   ```bash
   jupyter notebook data_processing.ipynb
   ```
2. Alternatively, execute the Python script:
   ```bash
   python data_processing.py
   ```

## Contributions

Contributions, suggestions, and feedback are welcome. Please create a pull request or open an issue to discuss your ideas.

---

This process involved extensive debugging, structuring, and testing through trial and error to ensure the desired outcome.
