This repository contains a series of functions designed for data processing and analysis, organized into reusable steps:

Data Collection: I collected data from the provided URL using BeautifulSoup.
Data Storage: The data was stored in a designated folder for easy access.
Text Cleaning (First Stage): I cleaned the data and removed stop words, saving the cleaned results in another folder.
Text Cleaning (Second Stage): I further cleaned the data by removing punctuation and saved it in a separate folder.
File Naming: Each file is named based on the URL from which it was sourced. These filenames are later used in the Excel processing section.
Scoring: I calculated scores based on the data, storing the results for further analysis.
Excel Integration: Using the OpenPyXL library, I read and wrote the cleaned data and scores into an Excel file.
Code Format: The code was initially developed in a Jupyter Notebook (.ipynb) format for ease of testing and visualization, then converted to a Python script (.py). For the best experience, please use the .ipynb file to observe the output.
Stop Word Handling: Please ensure that the ‘|’ character is removed from stop words, as it prevents proper reading of the data during execution.
Additionally, I have included my dependencies in the project; however, please exclude any outdated ones. The process took considerable effort to debug, structure, and test through trial and error to achieve the desired outcome.
