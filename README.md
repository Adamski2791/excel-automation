## README for Excel Automation with Python

This repository contains Python code to automate various tasks on Excel files, specifically designed to handle patient appointment data. It demonstrates how Python can be used to manipulate and format Excel data, replacing repetitive manual work.

**Source:** This project is based on the work of Raphael Schols, who published an article on Medium titled "Stop Wasting Time in Excel: Let Python Do the Work."

**What this project does:**

* Cleans and formats patient appointment data from Excel files.
* Generates separate reports for each insurance provider within the data.
* Calculates taxes on consultation fees.
* Creates charts visualizing data.

**Project Structure:**

```
data/
├── input/ # Stores raw patient appointment data
├── output/ # Stores processed reports
│   ├── transformed/ # Stores reports for each insurance provider
├── pipeline/
│   └── pipeline.py # Python script containing all the functionalities
├── requirements.txt # Lists required Python libraries
```

**Getting Started:**

1. **Clone the Repository:** Use `git clone https://github.com/raphaelschols/excel-automation.git` to clone this repository.
2. **Install Libraries:** Run `pip install -r requirements.txt` to install the necessary Python libraries (`openpyxl` in this case).
3. **Run the Script:** Execute the `pipeline.py` script to process your data. Modify the script's parameters as needed.

**Explanation of the Script:**

The `pipeline.py` script includes various functions that perform specific tasks on the Excel data, including:

* Reading the Excel file
* Deleting rows with missing patient IDs
* Formatting data types (e.g., applying currency format)
* Adding a new column (e.g., indicating insurance status)
* Applying conditional formatting (e.g., highlighting follow-up appointments)
* Calculating taxes on consultation fees
* Generating separate workbooks for each insurance provider
* Creating a summary sheet with total fees per diagnosis
* Adding a bar chart visualizing the diagnosis fees

**Learning Resources:**

The Medium article by Raphael Schols provides a detailed explanation of each function and the overall workflow. You can find it here: [Link to the Medium article "Stop Wasting Time in Excel..."]

**Feel free to explore the code, modify it to fit your specific needs, and play around with the functionalities!**

**Additional Notes:**

* This README provides a high-level overview. Refer to the comments within the code for detailed explanations.
* Consider using a virtual environment to manage Python libraries for this project.


I hope this README helps you understand the project's purpose and functionalities!
