#Advanced Excel Data Analyzer


##Description: This is a desktop application built using Python, Tkinter, and various data analysis libraries (Pandas,
               Matplotlib, Seaborn) that allows users to upload Excel files, clean data, and generate visualizations
               (Bar, Line, Pie charts) based on the selected columns. The app provides tools to aggregate data
               based on category and numeric values, then visualize the results.

##Features:
      - Upload Excel Files: Upload .xlsx, .xls, or .csv files.
      - Data Cleaning: Option to clean data by removing rows with missing values.
      - Aggregation: Aggregates data based on category (e.g., count, unique values) and numeric column
        (e.g., sum, mean).
      - Visualizations: Choose from bar, line, or pie charts to visualize the aggregated data.
      - Summary Output: Automatically saves an aggregated summary of the analysis as a CSV file.

##Requirements:
      - Python 3.x
      - pandas
      - matplotlib
      - seaborn
      - openpyxl (for reading .xlsx files)
      - tk (for GUI)

You can install the required libraries using the following command:
```
pip install -r requirements.txt
```

##How to Use:
  1. Run the application:
     - Download or clone this repository.
     - Open a terminal/command prompt.
     - Navigate to the project folder.
     - Run the following command:
 
 ```
 python app.py
 ```

     - The application window will open.
  2. Upload Excel File:
     - Click on the "Select Excel File" button to choose an Excel or CSV file from your system.
  3. Select Columns:
     - Select the appropriate category column and numeric column from the dropdown lists.
     - Choose the aggregation methods for both the category and numeric columns.
  4. Select Graph Type:
     - Choose the type of graph (Bar, Line, Pie) to visualize the aggregated data.
  5. Clean and Analyze:
     - Optionally, check the "Clean Data" box to remove rows with missing values.
     - Click "Analyze & Visualize" to generate the graph and save the summary as a CSV file.
Example:
   Once the analysis is complete, a graph will be displayed and a CSV summary file
   (aggregated_summary.csv) will be saved in the project directory.

