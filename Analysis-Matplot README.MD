Analysis-Matplot

**Project Description**
 This is a project that examines the dataset called **"1000 organizations.xlsx"**, to explore and visualize insights behind different organizations across countries and industries. Analysis will be made using the Python libraries **Pandas** and **Matplotlib**.
 
## Technologies Used

- **Python**: Programming language for data analysis.
- **Pandas**: Library mainly used for manipulation and analysis of data.
- **Matplotlib**: Library used for static, animated, interactive visualizations in Python.
- **OpenPyXL**: To read and write Excel files,

## Dataset

- **File Name**: `1000 organizations.xlsx`
- **Description**: An Excel file containing data on various organizations with the country of origin, industry type, and founding year.

## Code Overview

- **File Name**: `Analysis-Matplot.py`
- **Functionality**: The script will do the following:
  1. It will import the dataset from an Excel file.
  2. It will print out a preview of the first and last few rows of the dataset.
3. Check for missing values and prints types of data
 4. Creates visualizations
. Bar plot for the top 10 most numerous countries.
. Bar plot for the top 10 most numerous industries.
. Line plot for the organizations established in the last 30 years.
- Pie chart of organizations by the top 5 industries.

## Running the Code Instructions

### Requirements

You need to have the following packages installed:

```bash
pip install pandas matplotlib openpyxl
```

### Running the Code

#### Using Visual Studio Code (VS Code)

1. Open VS Code and go to your project directory.
2. Ensure you have your dataset: `1000 organizations.xlsx` placed at the following path for example, `C:\\Users\\adeel\\Downloads`.
3. Open the `Analysis-Matplot.py` file.
4. Execute the code
   - There is a terminal within the VS Code application by default. Open this. Using the navigation menu at the top right, locate `View` and click `Terminal`.
   - Right-click in the terminal and select `Paste`, which will paste the following command, then hit Enter.

   ```bash
   python Analysis-Matplot.py

1. Open PyCharm, create a new project or open an existing one.
2. Add the dataset `1000 organizations.xlsx` to the specified path, for example, `C:\\\Users\\\adeel\\\Downloads\`.
3. Create a new Python file called `Analysis-Matplot.py`, with the code given in that answer copied into it.
4. Run the Code:
   - Click the green "Run" button, or right-click inside the editor and select `Run 'Analysis-Matplot'`.
5. **Output:** The console will print the results, and the visualizations will open up automatically on a new window.

## Code Explanation

### 1. Importing Libraries

```python
import pandas as pd
import matplotlib.pyplot as plt
```
-Pandas is imported to manipulate and analyze the data
-Matplotlib.pyplot is used to plot the data

### 2. Reading the Excel File

```python
file_path = r"C:\Users\adeel\Downloads\1000 organizations.xlsx"
df = pd.read_excel(file_path)
```

Loads the dataset from the given path of the Excel file into the name `df` Pandas DataFrame

### 3. View of Data

```python
print("First few rows of the data:")
print(df.head())
print(" 
 Last few rows of the data:")
print(df.tail())
```
The first and last few rows of the dataset are printed in order to understand how its structure looks like.

### 4. Missing Values

```python
print(" 
 Missing values in the dataset:")
print(df.isnull().sum())
```

- This section checks for missing values in the dataset and prints the count for each column.
 
### 5. Data Type Display
 
```python
print("
Data types and information:")
df.info()
```
 
- The `.info()` function returns general information about a DataFrame: number of entries, column name, data type, and memory usage.
 
### 6. Visualization: Top 10 Countries by Most Organizations
 
```python
if 'Country' in df.columns:
    country_counts = df['Country'].value_counts().head(10)

    fig1 = plt.figure(figsize=(16, 8))
ax = country_counts.plot(kind='bar', color='teal')
    ax.set_facecolor('lightgray')
    plt.title("Top 10 Countries by Number of Organizations", fontsize=16, fontweight='bold')
    plt.xlabel('Country', fontsize=12)
    plt.ylabel('Number of Organizations', fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()

    fig1.canvas.manager.set_window_title("Figure 1: Countries by Organizations")
    plt.show()
else:
    print("'Country' column is missing.")
```

- It counts the organizations by country, checks for the column 'Country', and constructs a bar plot with the top 10 countries counted.

### 7. Visualization: Top 10 Industries with Most Organizations

```python
if 'Industry' in df.columns:
    industry_counts = df['Industry'].value_counts().head(10)

    fig2 = plt.figure(figsize=(16, 8))
    ax = industry_counts.plot(kind='bar', color='violet')
    ax.set_facecolor('black')
```
plt.title('Top 10 Industries by Number of Organizations', fontsize=16, fontweight='bold')
    plt.xlabel('Industry', fontsize=12)
    plt.ylabel('Number of Organizations', fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()

    fig2.canvas.manager.set_window_title("Figure 2: Industries by Organizations")
    plt.show()
else:
    print("'Industry' column is missing.")
END

Just as in the last chapter, this code also counts the occurrences and visualizes the top 10 industries using a bar plot.

### 8. Visualization: Organizations Founded in Last 30 Years

```python
if 'Founded' in df.columns:
    try:
        df['Founded'] = pd.to_numeric(df['Founded'], errors='coerce')
        df = df.dropna(subset=['Founded'])
```

df_recent = df[df['Founded'] >= df['Founded'].max() - 30]
        founded_counts = df_recent['Founded'].value_counts().sort_index()
        
        fig3 = plt.figure(figsize=(18, 9))
        ax = founded_counts.plot(kind='line', marker='o', linestyle='-', color='cyan', linewidth=2)
        ax.set_facecolor('gray')
plt.title("Organizations Founded in the Last 30 Years", fontsize=16, fontweight='bold')
        plt.xlabel('Year Founded', fontsize=12)
        plt.ylabel('Number of Organizations', fontsize=12)
        plt.grid(True, which='both', linestyle='--', linewidth=0.7)
        plt.tight_layout()

fig3.canvas.manager.set_window_title("Figure 3: Organizations Founded Over Time")
plt.show()
     
except Exception as e:
    print(f"Error processing 'Founded' column: {e}")
else:
    print("'Founded' column is missing.")
END

- This script cleans the 'Founded' column to only include organizations that were founded in the last 30 years and generates a line plot to represent the count of organizations founded over time.

### 9. Visualization: Top 5 Industries with Most Organizations (Pie Chart)

```python
if 'Industry' in df.columns:
    top_5_industries = df['Industry'].value_counts().head(5)
    fig4 = plt.figure(figsize=(10, 7))
    plt.pie(top_5_industries, labels=top_5_industries.index, autopct='%1.1f%%', colors=['purple', 'blue', 'cyan', 'teal', 'gray'], startangle=140)
```
plt.title("Distribution of Organizations by Top 5 Industries", fontsize=16, fontweight='bold')
    plt.tight_layout()
    
    fig4.canvas.manager.set_window_title("Figure 4: Distribution by Industry")
    plt.show()
else:
    print("'Industry' column is missing.")

This project is licensed under the MIT License. Feel free to modify and use as you need.
