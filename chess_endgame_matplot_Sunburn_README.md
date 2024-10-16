# Matplot-Sunburn-King-and-Rook-vs.-King-and-Pawn

This project includes the following files:

- **Matplot & Sunburn_King & Rook vs. King & Pawn.py**: The main Python script for analyzing chess endgames. This script loads the dataset, performs analysis, visualizes outcomes, and applies a Random Forest classifier to predict results.

- **kr-vs-kp.data**: The dataset containing the chess endgame positions, used for analysis.

- **kr-vs-kp.names**: A file that provides the names and descriptions of the attributes present in the dataset.

The files are located inside the folder **Chess (King-Rook Vs King-Pawn)**.

This is a Python project where an analysis of chess endgames, with a special focus on King and Rook vs. King and Pawn, is performed. Following the loading of a dataset, the script tallies main statistics and visualizes the game outcomes using Matplotlib and Seaborn. Applying a Random Forest classifier depicts the results of a game; conclusions are illustrated.

## Table of Contents
- [Dependencies](#dependencies)
- [Installation](#installation)
- [Running the Code](#running-the-code)
- [Editing File Location](#editing-file-location)
- [Code Explanation](#code-explanation)
- [Visualizations](#visualizations)
- [Conclusion](#conclusion)

## Dependencies

This project requires the following Python libraries:

- **numpy**: For numerical operations and data manipulation.
  ```bash
  pip install numpy


- **pandas**: For data analysis and manipulation.
```bash
  pip install pandas
  ```

- **matplotlib**: For creating static, animated, and interactive visualizations in Python.
  ```bash
  pip install matplotlib
  ```

- **seaborn**: A data visualization library built on top of Matplotlib that makes drawing attractive statistical graphics easy, using a high-level interface.
  ```bash
  pip install seaborn
  ```

- **scikit-learn**: A machine learning library for Python, providing simple and efficient tools for data mining and data analysis.
  ```bash
  pip install scikit-learn
  ```

- **tracemalloc**: A built-in Python library for tracing memory usage.
  ```bash
  # Installed by default, so not included here
  ```

## Installation

1. **Clone the Repository**: 
   Clone this repository to your local machine using the following command:
   ```bash
   git clone https://github.com/Diocese-trees/Matplot-Sunburn-King-and-Rook-vs.-King-and-Pawn-
   ```

2. **Navigate to the Project Directory**: 
   Open your terminal and navigate to the project directory:
   ```bash
   cd "Chess Endgame Analysis with Matplot Sunburn"
   ```

3. **Install the Required Dependencies**: 
   Use pip to install the necessary libraries:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Code

1. Ensure that your dataset file (`kr-vs-kp.data`) is located in the `Chess (King-Rook Vs King-Pawn)` folder of the project directory.

2. Run the Python script with the command:
   ```bash
   python "discreet dataset matplotlib V2.py"
   ```

## Editing File Location

To change the file location for the dataset:

1. Open the Python script in a text editor or IDE, `discreet dataset matplotlib V2.py`.

2. Locate the following line in your code:
   ```python
   data_path = r"C:\Users\adeel\OneDrive\Documents\Internship\Chess Endgame Analysis with Matplot Sunburn\Chess (King-Rook Vs King-Pawn)"
   ```

3. Replace it with your new location for the dataset.

4. Save the edit and run it again.

## Code Explanation

This project analyzes chess endgame positions, specifically focusing on King and Rook versus King and Pawn. The script includes several key components:

### Data Loading
- The dataset is loaded using the Pandas library, handling exceptions to ensure the script runs smoothly.

### Data Analysis
- The code provides a comprehensive summary of the dataset, including statistical descriptions, class distributions, and frequencies of categorical features.
- It checks class distribution to evaluate the balance between winning and losing positions.

### Visualizations
Various types of visualizations are created using Matplotlib and Seaborn, such as:
- **Bar Plots**: Show the count of winning and losing positions.
- **Pie Charts**: Illustrate the distribution of outcomes as a percentage.
- **Dot Plots**: Depict the position of the White King and the outcome (won or lost) with added jitter for clarity.
- **Box Plots with Swarm Plots**: Provide insights into the distribution of the White King's position based on the outcome.

### Model Training
- A Random Forest classifier predicts game results based on encoded features of the dataset.
- The accuracy of the model is evaluated, providing performance metrics such as precision, recall, and F1-score.

### Findings Documentation
- Key insights from the analysis are documented, focusing on important features and strategies that enhance winning chances in similar chess endgames.

### Memory Usage Monitoring
- The script tracks memory usage during execution to ensure efficient resource management.

## Visualizations

The visualizations generated include:

1. **Bar Plot**: Displays the count of winning and losing positions, providing an overview of the class distribution.
   
2. **Pie Chart**: Shows the distribution of outcomes as a percentage, helping visualize the proportion of winning positions.
   
3. **Dot Plot**: Depicts the position of the White King against the game outcome, making it easy to see how different positions correlate with wins or losses.
   
4. **Box Plot with Swarm Plot**: Combines summary statistics and individual data points to provide a comprehensive view of the White King's position by outcome.

These visualizations enhance understanding of the dataset's characteristics and illustrate how different features impact a game's outcome.

## Conclusion

This project provides a detailed analysis of chess endgame positions and demonstrates how to use Python for data analysis and visualization. The incorporation of machine learning techniques allows for predictive modeling of game outcomes, while comprehensive visualizations facilitate better understanding and strategic insights.
```
