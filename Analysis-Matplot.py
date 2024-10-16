import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = r"C:\Users\adeel\Downloads\organizations-1000.xlsx"
df = pd.read_excel(file_path)

# Display first and last few rows
print("First few rows of the data:")
print(df.head())
print("\nLast few rows of the data:")
print(df.tail())

# Check for missing values
print("\nMissing values in the dataset:")
print(df.isnull().sum())

# Data types info
print("\nData types and information:")
df.info()

# 1. Bar plot for Top 10 Countries with Most Organizations
if 'Country' in df.columns:
    country_counts = df['Country'].value_counts().head(10)  # Limit to top 10 countries

    # Plot: Number of Organizations per Country
    fig1 = plt.figure(figsize=(16, 8))
    ax = country_counts.plot(kind='bar', color='teal')
    ax.set_facecolor('lightgray')
    plt.title("Top 10 Countries by Number of Organizations", fontsize=16, fontweight='bold')
    plt.xlabel('Country', fontsize=12)
    plt.ylabel('Number of Organizations', fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Set the window title for the figure
    fig1.canvas.manager.set_window_title("Figure 1: Countries by Organizations")
    plt.show()
else:
    print("'Country' column is missing.")

# 2. Bar plot for Top 10 Industries with Most Organizations
if 'Industry' in df.columns:
    industry_counts = df['Industry'].value_counts().head(10)  # Limit to top 10 industries

    # Plot: Number of Organizations per Industry
    fig2 = plt.figure(figsize=(16, 8))
    ax = industry_counts.plot(kind='bar', color='violet')
    ax.set_facecolor('black')
    plt.title("Top 10 Industries by Number of Organizations", fontsize=16, fontweight='bold')
    plt.xlabel('Industry', fontsize=12)
    plt.ylabel('Number of Organizations', fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Set the window title for the figure
    fig2.canvas.manager.set_window_title("Figure 2: Industries by Organizations")
    plt.show()
else:
    print("'Industry' column is missing.")

# 3. Line plot for Organizations Founded in Recent 30 Years
if 'Founded' in df.columns:
    try:
        # Convert 'Founded' column to numeric (year), handle errors
        df['Founded'] = pd.to_numeric(df['Founded'], errors='coerce')
        df = df.dropna(subset=['Founded'])  # Drop rows where 'Founded' is missing
        
        # Filter the last 30 years (if applicable)
        df_recent = df[df['Founded'] >= df['Founded'].max() - 30]
        founded_counts = df_recent['Founded'].value_counts().sort_index()

        # Plot: Number of Organizations Founded Over Time (last 30 years)
        fig3 = plt.figure(figsize=(18, 9))
        ax = founded_counts.plot(kind='line', marker='o', linestyle='-', color='cyan', linewidth=2)
        ax.set_facecolor('gray')
        plt.title("Organizations Founded in the Last 30 Years", fontsize=16, fontweight='bold')
        plt.xlabel('Year Founded', fontsize=12)
        plt.ylabel('Number of Organizations', fontsize=12)
        plt.grid(True, which='both', linestyle='--', linewidth=0.7)
        plt.tight_layout()

        # Set the window title for the figure
        fig3.canvas.manager.set_window_title("Figure 3: Organizations Founded Over Time")
        plt.show()
        
    except Exception as e:
        print(f"Error processing 'Founded' column: {e}")
else:
    print("'Founded' column is missing.")

# 4. Pie chart for Top 5 Industries with Most Organizations
if 'Industry' in df.columns:
    top_5_industries = df['Industry'].value_counts().head(5)

    # Plot: Pie chart of organizations by industry
    fig4 = plt.figure(figsize=(10, 7))
    plt.pie(top_5_industries, labels=top_5_industries.index, autopct='%1.1f%%', colors=['purple', 'blue', 'cyan', 'teal', 'gray'], startangle=140)
    plt.title("Distribution of Organizations by Top 5 Industries", fontsize=16, fontweight='bold')
    plt.tight_layout()

    # Set the window title for the figure
    fig4.canvas.manager.set_window_title("Figure 4: Distribution by Industry")
    plt.show()
else:
    print("'Industry' column is missing.")
