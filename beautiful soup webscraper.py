import requests
from bs4 import BeautifulSoup
import csv
import os

# Define the URL of the website
url = "https://realpython.github.io/fake-jobs/"

# Send an HTTP request and fetch the HTML content of the webpage
response = requests.get(url)

# Parse the HTML using BeautifulSoup
soup = BeautifulSoup(response.content, 'html.parser')

# Find all job postings in the HTML
jobs = soup.find_all('div', class_='card-content')

# Extract details for each job
job_list = []
for job in jobs:
    title = job.find('h2', class_='title').text.strip()  # Job title
    company = job.find('h3', class_='company').text.strip()  # Company name
    location = job.find('p', class_='location').text.strip()  # Location
    link = job.find('a')['href']  # Link to job details

    # Append the job details to the list
    job_list.append([title, company, location, link])

#  Determine the directory to save the CSV file
# Try to use the directory where the Python script is located, but if _file_ is not available, use current working directory
try:
    current_directory = os.path.dirname(os.path.abspath(__file__))
except NameError:
    current_directory = os.getcwd()  # Fallback for interactive environments

# Define the CSV file path in the same directory as the Python script
csv_file = os.path.join(current_directory, "fake_python_jobs.csv")

# Write the extracted data into the CSV file
with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow(['Job Title', 'Company', 'Location', 'More Info'])
    # Write the job rows
    writer.writerows(job_list)

# Display a success message with the location of the CSV file
print(f"🎉 Scraping the website successfully finished! A total of {len(job_list)} job postings were scraped and saved to '{csv_file}' 💾")
