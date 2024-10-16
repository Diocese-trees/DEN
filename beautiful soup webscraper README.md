## Project Title: Web Scraper for Job Postings using Python and Beautiful Soup

### Description:
This is a Python project, a scraper for a web scraper, which will scrape job postings from an HTTP website using Beautiful Soup. The extracted and stored data are the job titles, company names, locations, and links to job details in a CSV file that can easily be analyzed. It showcases the usage of web scraping techniques with Python's requests and Beautiful Soup libraries by outputting the collected data into an ordered CSV format.

### Features:
- **Handling HTTP Request**: It uses a `requests` library to fetch content - HTML from the target site
- **HTML Parsing**: Beautiful Soup will then parse the webpage and extract the relevant information about jobs in the form of job titles, company names, locations, and links to more details
- **Export to CSV**: All the scraped data is saved in a comma separated value file for further analysis or sometimes integration into other workflows of data processing.
- **Cross-environment Compatibility**: The scraper chooses the path in which it will save the file based on whether it is running in a script or interactive environment. So, whether your CSV file is supposed to be saved in a script environment or in an interactive environment, it will get it right.

### Prerequisites:
To run this web scraper, ensure that you have these installed:
- **Python 3.x**
- **Beautiful Soup**: Run the installation command: ```bash pip install beautifulsoup4```
pip install requests
```
Installation:
1. **Clone the Repository**:
Download the script or clone the repository using:
```bash
git clone <repository-url>
```
2. **Install Required Libraries**:
Run the following command to install necessary dependencies:
```bash
pip install -r requirements.txt
*(Make sure to first create a `requirements.txt` file that is empty except for the lines `requests` and `beautifulsoup4`.)*

### Use:
1. **Run the Script**:
Execute the scraper by running the following Python script:
```bash
python scraper.py
The script will write the job postings to a CSV file, `fake_python_jobs.csv`, in the directory where this script resides. The columns for the CSV file will be as follows:
Job Title: the title of the posting
Company: the company offering the job
Location: the location.
- **More**: A link to extra information on the job posting.

3. **Success Notification**:
Upon successful run, it prints out on a console a notification message showing the quantity of the scraped job postings as well as where the CSV file is saved:
```
### Example:
Here's an example of what the output CSV might look like:
| Job Title | Company | Location | More Info |
|-------------------------|----------------------------|-----------------------|------------------------------------|
| Senior Python Developer | Payne, Roberts and Davis | Stewartbury, AA | https://www.realpython.com |
| Energy engineer | Vasquez-Davidson | Christopherville, AA | https://www.realpython.com |
| Legal manager | Jackson, Chambers and Levy | Port Ericaburgh, AA | https://www.realpython.com | ]
| Gym Manager | Savage-Bradley | East Seanview, AP | https://www.realpython.com |
| Product manager | Ramirez Inc | North Jamieview, AP | https://www.realpython.com |

### Project Structure:
```
.
├── scraper.py # Main Python web scraping program
├── requirements.txt # File listing the pip dependencies if provided
└── README.md
fake_python_jobs.csv # Output CSV created by this script (when run)
```

### Personalization:
- **Target URL**: Update the url variable of the script to fetch another URL.
- **HTML Parse Logic**: Update the logic of parsing the Beautiful Soup which defines where to look in the HTML structure of the target webpage.
This is an eduction project only. Always check out a site's **robots.txt** file and **terms of service** before scraping to ensure that you will not break their policies.

---

Scrape on!