import os
import re
import requests
import csv
from dotenv import load_dotenv
from serpapi import GoogleSearch

# Load environment variables
load_dotenv()
api_key = os.getenv('SERPAPI_KEY')

# Search parameters
params = {
    "q": "@gmail.com, dental",
    "engine": "google",
    "location": "Zapopan",
    "hl": "en",
    "gl": "us",
    "num": 120,
    "api_key": api_key
}

# Perform search
search = GoogleSearch(params)
result = search.get_dict()

def fetch_url_content(url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.text
    except requests.RequestException:
        return ""
    return ""

def extract_pattern_from_text(pattern, text):
    return set(re.findall(pattern, text))

def extract_emails_from_text(text):
    email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    emails = extract_pattern_from_text(email_pattern, text)
    return set(list(emails)[:4])  # Limit to max 4 emails

def extract_phone_numbers_from_text(text):
    phone_pattern = r"\(\d{3}\)\s?\d{3}[-.\s]?\d{4}"
    phone_numbers = extract_pattern_from_text(phone_pattern, text)
    return set(list(phone_numbers)[:4])  # Limit to max 4 phone numbers

def extract_dentist_names_from_text(text):
    dentist_name_pattern = r"(Dr\.\s?[A-Za-z]+[A-Za-z]*\s[A-Za-z]+|[A-Za-z]+\s[A-Za-z]+,\sDMD)"
    dentist_names = extract_pattern_from_text(dentist_name_pattern, text)
    return set(list(dentist_names)[:4])  # Limit to max 4 dentist names

def extract_addresses_from_text(text):
    address_pattern = r"\d{1,5}\s[A-Za-z0-9\s,.-]+(?:\s[A-Za-z]{2,}){1,2}\s\d{5,9}"
    return extract_pattern_from_text(address_pattern, text)

def process_search_results(results):
    data = []
    if "organic_results" in results:
        for item in results["organic_results"]:
            title = item.get('title', 'No title')
            link = item.get('link', '')

            if link:
                content = fetch_url_content(link)
                if content:
                    dentist_names = extract_dentist_names_from_text(content)
                    emails = extract_emails_from_text(content)
                    phone_numbers = extract_phone_numbers_from_text(content)
                    addresses = extract_addresses_from_text(content)

                    row = {
                        "Title": title,
                        "Website URL": link,
                        "Address": ", ".join(addresses) if addresses else "N/A"
                    }

                    for i, name in enumerate(dentist_names, start=1):
                        row[f"Dentist Name {i}"] = name

                    for i, email in enumerate(emails, start=1):
                        row[f"Email {i}"] = email

                    for i, phone in enumerate(phone_numbers, start=1):
                        row[f"Phone Number {i}"] = phone

                    data.append(row)

    return data

# Process search results
data = process_search_results(result)

# Collect all possible column headers in order
csv_columns = ["Title", "Website URL", "Address"]
additional_columns = set()

for row in data:
    additional_columns.update(row.keys())

# Remove predefined columns and sort the rest to maintain order
additional_columns = sorted(additional_columns - set(csv_columns))
csv_columns.extend(additional_columns)

# Save data to CSV file
csv_file = "Zapopan_MEXICO.csv"

try:
    with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
        writer.writeheader()

        for row in data:
            # Fill missing columns with "N/A"
            ordered_row = {column: row.get(column, "N/A") for column in csv_columns}
            writer.writerow(ordered_row)

    print(f"Data successfully written to {csv_file}")
except IOError:
    print("I/O error")