import os
import psycopg2
from sqlalchemy import create_engine
import pandas as pd
import requests
import logging
from getpass import getpass
import subprocess
from datetime import datetime

from sqlalchemy.testing.config import Config

"""
This script automates the process of exporting data from a PostgreSQL database 
to a Microsoft SharePoint site and to a local folder.

The script performs the following steps:

1. Fetches data from the PostgreSQL database using a provided SQL query.
2. Saves the fetched data to an excel file.
3. Uploads the excel file to a Microsoft SharePoint site using the Microsoft Graph API.

Environment Variables:
    DB_URL: Connection URL for the PostgreSQL database.
    OAUTH_TOKEN: OAuth token to authenticate with the Microsoft Graph API.

The script uses Python's logging module to log information about each step 
and any errors that occur.

Requires:
    psycopg2: For connecting to the PostgreSQL database.
    pandas: For handling the fetched data and saving it to a CSV file.
    requests: For making HTTP requests to the Microsoft Graph API.
"""
# Set up logging
now = datetime.now()
now_str = now.strftime("%Y%m%d%H%M%S")
log_filename = f"{now_str}_log.log"
logging.basicConfig(filename=log_filename, filemode='w', format='%(name)s - %(levelname)s - %(message)s')


# Database Export
def read_sql_file(sql_file_path):
    try:
        with open(sql_file_path, 'r') as file:
            sql_file = file.read()
        return sql_file
    except Exception as e:
        print(f"Error occurred: {e}")
        return None


def fetch_data(file_name, db_url):
    """
    Connects to a PostgreSQL database using configuration from environment variables and fetches data
    using a provided SQL query.

    Args:
        file_name (str): Path to the SQL file.
        db_url (str): Connection URL for the database.

    Returns:
        A pandas DataFrame containing the fetched data, or None if an error occurs.
    """
    try:
        # Load configuration
        config = Config()
        user = config.get('DB_USER')
        password = config.get('DB_PASSWORD')
        host = config.get('DB_HOST')
        port = config.get('DB_PORT')
        database = config.get('DB_NAME')

        # Create a connection engine
        engine_url = f"postgresql://{user}:{password}@{host}:{port}/{database}"
        engine = create_engine(engine_url, pool_size=10, max_overflow=20)

        # Read the SQL file
        with open(file_name, 'r') as file:
            query = file.read()

        # Execute the query and fetch data
        data = pd.read_sql_query(query, engine)

        return data
    except Exception as e:
        logging.error("Error fetching data: %s", e)
        return None


# Save data to csv
def save_to_excel(data, filename):
    """
    Saves a pandas DataFrame to an Excel file.

    Args:
        data (pd.DataFrame): The data to save.
        filename (str): The name of the file to save the data in.
    """
    try:
        data.to_excel(filename, index=False)
    except Exception as e:
        logging.error("Error saving data to Excel: ", e)


# Upload to SharePoint
def upload_to_sharepoint(filename):
    """
    Uploads a file to SharePoint using the Microsoft Graph API.

    Args:
        filename (str): The name of the file to upload.
    """
    try:
        token = os.getenv('OAUTH_TOKEN')  # get OAuth token from an environment variable
        site_id = "your_sharepoint_site_id"
        item_path = "path/to/your/item.xlsx"
        headers = {"Authorization": f"Bearer {token}", }
        with open(filename, 'rb') as f:
            file_content = f.read()
        response = requests.put(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{item_path}:/content",
                                headers=headers, data=file_content)
        response.raise_for_status()
    except Exception as e:
        logging.error("Error uploading to SharePoint: ", e)


def main():
    """
    Main function that orchestrates the data fetching, saving data to an excel, and uploading to SharePoint.
    """
    # Get current date and time
    right_now = datetime.now()

    # Format as a string
    right_now_str = right_now.strftime("%Y-%m-%d_%H-%M-%S")

    # Use in file name
    filename = f"{right_now_str} - data_name.xlsx"
    data = fetch_data('name_of_sql_file', 'URL of DB')

    if data is not None:
        save_to_excel(data, filename)
        # TODO: Still need to test sharepoint upload functionality
        # upload_to_sharepoint('data.csv')
    else:
        logging.error("No data fetched from the database")


if __name__ == '__main__':
    main()
