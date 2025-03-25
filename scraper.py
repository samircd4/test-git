"""
BCP Events Scraper with API key authentication
Fetches tabletop gaming events from Best Coast Pairings API and appends new entries to an Excel file.
"""

from rich import print
import json
import requests
from datetime import datetime, timedelta
import pandas as pd
import os
import logging
from zoneinfo import ZoneInfo

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("events_scraper.log"),  # Save logs to file
        logging.StreamHandler()  # Show logs in console
    ]
)

def validate_api_key(api_key):
    """
    Validate the API key by making a test request to the API.
    
    Args:
        api_key (str): The API key to validate.
    
    Returns:
        bool: True if the API key is valid, False otherwise.
    """
    url = 'https://newprod-api.bestcoastpairings.com/v1/events'
    headers = {
        "accept": "*/*",
        "client-id": "web-app",
        "authorization": f"Bearer {api_key}",  # Add API key to headers
    }
    
    try:
        response = requests.get(url, headers=headers, params={"limit": 1})
        if response.status_code == 200:
            logging.info("API key is valid")
            return True
        else:
            logging.error(f"API key validation failed: {response.status_code} - {response.text}")
            return True
    except Exception as e:
        logging.error(f"API key validation error: {str(e)}")
        return False

def get_events(days_window=90):
    """
    Fetch events from BCP API with dynamic date range.
    
    Args:
        api_key (str): The API key for authentication.
        days_window (int): Number of days from today to include in search.
    
    Returns:
        list: Dictionary of event records.
    """
    url = 'https://newprod-api.bestcoastpairings.com/v1/events'
    
    # Define timezones
    est_tz = ZoneInfo("America/New_York")  # EST/EDT timezone
    utc_tz = ZoneInfo("UTC")  # UTC timezone for API

    # Calculate dynamic date range in EST
    current_est = datetime.now(est_tz)
    logging.info(f"Current EST time: {current_est.strftime('%Y-%m-%d %H:%M:%S %Z')}")

    # Start date is today at midnight in EST
    start_date_est = current_est.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date_est = start_date_est + timedelta(days=days_window)

    # Log the date range in EST
    logging.info(f"Fetching events between {start_date_est.strftime('%Y-%m-%d')} and "f"{end_date_est.strftime('%Y-%m-%d')} (EST)")

    # Convert EST dates to UTC for API parameters
    start_date_utc = start_date_est.astimezone(utc_tz)
    end_date_utc = end_date_est.astimezone(utc_tz)
    
    # API parameters with UTC-converted dates
    params = {
        'limit': '40',
        'startDate': start_date_utc.isoformat(timespec='seconds').replace('+00:00', 'Z'),  # ISO format with Zulu time
        'endDate': end_date_utc.isoformat(timespec='seconds').replace('+00:00', 'Z'),
        'sortKey': 'eventDate',
        'sortAscending': 'true',
        'sortAsc': 'true',
        'sortType': 'Start Date Ascending',
    }
    
    headers = {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,bn;q=0.8",
        "client-id": "web-app",
        "content-type": "application/json",
        # "authorization": f"Bearer {api_key}",  # Add API key to headers
        "env": "bcp",
        "origin": "https://bestcoastpairings.com.",
        "priority": "u=1, i",
        "referer": "https://bestcoastpairings.com./",
        "sec-ch-ua-mobile": "?0",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "cross-site",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36"
    }
    
    results = []
    page_num = 1
    
    try:
        while True:
            # API request with error handling
            try:
                response = requests.get(url, params=params, headers=headers)
                response.raise_for_status()  # Raise exception for 4xx/5xx status
            except requests.exceptions.RequestException as e:
                logging.error(f"API request failed: {str(e)}")
                break

            # Parse JSON response
            try:
                data_str = json.loads(response.text)
            except json.JSONDecodeError:
                logging.error("Failed to parse JSON response")
                break

            next_key = data_str.get('nextKey')
            params['nextKey'] = next_key
            data = data_str.get('data', [])
            
            if not data:
                logging.info("No more data in response")
                break

            # Process event records
            for d in data:
                try:
                    # Original ISO date string
                    iso_start_date = d.get('eventDate')

                    # Parse the ISO string
                    start_date_obj = datetime.strptime(iso_start_date, "%Y-%m-%dT%H:%M:%S.%fZ")

                    # Format date and time
                    start_date = start_date_obj.strftime("%d/%m/%Y")  # dd/mm/yyyy
                    temp_time = start_date_obj.strftime("%I:%M %p")  # hh:mm AM/PM
                    start_time = temp_time[1:] if temp_time.startswith('0') else temp_time
                    
                    # Original ISO date string
                    iso_end_date = d.get('eventEndDate')

                    # Parse the ISO string
                    end_date_obj = datetime.strptime(iso_end_date, "%Y-%m-%dT%H:%M:%S.%fZ")

                    # Format date and time
                    end_date = end_date_obj.strftime("%d/%m/%Y")  # dd/mm/yyyy
                    temp_time = end_date_obj.strftime("%I:%M %p")  # hh:mm AM/PM
                    end_time = temp_time[1:] if temp_time.startswith('0') else temp_time
                    
                    
                    event = {
                        'event_id': d.get('id'),
                        'name': d.get('name'),
                        'game_system': d.get('gameSystemName'),
                        'start_date': start_date,
                        'end_date': end_date,
                        'start_time': start_time,
                        'end_time': end_time,
                        'location': f"{d.get('city', '')} {d.get('state', '')} {d.get('country', '')}".strip(),
                        'street': f"{d.get('streetNum', '')} {d.get('streetName', '')}".strip(),
                        'city': d.get('city'),
                        'state': d.get('state'),
                        'country': d.get('country'),
                        'owner_name': f"{d.get('ownerFirstName', '')} {d.get('ownerLastName', '')}".strip(),
                        'img_url': d.get('photoUrl'),
                        'event_link': f'https://bestcoastpairings.com/event/{d.get("id")}'
                    }
                    results.append(event)
                    print(event)
                except Exception as e:
                    logging.error(f"Error processing event {d.get('id')}: {str(e)}")
                    continue

            logging.info(f"Processed page {page_num} with {len(data)} events")
            page_num += 1

            # Break conditions
            if not next_key:
                logging.info("Reached end of pagination")
                break
            # else:
            #     break  # Remove this line to fetch all pages

    except Exception as e:
        logging.error(f"Critical error in get_events: {str(e)}")
        return []

    return results



def remove_csv_file(filename: str) -> bool:
    """
    Remove a CSV file if it exists.
    
    Args:
        filename (str): Name of the CSV file to remove
        
    Returns:
        bool: True if file was removed successfully, False otherwise
    """
    try:
        if os.path.exists(filename):
            os.remove(filename)
            logging.info(f"Successfully removed {filename}")
            return True
        else:
            logging.info(f"File {filename} does not exist")
            return False
    except Exception as e:
        logging.error(f"Error removing {filename}: {str(e)}")
        return False



def get_new_events():

    filename = 'Events_data.xlsx'
    existing_data = pd.DataFrame()
    
    try:
        # Load existing data if available
        if os.path.exists(filename):
            existing_data = pd.read_excel(filename)
            logging.info(f"Loaded existing data with {len(existing_data)} records")
        else:
            logging.info("No existing data file found - creating new one")

        # Fetch new events
        new_events = get_events(days_window=90)  # 3 month window
        new_df = pd.DataFrame(new_events)
        
        if new_df.empty:
            logging.info("No new events found in API response")
            return

        # Filter new entries
        if not existing_data.empty:
            existing_ids = existing_data['event_id'].unique()
            new_df = new_df[~new_df['event_id'].isin(existing_ids)]
            
        if not new_df.empty:
            # Save combined data
            combined_df = pd.concat([existing_data, new_df], ignore_index=True)
            combined_df.to_excel(filename, index=False)
            logging.info(f"Added {len(new_df)} new events to {filename}")
            print(f"[bold green]Success: Added {len(new_df)} new events[/bold green]")
            new_df.to_excel('new_events.xlsx', index=False)
        else:
            logging.info("No new events to add")
            print("[bold yellow]No new records found[/bold yellow]")

    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")
        print("[bold red]Script failed - check logs for details[/bold red]")

if __name__ == '__main__':
    get_new_events()