import logging
from playwright.sync_api import sync_playwright, Page
from scraper import get_new_events
import pyperclip
import os
import requests
import shutil
from rich import print
import pandas as pd
from time import sleep
from urllib3.exceptions import HeaderParsingError

# Configure logging
logging.basicConfig(
    filename='event_scraper.log', 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Disable logging for urllib3 errors (specifically HeaderParsingError)
logging.getLogger("urllib3").setLevel(logging.ERROR)

def download_image(event_id: str, url: str) -> bool:
    url = '' if url == 'nan' else url
    logging.info(f'Downloading image for event {event_id}')
    
    if not url:
        logging.info(f'No image found for event {event_id}')
        return False
    try:
        folder = 'images'
        os.makedirs(folder, exist_ok=True)  # Ensures directory exists

        file_extension = url.split(".")[-1].split("?")[0]  # Extract extension safely
        file_name = f'{event_id}.{file_extension}'
        file_path = os.path.join(os.getcwd(), folder, file_name)

        # Send GET request to the image URL
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Raise an exception for HTTP errors

        # Open the file and save the image
        with open(file_path, 'wb') as file:
            for chunk in response.iter_content(1024):
                file.write(chunk)
        
        logging.info(f"Image successfully downloaded as {file_name}")
        return True

    except HeaderParsingError:
        # Suppress HeaderParsingError and don't log it
        logging.warning(f"Header parsing error occurred for {event_id} image, but image download proceeded.")
        return False
        
    except requests.exceptions.RequestException as e:
        logging.error(f"Error downloading image for event {event_id}: {e}")
        return False

def login():
    with sync_playwright() as p:
        logging.info('Starting login process')
        browser = p.chromium.launch(headless=False, slow_mo=2000)
        context = browser.new_context()
        page = context.new_page()
        try:
            page.goto("https://www.lfgnexus.com/login")
            page.query_selector_all("a[aria-label='Member Login']")[0].click()
            page.locator("input#member_login_190-element-9").first.fill("samircd4@gmail.com")
            page.locator("input#member_login_190-element-10").first.fill("s72647D378#")
            page.locator("input#member_login_190-element-12").first.click()
            page.wait_for_timeout(5000)
            page.context.storage_state(path="state.json")
            logging.info('Login successful')
        except Exception as e:
            logging.error(f'Login failed: {e}')
        finally:
            page.close()
            context.close()
            browser.close()

def get_description(page: Page, url: str) -> str:
    logging.info(f'Fetching description from {url}')
    url = f'{url}?active_tab=overview'
    page.goto(url)
    
    try:
        # Try both possible selectors
        selectors = [
            'h6:has-text("Event Details:") + div',  # Selector for div
            'h6:has-text("Event Details:") + span'  # Selector for span
        ]
        
        desc = ""
        for selector in selectors:
            try:
                details_header = page.locator(selector)
                details_header.wait_for(state="visible", timeout=3000)
                desc = details_header.inner_html()
                if desc:  # If description is found, break the loop
                    logging.info(f'Description found for {url}')
                    break
            except Exception:
                continue  # Try next selector if this one fails

        # If no description is found
        if not desc:
            logging.warning(f'Description not found for {url}')
        
        page.wait_for_timeout(1000)
        return desc
    
    except Exception as e:
        logging.error(f'Error fetching description from {url}: {e}')
        return ""
    
    finally:
        # Close the page after getting the description
        page.close()

def post_event(data: dict, index:int):
    logging.info(f'{index}: Posting new event: {data["name"]}')
    print(data)
    
    has_image = download_image(data['event_id'], str(data['img_url']))
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, slow_mo=1000)
        context = browser.new_context(storage_state="state.json")
        page = context.new_page()
        try:
            page.goto("https://www.lfgnexus.com/account/events/add")
            
            if 'login_direct_url' in page.url:
                logging.warning('Login required before posting event')
                return 'login required'
            
            with page.expect_file_chooser() as fc_info:
                page.get_by_text("Upload Image").click()
            file_chooser = fc_info.value
            if has_image:
                file_chooser.set_files(f"images/{data['event_id']}.png")
            else:
                file_chooser.set_files(f"images/default-image.png")
            
            page.get_by_label("Yes").check()
            page.get_by_placeholder("Enter the post title").fill(f"{data['name']} ({data['game_system']})")
            page.locator('input#event_fields_322-element-15').fill(data['game_system'])
            try:
                page.get_by_label("event_fields-element-15-1").select_option(data['start_time'], timeout=1000)
            except:
                pass
            try:
                page.get_by_label("event_fields-element-15-2").select_option(data['end_time'],timeout=1000)
            except:
                pass
            page.locator("#stardatepicker").fill(data['start_date'])
            page.locator("#enddatepicker").fill(data['end_date'])
            page.get_by_label("External Web Link").fill(data['event_link'])
            page.get_by_placeholder("Enter a location: 350 Fifth").fill(data['location'])
            page.get_by_placeholder("Enter a location: 350 Fifth").press("Enter")
            # page.locator("input#pac-input").fill(data['location']).press("Enter")
            
            description = get_description(page=context.new_page(), url=data['event_link'])
            pyperclip.copy(description)
            page.locator('button#html-1').click()
            page.wait_for_timeout(1000)
            page.keyboard.press("Control+V")
            
            page.get_by_role("button", name="Save Changes").click()
            page.wait_for_timeout(5000)
            logging.info(f'Event {data["name"]} posted successfully')
            
        except Exception as e:
            logging.error(f'Error posting event: {data["event_id"]}: {data["name"]}')
            with open('error.txt', 'a') as err_file:
                err_file.write(f'{data["event_id"]}\n')
        finally:
            page.close()
            context.close()
            browser.close()
        return 'success'


def updated_data(filename:str, event_ids_to_remove: list):
    """
    Removes rows from an Excel file where the event_id matches any in the given list.
    
    :param filename: str - The path to the Excel file.
    :param event_ids_to_remove: list - A list of event_id values to remove.
    """
    try:
        # Load the existing data
        existing_data = pd.read_excel(filename)

        # Filter out rows where event_id is in the list
        new_df = existing_data[~existing_data['event_id'].isin(event_ids_to_remove)]

        # Overwrite the original file with the updated data
        new_df.to_excel(filename, index=False)

        print(f"Successfully removed {len(existing_data) - len(new_df)} events from {filename}.")
    
    except Exception as e:
        print(f"Error: {e}")


def read_error_event():
    """
    Reads event IDs from 'error.txt', deletes the file after reading, and returns the event IDs as a list.
    
    :return: list of event IDs (each line as a separate item)
    """
    try:
        with open('error.txt', 'r') as file:
            event_ids = file.read().splitlines()
        
        # Delete the file after reading
        os.remove('error.txt')
        
        return event_ids
    except FileNotFoundError:
        print("Error: 'error.txt' not found.")
        return []
    except Exception as e:
        print(f"An error occurred: {e}")
        return []


if __name__ == "__main__":
    logging.info('Starting event scraping process')
    old_file_name = 'Events_data.xlsx'
    updated_data(old_file_name, read_error_event())
    
    get_new_events()
    
    if os.path.exists('new_events.xlsx'):
        data = pd.read_excel('new_events.xlsx').to_dict('records')
        
        for index, event_details in enumerate(data, start=1):
            if index<1:
                continue
            response = post_event(event_details, index)
            
            if response == 'login required':
                logging.info('Re-attempting login')
                login()
                post_event(event_details)
    logging.info('Event scraping process completed')
