import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import os
import re
import pyautogui
import pyperclip
import shutil
import datetime
import keyboard  # Add this import for keyboard event handling
import threading
import signal
import sys

# Global flag to track if automation should stop
stop_automation = False

def keyboard_listener():
    """
    Listen for Ctrl+E keyboard combination to stop automation
    """
    global stop_automation
    print("Keyboard listener started. Press Ctrl+E to stop automation.")
    keyboard.wait('ctrl+e')
    stop_automation = True
    print("\n*** Ctrl+E detected! Stopping automation after current contact... ***")

def read_excel_data(file_path):
    """
    Read data from an Excel file and return a pandas DataFrame
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        pandas.DataFrame: Data from the Excel file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        print(f"Successfully read data from {file_path}")
        print(f"Found {len(df)} rows and {len(df.columns)} columns")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def display_data(df):
    """
    Display the data from the DataFrame
    
    Args:
        df (pandas.DataFrame): DataFrame to display
    """
    if df is not None:
        # Display the first 5 rows
        print("\nFirst 5 rows of data:")
        print(df.head())
        
        # Display column names
        print("\nColumn names:")
        print(df.columns.tolist())

def read_caption_template():
    """
    Read caption template from WHATSDRAFT.txt
    
    Returns:
        str: Caption template text
    """
    try:
        with open("WHATSDRAFT.txt", "r") as f:
            caption = f.read().strip()
        print("Successfully read caption template from WHATSDRAFT.txt")
        return caption
    except Exception as e:
        print(f"Error reading caption template: {e}")
        return ""

def log_feedback(index, total, name, phone, status):
    """
    Log feedback to feedback.txt file
    
    Args:
        index (int): Current contact index
        total (int): Total number of contacts
        name (str): Contact name
        phone (str): Contact phone number
        status (str): Status message (success or failure message)
    """
    try:
        with open("feedback.txt", "a") as f:
            f.write(f"[{index}/{total}] {name} -- {phone} : {status}\n")
    except Exception as e:
        print(f"Error logging feedback: {e}")

def log_summary(total_contacts, successful_sends, failed_sends, force_stopped=False):
    """
    Log summary to feedback.txt file
    
    Args:
        total_contacts (int): Total number of contacts processed
        successful_sends (int): Number of successful sends
        failed_sends (int): Number of failed sends
        force_stopped (bool): Whether automation was forcefully stopped
    """
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("feedback.txt", "a") as f:
            f.write("\n" + "-"*50 + "\n")
            
            if force_stopped:
                f.write(f"Automation forcefully stopped at {timestamp}\n")
            else:
                f.write(f"Automation successfully finished at {timestamp}\n")
                
            f.write(f"Total contacts: {total_contacts}\n")
            f.write(f"Successfully sent: {successful_sends}\n")
            f.write(f"Failed: {failed_sends}\n")
            
            if force_stopped:
                remaining = total_contacts - (successful_sends + failed_sends)
                f.write(f"Remaining contacts: {remaining}\n")
                
            f.write("-"*50 + "\n")
            # Add signature at the bottom right corner
            signature = "Rashed--SE@UTM"
            padding = " " * (50 - len(signature))
            f.write(f"\n{padding}{signature}\n\n")
    except Exception as e:
        print(f"Error logging summary: {e}")

def wait_for_chat_to_load(driver, timeout=45):
    """
    Wait for WhatsApp chat to load by trying multiple selectors
    
    Args:
        driver: Selenium WebDriver
        timeout: Maximum time to wait in seconds
        
    Returns:
        bool: True if chat loaded, False otherwise
    """
    print(f"Waiting for chat to load (max {timeout} seconds)...")
    start_time = time.time()
    
    # List of possible selectors for chat input field
    selectors = [
        (By.XPATH, "//div[@title='Type a message']"),
        (By.XPATH, "//div[@data-testid='conversation-compose-box-input']"),
        (By.XPATH, "//div[contains(@class, 'copyable-text') and @contenteditable='true']"),
        (By.XPATH, "//footer//div[@contenteditable='true']"),
        (By.CSS_SELECTOR, "div.copyable-text[contenteditable='true']"),
        (By.CSS_SELECTOR, "footer div[contenteditable='true']"),
    ]
    
    while time.time() - start_time < timeout:
        for selector_type, selector_value in selectors:
            try:
                # Check if the element is present
                element = driver.find_element(selector_type, selector_value)
                if element.is_displayed():
                    print(f"Chat loaded successfully! Found element with {selector_type}: {selector_value}")
                    return True
            except:
                pass
        
        # Wait a bit before trying again
        time.sleep(1)
        
        # Check if we're still on the loading page or if there's an error
        try:
            if driver.find_element(By.XPATH, "//div[contains(text(), 'Phone number shared via url is invalid.')]").is_displayed():
                print("Error: Invalid phone number")
                return False
        except:
            pass
    
    print("Timed out waiting for chat to load")
    return False

def send_image_with_caption_to_chat(driver, image_path, caption):
    """
    Send an image with caption in an already opened WhatsApp chat using clipboard
    
    Args:
        driver: Selenium WebDriver
        image_path: Path to the image file
        caption: Caption text to send with the image
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get absolute path of the image
        image_path_abs = os.path.abspath(image_path)
        print(f"Sending image: {image_path_abs}")
        print(f"With caption: {caption}")
        
        # Focus on the chat input field
        try:
            input_field = driver.find_element(By.XPATH, "//footer//div[@contenteditable='true']")
            input_field.click()
            print("Clicked on input field")
        except Exception as e:
            print(f"Could not click input field: {e}")
            return False
        
        # Copy the image to clipboard
        # For Windows, we can use a temporary approach by copying the file
        # to clipboard using pyautogui
        
        # First, copy the file to clipboard
        # This is a Windows-specific approach
        pyperclip.copy('')  # Clear clipboard
        
        # Use Windows Explorer to copy the file
        # Create a temporary script to copy the file to clipboard
        temp_script = "temp_copy_script.ps1"
        with open(temp_script, "w") as f:
            f.write(f"Add-Type -AssemblyName System.Windows.Forms\n")
            f.write(f"$filePath = '{image_path_abs}'\n")
            f.write(f"$image = [System.Drawing.Image]::FromFile($filePath)\n")
            f.write(f"[System.Windows.Forms.Clipboard]::SetImage($image)\n")
        
        # Execute the PowerShell script
        os.system(f"powershell -ExecutionPolicy Bypass -File {temp_script}")
        
        # Delete the temporary script
        os.remove(temp_script)
        
        print("Copied image to clipboard")
        
        # Wait a moment
        time.sleep(1)
        
        # Paste the image into the chat
        pyautogui.hotkey('ctrl', 'v')
        print("Pressed Ctrl+V to paste image")
        
        # Wait for image to appear
        time.sleep(2)
        
        # Type the caption
        pyperclip.copy(caption)  # Copy caption to clipboard
        pyautogui.hotkey('ctrl', 'v')  # Paste caption
        print("Added caption to image")
        
        # Wait a moment
        time.sleep(1)
        
        # Press Enter to send
        pyautogui.press('enter')
        print("Pressed Enter to send image with caption")
        
        # Wait for the message to be sent
        time.sleep(2)
        print("Image with caption sent successfully using keyboard shortcuts")
        return True
        
    except Exception as e:
        print(f"Error sending image with caption: {e}")
        return False

def open_whatsapp_chat_and_send_image(driver, phone_number, name, image_path, caption_template, index, total, max_retries=3):
    """
    Open WhatsApp Web chat with the specified phone number and send an image with caption
    
    Args:
        driver (WebDriver): Selenium WebDriver instance
        phone_number (str): Phone number to open chat with (including country code)
        name (str): Contact name to use in caption
        image_path (str): Path to the image file to send
        caption_template (str): Caption template to use
        index (int): Current contact index
        total (int): Total number of contacts
        max_retries: Maximum number of retry attempts
        
    Returns:
        bool: True if successful, False otherwise
    """
    retry_count = 0
    while retry_count <= max_retries:
        try:
            # Format the WhatsApp URL with the phone number
            whatsapp_url = f"https://web.whatsapp.com/send?phone={phone_number}"
            
            print(f"Opening WhatsApp chat with {phone_number} ({name})")
            print(f"URL: {whatsapp_url}")
            
            # Open the WhatsApp URL
            driver.get(whatsapp_url)
            
            # Wait for the chat to load
            if not wait_for_chat_to_load(driver, timeout=45):
                print(f"Chat did not load properly for {phone_number}. Retrying...")
                retry_count += 1
                continue
            
            # Replace {name} in caption template with actual name
            caption = caption_template.replace("{name}", name)
            
            # Try to send the image with caption
            if send_image_with_caption_to_chat(driver, image_path, caption):
                print(f"Successfully sent image with caption to {name} ({phone_number})")
                log_feedback(index, total, name, phone_number, "successfully sent")
                return True
            else:
                print(f"Failed to send image with caption to {name} ({phone_number}). Retrying...")
                retry_count += 1
                continue
                
        except Exception as e:
            print(f"Error: {e}")
            if retry_count < max_retries:
                print(f"Retrying... (Attempt {retry_count + 1} of {max_retries})")
                retry_count += 1
                time.sleep(5)  # Wait before retrying
            else:
                print("Maximum retries reached. Moving to next contact.")
                log_feedback(index, total, name, phone_number, f"failed - {str(e)[:100]}")
                return False
    
    log_feedback(index, total, name, phone_number, "failed after maximum retries")
    return False

def auto_detect_phone_column(df):
    """
    Automatically detect which column contains phone numbers
    
    Args:
        df (pandas.DataFrame): DataFrame to analyze
        
    Returns:
        str: Name of the column containing phone numbers
    """
    # Common column names for phone numbers
    phone_keywords = ['phone', 'mobile', 'cell', 'contact', 'tel', 'number', 'whatsapp']
    
    # Check column names first
    for col in df.columns:
        col_lower = col.lower()
        for keyword in phone_keywords:
            if keyword in col_lower:
                return col
    
    # If no column name matches, try to detect by content
    for col in df.columns:
        # Check first few values in each column
        sample = df[col].astype(str).head(5)
        
        # Count how many values look like phone numbers
        phone_count = 0
        for val in sample:
            # Remove non-digits
            digits_only = ''.join(filter(str.isdigit, val))
            # Check if it has a reasonable length for a phone number (7-15 digits)
            if 7 <= len(digits_only) <= 15:
                phone_count += 1
        
        # If most values look like phone numbers, use this column
        if phone_count >= 3:  # If at least 3 out of 5 look like phone numbers
            return col
    
    # If no suitable column found, use the first column
    return df.columns[0]

def auto_detect_name_column(df, phone_column):
    """
    Automatically detect which column contains names
    
    Args:
        df (pandas.DataFrame): DataFrame to analyze
        phone_column (str): Name of the column containing phone numbers
        
    Returns:
        str: Name of the column containing names
    """
    # Common column names for names
    name_keywords = ['name', 'contact', 'person', 'customer', 'client', 'user']
    
    # Check column names first
    for col in df.columns:
        if col == phone_column:  # Skip the phone column
            continue
            
        col_lower = col.lower()
        for keyword in name_keywords:
            if keyword in col_lower:
                return col
    
    # If no suitable column found, use the first column that's not the phone column
    for col in df.columns:
        if col != phone_column:
            return col
    
    # If there's only one column, use it
    return df.columns[0]

def process_contacts_and_send_image(df, image_path, caption_template):
    """
    Process all contacts from DataFrame, open WhatsApp chats and send image with caption
    
    Args:
        df (pandas.DataFrame): DataFrame containing contact information
        image_path (str): Path to the image file to send
        caption_template (str): Caption template to use
    """
    global stop_automation
    
    if df is None or len(df) == 0:
        print("Error: DataFrame is empty or None")
        return
    
    # Auto-detect which column contains phone numbers
    phone_column = auto_detect_phone_column(df)
    print(f"\nAutomatically selected column '{phone_column}' for phone numbers")
    
    # Auto-detect which column contains names
    name_column = auto_detect_name_column(df, phone_column)
    print(f"Automatically selected column '{name_column}' for names")
    
    # Get total number of contacts
    total_contacts = len(df)
    print(f"Processing {total_contacts} contacts...")
    
    # In the process_contacts_and_send_image function, modify the feedback file initialization:
    
    # Clear feedback file before starting
    try:
        with open("feedback.txt", "w") as f:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"WhatsApp Automation Started at {timestamp}\n")
            f.write("-"*50 + "\n\n")
            # Remove these lines that add the signature at the top
            # signature = "Rashed--SE@UTM"
            # padding = " " * (50 - len(signature))
            # f.write(f"{padding}{signature}\n\n")
    except Exception as e:
        print(f"Error clearing feedback file: {e}")
    
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    
    # Initialize the Chrome driver once
    try:
        driver = webdriver.Chrome(options=chrome_options)
        
        # First load WhatsApp Web to authenticate (if needed)
        driver.get("https://web.whatsapp.com/")
        print("\nPlease scan the QR code to log in to WhatsApp Web (if required)")
        print("Waiting 45 seconds for authentication...")
        time.sleep(45)  # Wait for authentication
        
        # Process all contacts
        successful_sends = 0
        failed_sends = 0
        
        for index, row in df.iterrows():
            # Check if stop flag is set
            if stop_automation:
                print("\nAutomation stopped by user (Ctrl+E pressed)")
                break
                
            phone = str(row[phone_column]).strip()
            # Remove any non-numeric characters
            phone = ''.join(filter(str.isdigit, phone))
            
            # Get name (or use phone number if name column doesn't exist)
            name = str(row.get(name_column, phone)).strip()
            if not name:  # If name is empty, use phone number
                name = phone
            
            if phone:
                print(f"\nProcessing contact {index+1}/{total_contacts}: {name} ({phone})")
                success = open_whatsapp_chat_and_send_image(driver, phone, name, image_path, caption_template, index+1, total_contacts)
                
                if success:
                    successful_sends += 1
                else:
                    failed_sends += 1
                
                # Brief pause between contacts
                time.sleep(2)
    except Exception as e:
        print(f"Error initializing browser: {e}")
        log_feedback(0, total_contacts, "SYSTEM", "ERROR", f"Browser initialization failed: {str(e)[:100]}")
    finally:
        # Close the browser when done
        try:
            driver.quit()
        except:
            pass
        
        # Log summary with force_stopped flag if automation was stopped by user
        log_summary(total_contacts, successful_sends, failed_sends, force_stopped=stop_automation)
        
        print(f"\nAll contacts processed. Browser closed.")
        print(f"Summary: {successful_sends} successful sends, {failed_sends} failed sends")
        if stop_automation:
            remaining = total_contacts - (successful_sends + failed_sends)
            print(f"Remaining contacts: {remaining} (automation stopped by user)")
        print(f"Detailed feedback saved to feedback.txt")

def main():
    # Start keyboard listener in a separate thread
    listener_thread = threading.Thread(target=keyboard_listener)
    listener_thread.daemon = True  # Thread will exit when main program exits
    listener_thread.start()
    
    # Path to the Excel file
    file_path = "data.xlsx"
    
    # Path to the image file
    image_path = "image.png"
    
    # Check if image exists
    if not os.path.exists(image_path):
        print(f"Error: Image file '{image_path}' not found")
        return
    
    # Read caption template
    caption_template = read_caption_template()
    if not caption_template:
        print("Warning: Caption template is empty. Proceeding without caption.")
    
    # Read the data
    data = read_excel_data(file_path)
    
    # Display the data
    display_data(data)
    
    if data is not None:
        # Process contacts and send image with caption
        process_contacts_and_send_image(data, image_path, caption_template)

if __name__ == "__main__":
    main()