import pyautogui
import time
import cv2
import numpy as np
import mss
import os
import csv
from paths import paths
from tkinter import messagebox

def capture_screen():
    """ Capture the current screen and return it as an OpenCV image. """
    with mss.mss() as sct:
        monitor = sct.monitors[1]  # Take the first monitor
        screen_img = np.array(sct.grab(monitor))
        return cv2.cvtColor(screen_img, cv2.COLOR_RGBA2RGB)

def detect_element(screen_img, template_path, threshold=0.8):
    """ Use template matching to detect the element in the screenshot. """
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    result = cv2.matchTemplate(screen_img, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    
    if max_val >= threshold:
        return (max_loc[0], max_loc[1], 
                max_loc[0] + template.shape[1], max_loc[1] + template.shape[0])
    return None

def click_element(coords, click_type='single'):
    """ Simulate a mouse click based on click type. """
    x1, y1, x2, y2 = coords
    center_x = int((x1 + x2) / 2)
    center_y = int((y1 + y2) / 2)

    pyautogui.moveTo(center_x, center_y)
    time.sleep(0.5)

    if click_type == 'double':
        pyautogui.doubleClick()
        print(f"Double-clicked at: {center_x}, {center_y}")
    else:
        pyautogui.click()
        print(f"Clicked at: {center_x}, {center_y}")

def send_message_to_csv(content):
    """Send a message to the send.csv file."""
    send_csv_file = paths['to_send_message_path']
    try:
        with open(send_csv_file, mode='a', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([content, 'message', 'application'])  # Log the message
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write to send.csv: {e}")

def find_and_click_element(template_path, click_type='single', max_attempts=10, delay=2):
    """ Find and click the specified element. """
    consecutive_failures = 0  # Counter for consecutive failures
    for _ in range(max_attempts):
        screen_img = capture_screen()
        print("Capturing screen...")
        detected_coords = detect_element(screen_img, template_path)
        
        if detected_coords is not None:
            print(f"Element detected, clicking it...")
            click_element(detected_coords, click_type)
            return True
        
        consecutive_failures += 1
        print("No new commands to execute.")

        if consecutive_failures >= 3:
            # Send message to CSV and print finished message
            send_message_to_csv(f"Failed to find element '{template_path}' after 3 consecutive attempts.")
            print(f"Finished: Failed to find element '{template_path}' after 3 attempts.")
            return False
        
        time.sleep(delay)
    
    print(f"Failed to find element after {max_attempts} attempts.")
    return False

def path_of_button(button):
    """ Return the file path of the button image based on the button name. """
    files_dir = paths['button_dataset_dir']
    
    expected_file_name = f"{button}.png"  # Change the extension if necessary
    
    for file_name in os.listdir(files_dir):
        if file_name.lower() == expected_file_name.lower():  # Case insensitive match
            return os.path.join(files_dir, file_name)
    
    print(f"Button image for '{button}' not found.")
    return None

def CLICK_ELEMENT(button, click_type='single'):
    """ Click the specified button using single or double click. """
    template_path = path_of_button(button)
    print(template_path)
    if template_path is None:
        return  # Exit if the template path is invalid

    attempts = 0  # Keep track of total attempts
    max_attempts = 10  # Maximum attempts allowed before failure

    while attempts < max_attempts:
        if find_and_click_element(template_path, click_type):
            print(f"Successfully clicked on '{button}' button. Exiting loop.")
            break  # Exit the loop after a successful click

        attempts += 1  # Increment the attempt count
        print(f"Attempt {attempts}/{max_attempts} failed to find '{button}'.")

        if attempts >= 3:  # After 3 consecutive failures, log the message
            send_message_to_csv(f"Failed to find element '{template_path}' after {attempts} consecutive attempts.")
            print(f"Finished: Failed to find element '{template_path}' after {attempts} attempts.")
            break  # Exit the loop to prevent infinite attempts
