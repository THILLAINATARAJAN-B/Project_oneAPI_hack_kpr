import cv2
import numpy as np
import mss
import os
from paths import paths

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

def path_of_button(button):
    """ Return the file path of the button image based on the button name. """
    files_dir = paths['button_dataset_dir']
    
    # Construct the expected file name
    expected_file_name = f"{button}.png"
    
    # Iterate through the files in the directory
    for file_name in os.listdir(files_dir):
        if file_name.lower() == expected_file_name.lower():  # Case insensitive match
            return os.path.join(files_dir, file_name)
    
    print(f"Button image for '{button}' not found.")
    return None

def verify_screen_status(templates):
    """ Verify the presence of multiple screen elements. """
    screen_img = capture_screen()
    detected_elements = {}
    
    for template in templates:
        template_path = path_of_button(template)
        if template_path is None:
            detected_elements[template] = False
            continue
        
        detected_coords = detect_element(screen_img, template_path)
        detected_elements[template] = detected_coords is not None
    
    return detected_elements

# Example usage
if __name__ == "__main__":
    results = verify_screen_status(['ms_search_serach_result', 'ms_search_app'])
    for element, is_found in results.items():
        if is_found:
            print(f"Element '{element}' detected.")
        else:
            print(f"Element '{element}' not found.")
