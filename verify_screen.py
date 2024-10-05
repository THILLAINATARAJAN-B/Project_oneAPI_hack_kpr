import pyautogui
import time
from ultralytics import YOLO
from PIL import Image
import numpy as np
import cv2
import mss

# Load YOLO model
model = YOLO(r"D:\CODES\PROJECTS\Assist\YOLOv9\runs-20240924T163833Z-001\runs\detect\train\weights\best.pt")

def capture_screen():
    """ Capture the current screen and return it as an image. """
    with mss.mss() as sct:
        # Get monitor screen dimensions
        monitor = sct.monitors[1]  # Take the first monitor
        screen_img = sct.grab(monitor)
        img = Image.frombytes("RGB", (screen_img.width, screen_img.height), screen_img.rgb)
        return img

def detect_element(screen_img,element):
    """ Use the YOLO model to detect the recycle bin in the screenshot. """
    # Convert to OpenCV image (as YOLO expects OpenCV format)
    opencv_img = np.array(screen_img)
    opencv_img = cv2.cvtColor(opencv_img, cv2.COLOR_RGB2BGR)
    
    # Perform prediction
    results = model.predict(opencv_img, imgsz=640, conf=0.3)
    
    # Loop through results to find 'Recycle Bin'
    for result in results:
        for detection in result.boxes:
            label = result.names[int(detection.cls)]
            if label == element:
                # Return bounding box coordinates
                return detection.xyxy.cpu().numpy()[0]
    
    return None

def click_elemet(coords):
    """ Simulate a mouse double-click. """
    x1, y1, x2, y2 = coords
    center_x = int((x1 + x2) / 2)
    center_y = int((y1 + y2) / 2)

    # Move and perform double-click
    pyautogui.moveTo(center_x, center_y)
    time.sleep(0.5)
    pyautogui.doubleClick()  # Perform a double-click
    print(f"Double-clicked at: {center_x}, {center_y}")

def main():
    while True:
        # Capture the screen
        screen_img = capture_screen()

        element='Recycle Bin'
        # Detect the recycle bin on the screen
        detected_coords = detect_element(screen_img,element)
        
        if detected_coords is not None:
            print("Recycle Bin detected, opening it...")
            click_elemet(detected_coords)
            break  # Stop once the Recycle Bin has been clicked
        
        # Wait before capturing the screen again
        time.sleep(2)

if __name__ == "__main__":
    # Simulate user input to open Recycle Bin
    user_input ="open recycle bin"
    
    if user_input.lower() == "open recycle bin":
        main()
