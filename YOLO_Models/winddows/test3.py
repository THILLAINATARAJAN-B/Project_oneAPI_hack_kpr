import pyautogui
import time
from ultralytics import YOLO
from PIL import Image
import numpy as np
import cv2
import mss

# Load YOLO model
model = YOLO(r"D:\CODES\PROJECTS\PC ASSISTANT\APP\MODEL 5\AI-PC-Assist\YOLO_Models\winddows\runs-windows_dataset\detect\train\weights\best.onnx")

def capture_screen():
    """ Capture the current screen and return it as an image. """
    with mss.mss() as sct:
        # Get monitor screen dimensions
        monitor = sct.monitors[1]  # Take the first monitor
        screen_img = sct.grab(monitor)
        img = Image.frombytes("RGB", (screen_img.width, screen_img.height), screen_img.rgb)
        return img

def detect_recycle_bin(screen_img):
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
            if label == "Recycle Bin":
                # Return bounding box coordinates
                return detection.xyxy.cpu().numpy()[0]
    
    return None

import pyautogui
import time

def click_recycle_bin(coords):
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

        # Detect the recycle bin on the screen
        recycle_bin_coords = detect_recycle_bin(screen_img)
        
        if recycle_bin_coords is not None:
            print("Recycle Bin detected, opening it...")
            click_recycle_bin(recycle_bin_coords)
            break  # Stop once the Recycle Bin has been clicked
        
        # Wait before capturing the screen again
        time.sleep(2)

if __name__ == "__main__":
    # Simulate user input to open Recycle Bin
    user_input ="open recycle bin"
    
    if user_input.lower() == "open recycle bin":
        main()
