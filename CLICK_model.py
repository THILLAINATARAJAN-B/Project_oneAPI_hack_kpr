import pyautogui
import time
from ultralytics import YOLO
from PIL import Image
import numpy as np
import cv2
import mss

# Load YOLO model
model = YOLO(r"d:\CODES\PROJECTS\PC ASSISTANT\YOLO\YOLOv9\winddows\runs-windows_dataset\detect\train\weights\best.pt")

def capture_screen():
    """ Capture the current screen and return it as an image. """
    with mss.mss() as sct:
        monitor = sct.monitors[1]  # Take the first monitor
        screen_img = sct.grab(monitor)
        img = Image.frombytes("RGB", (screen_img.width, screen_img.height), screen_img.rgb)
        return img

def detect_element(screen_img, element):
    """ Use the YOLO model to detect the specified element in the screenshot. """
    opencv_img = np.array(screen_img)
    opencv_img = cv2.cvtColor(opencv_img, cv2.COLOR_RGB2BGR)
    
    results = model.predict(opencv_img, imgsz=640, conf=0.3)
    
    for result in results:
        for detection in result.boxes:
            label = result.names[int(detection.cls)]
            if label == element:
                return detection.xyxy.cpu().numpy()[0]
    
    return None

def click_element(coords):
    """ Simulate a mouse double-click. """
    x1, y1, x2, y2 = coords
    center_x = int((x1 + x2) / 2)
    center_y = int((y1 + y2) / 2)

    pyautogui.moveTo(center_x, center_y)
    time.sleep(0.5)
    pyautogui.doubleClick()
    print(f"Double-clicked at: {center_x}, {center_y}")

def find_and_click_element(element, max_attempts=10, delay=2):
    """ Find and click the specified element. """
    for _ in range(max_attempts):
        screen_img = capture_screen()
        print("screen detected")
        detected_coords = detect_element(screen_img, element)
        
        if detected_coords is not None:
            print(f"{element} detected, clicking it...")
            click_element(detected_coords)
            return True
        
        time.sleep(delay)
    
    print(f"Failed to find {element} after {max_attempts} attempts.")
    return False

def main():
    while True:
        user_input = input("Enter the element to find and click (or 'quit' to exit): ")
        if user_input.lower() == 'quit':
            break
        find_and_click_element(user_input)

if __name__ == "__main__":
    main()