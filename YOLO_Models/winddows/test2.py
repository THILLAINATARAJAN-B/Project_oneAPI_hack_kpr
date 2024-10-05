import cv2
import numpy as np
import mss
from ultralytics import YOLO
import tkinter as tk
from PIL import Image, ImageTk
import win32con
import win32gui
import threading
import time

# Load the YOLO model
model = YOLO(r"D:\CODES\PROJECTS\PC ASSISTANT\YOLO\YOLOv9\winddows\runs-windows_dataset\detect\train\weights\best.pt")

# Define your screen resolution (update these values as needed)
screen_width = 1920  # Example: Full HD width
screen_height = 1080  # Example: Full HD height

# Define screen capture dimensions
monitor = {"top": 0, "left": 0, "width": screen_width, "height": screen_height}

# Set up the tkinter window
root = tk.Tk()
root.overrideredirect(True)  # Remove window borders
root.geometry(f"{monitor['width']}x{monitor['height']}+{monitor['left']}+{monitor['top']}")  # Set size and position
root.attributes('-topmost', True)
root.attributes('-alpha', 0.5)  # Set transparency level (0.0 to 1.0)
root.configure(bg='black')  # Use black as the key color

# Function to set window as click-through
def set_click_through():
    hwnd = win32gui.FindWindow(None, "tk")  # The window title
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, 
                           win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | 
                           win32con.WS_EX_LAYERED | 
                           win32con.WS_EX_TRANSPARENT)

# Function to run the screen capture and detection in a separate thread
def capture_and_detect():
    with mss.mss() as sct:
        frame_count = 0
        while True:
            # Capture the screen
            img = sct.grab(monitor)

            # Convert the screen capture to a format OpenCV can work with
            frame = np.array(img)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)

            # Make predictions every 5 frames
            if frame_count % 5 == 0:
                results = model.predict(frame, imgsz=640, conf=0.2)  # Lower confidence threshold

                # Check if any detections are made
                if results and results[0].boxes:
                    # Draw bounding boxes on a separate image
                    annotated_frame = results[0].plot()
                else:
                    annotated_frame = frame  # No detections, show original frame

                # Convert to a format suitable for tkinter
                img_pil = Image.fromarray(cv2.cvtColor(annotated_frame, cv2.COLOR_BGR2RGB))
                img_tk = ImageTk.PhotoImage(img_pil)

                # Update the label with the new image
                label.img = img_tk  # Keep a reference to avoid garbage collection
                label.configure(image=img_tk)

            frame_count += 1
            time.sleep(0.01)  # Adjust to control the frame rate

# Start the click-through setting and the capture thread
set_click_through()
label = tk.Label(root, bg='black')
label.pack()

# Start the capture and detection thread
threading.Thread(target=capture_and_detect, daemon=True).start()

# Main loop for tkinter
root.mainloop()
