from ultralytics import YOLO
model=YOLO(r"D:\CODES\PROJECTS\PC ASSISTANT\YOLO\YOLOv9\winddows\runs-windows_dataset\detect\train\weights\last.pt")

result=model.predict(r"D:\CODES\PROJECTS\Assist\YOLOv9\Screenshot (29).png",imgsz=640,conf=0.7,save=True,show=True)
