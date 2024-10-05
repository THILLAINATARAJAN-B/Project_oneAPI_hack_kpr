from ultralytics import YOLO
model=YOLO(r"D:\CODES\PROJECTS\PC ASSISTANT\YOLO\YOLOv9\gmail\runs\detect\train3\weights\best.pt")

result=model.predict(r"D:\CODES\PROJECTS\PC ASSISTANT\YOLO\YOLOv9\gmail\Screenshot 2024-10-04 183848.png",imgsz=640,conf=0.7,save=True,show=True)
