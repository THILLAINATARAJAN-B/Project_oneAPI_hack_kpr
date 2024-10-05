import numpy as np
import cv2
from openvino.runtime import Core

# Load the model
core = Core()
model_path = "D:\\CODES\\PROJECTS\\Assist\\YOLOv9\\runs-20240924T163833Z-001\\runs\\detect\\train\\weights\\best.onnx"
model = core.read_model(model_path)
compiled_model = core.compile_model(model, "CPU")

# Load an image
image_path = "D:\\CODES\\PROJECTS\\Assist\\YOLOv9\\Screenshot (29).png"
image = cv2.imread(image_path)
input_image = cv2.resize(image, (640, 640))
input_image = np.expand_dims(input_image, axis=0)

# Run inference
output = compiled_model(input_image)

# Print the output details
print("Model Output:")
for i, out in enumerate(output):
    print(f"Output {i}: shape={out.shape}, dtype={out.dtype}")

# Process output based on actual structure
if len(output) > 0:
    boxes = output[0]  # Assuming this is where boxes are
    scores = output[1] if len(output) > 1 else None  # Only access if it exists
    class_ids = output[2] if len(output) > 2 else None  # Only access if it exists

    # Visualize the detections (if boxes are present)
    if boxes is not None:
        for i in range(boxes.shape[1]):  # Assuming boxes shape is [1, N, 4]
            box = boxes[0, i]
            if scores is not None and scores[0, i] > 0.3:  # Confidence threshold
                class_id = class_ids[0, i] if class_ids is not None else None
                # Draw boxes and labels on the image as needed
                cv2.rectangle(image, (int(box[0]), int(box[1])), (int(box[2]), int(box[3])), (255, 0, 0), 2)
                if class_id is not None:
                    cv2.putText(image, f'Class: {class_id}', (int(box[0]), int(box[1]-10)), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 2)

# Display the result image
cv2.imshow('Detection', image)
cv2.waitKey(0)
cv2.destroyAllWindows()
