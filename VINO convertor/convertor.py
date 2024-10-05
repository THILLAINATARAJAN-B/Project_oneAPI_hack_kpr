import cv2
import numpy as np
from openvino.runtime import Core

# Load OpenVINO Runtime
core = Core()

# Load the IR model (.xml and .bin files)
model_path = r"YOLO_Models\winddows\runs-windows_dataset\detect\train\weights\best.onnx"  # Updated to the OpenVINO model path
model = core.read_model(model=model_path)
compiled_model = core.compile_model(model=model, device_name="CPU")

# Load an image
image_path = r"D:\CODES\PROJECTS\PC ASSISTANT\APP\MODEL 5\AI-PC-Assist\test\img1.png"  # Verify the model accuracy
image = cv2.imread(image_path)
image_resized = cv2.resize(image, (640, 640))  # Resize to 640x640
input_image = image_resized.transpose(2, 0, 1)  # Change from HWC to CHW format
input_image = np.expand_dims(input_image, axis=0)  # Add batch dimension
input_image = input_image.astype(np.float32)  # Convert to float32

# Run inference
input_layer = compiled_model.input(0)  # Get input layer
output_layer = compiled_model.output(0)  # Get output layer
result = compiled_model([input_image])[output_layer]  # Run the model

# Process the result
print(result)
