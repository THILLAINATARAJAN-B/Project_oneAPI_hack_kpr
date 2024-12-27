# Intelligent Remote PC Task Automation System

This project aims to develop an intelligent system capable of executing tasks on a PC remotely through natural language commands. These commands are issued via a mobile or web app, leveraging cutting-edge **Natural Language Processing (NLP)** and **Computer Vision (CV)** techniques. The system can detect task intents and perform actions accordingly, offering a seamless experience for the user.

The system utilizes **Intel DeepNN** for NLP tasks and the **YOLO model** (You Only Look Once) for computer vision tasks. Additionally, the project includes functionality for file transfer between the mobile device and the PC when both are connected.

## Features

- **Remote Task Execution:** Perform tasks on your PC by sending natural language commands through a mobile or web app.
- **Advanced NLP Integration:** Uses **Intel DeepNN** for understanding and processing user commands.
- **Computer Vision (CV) Capabilities:** **YOLO** model is used for vision-related tasks such as object detection and image analysis.
- **File Transfer Support:** Seamlessly transfer files between mobile and PC over a secure connection.
- **Mobile/Web Interface:** User-friendly interface for sending commands and interacting with the system.

## Technologies Used

- **Natural Language Processing (NLP):** 
  - **Intel DeepNN** model for detecting task intents and automating execution.
- **Computer Vision (CV):** 
  - **YOLO** model for vision-related tasks, such as object detection and image classification.
- **App Development:** 
  - Developed both mobile and web app interfaces to interact with the system.
- **Automation Techniques:** 
  - Automated task execution based on user commands via the app.

## System Workflow

1. **User Interaction:** The user sends a command through the mobile or web app.
2. **Command Processing:** The system uses **Intel DeepNN** to interpret the command and identify the task intent.
3. **Task Execution:** Once the intent is recognized, the system performs the task on the connected PC.
4. **File Transfer:** If needed, files are transferred between the mobile device and the PC.
5. **Vision Tasks (Optional):** When required, the YOLO model processes images for object detection and vision-based tasks.

## Installation

To get started with the system, follow these steps:

### 1. Clone the Repository

Clone the repository to your local machine:

```bash
git clone https://github.com/THILLAINATARAJAN-B/Project_oneAPI_hack_kpr.git
cd Project_oneAPI_hack_kpr
```

### 2. Set up Dependencies

Install the required dependencies for NLP, CV, and app functionality:

```bash
pip install -r requirements.txt
```

### 3. Install Intel DeepNN and YOLO

Make sure you have **Intel DeepNN** and **YOLO** installed. Follow the official documentation for installation:

- [Intel DeepNN](https://software.intel.com/content/www/us/en/develop/tools/deep-learning-deployment-sdk.html)
- [YOLO](https://pjreddie.com/darknet/yolo/) - Install YOLO and its dependencies.

### 4. Configure the App

Ensure that the mobile/web app is set up to communicate with the system. Follow the setup guide in the `app_setup.md` file for the correct configuration.

### 5. Run the System

Once everything is set up, start the system by running:

```bash
python main.py
```

Once the system is running, you can interact with it via the mobile or web app to execute commands and manage file transfers.

## Usage

After setting up the system, you can use it as follows:

- **Send Commands:** Open the mobile or web app and input commands such as "Open Chrome," "Start Music," "Play Video," or "Transfer Files."
- **File Transfers:** Choose the files to be transferred from the mobile/web app and send them to the connected PC. The system will handle the transfer automatically.
- **Computer Vision Tasks:** When needed, the system uses the YOLO model to detect objects in images or video streams and perform associated tasks.

The system processes the commands, detects the intent, and performs the task on the PC.

## Contribution

If you'd like to contribute to the project, feel free to fork the repository, make improvements, and submit a pull request. Contributions are always welcome!

Steps to contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new pull request.

---

## Acknowledgments

- **Intel:** For providing the **DeepNN** model for NLP tasks.
- **YOLO:** For enabling real-time object detection and image analysis.
- **Hackathon Organizers:** For supporting the development of innovative solutions.
- **Mentors:** For guiding and assisting throughout the project.


### Key Changes:
1. **YOLO Model**: The README now reflects that the system uses **YOLO** for computer vision tasks instead of **Intel OpenVINO**.
2. **Intel DeepNN**: Continued use of **Intel DeepNN** for NLP tasks.
3. **File Transfer and App Setup**: The instructions remain the same for file transfer and app communication setup.

Feel free to copy and paste this updated content into your `README.md` file on GitHub! Let me know if any further adjustments are needed.
