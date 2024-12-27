Intelligent Remote PC Task Automation System
This project aims to develop an intelligent system capable of executing tasks on a PC remotely through natural language commands. These commands are issued via a mobile or web app, leveraging cutting-edge Natural Language Processing (NLP) and Computer Vision (CV) techniques. The system can detect task intents and perform actions accordingly, offering a seamless experience for the user.

The system utilizes Intel DeepNN for NLP tasks and Intel OpenVINO for computer vision tasks. Additionally, the project includes functionality for file transfer between the mobile device and the PC when both are connected.

Features
Remote Task Execution: Perform tasks on your PC by sending natural language commands through a mobile or web app.
Advanced NLP Integration: Uses Intel DeepNN for understanding and processing user commands.
Computer Vision (CV) Capabilities: Intel OpenVINO enables task execution related to vision-based tasks.
File Transfer Support: Seamlessly transfer files between mobile and PC over a secure connection.
Mobile/Web Interface: User-friendly interface for sending commands and interacting with the system.
Technologies Used
Natural Language Processing (NLP):
Intel DeepNN model for detecting task intents and automating execution.
Computer Vision (CV):
Intel OpenVINO model for vision-related tasks.
App Development:
Developed both mobile and web app interfaces to interact with the system.
Automation Techniques:
Automated task execution based on user commands via the app.
System Workflow
User Interaction: The user sends a command through the mobile or web app.
Command Processing: The system uses Intel DeepNN to interpret the command and identify the task intent.
Task Execution: Once the intent is recognized, the system performs the task on the connected PC.
File Transfer: If needed, files are transferred between the mobile device and the PC.
Installation
To get started with the system, follow these steps:

1. Clone the Repository
bash
Copy code
git clone https://github.com/THILLAINATARAJAN-B/Project_oneAPI_hack_kpr.git
cd Project_oneAPI_hack_kpr
2. Set up Dependencies
Install the necessary dependencies for NLP, CV, and app functionality.

bash
Copy code
pip install -r requirements.txt
3. Install Intel DeepNN and OpenVINO
Make sure you have Intel DeepNN and Intel OpenVINO installed. Follow the official documentation to install both tools:

Intel DeepNN
Intel OpenVINO
4. Configure the App
Set up the mobile/web app by following the instructions in the app_setup.md file. Ensure that the app is properly connected to the system for remote interaction.

5. Run the System
Start the system with the following command:

bash
Copy code
python main.py
Once the system is running, you can interact with it via the mobile or web app to execute commands and manage file transfers.

Usage
Send Commands: Open the mobile or web app and input commands such as "Open Chrome," "Start Music," or "Transfer Files."
File Transfers: Select files from the app to transfer between your mobile device and the connected PC.
The system processes the commands, detects the intent, and automates the task accordingly.

Contribution
Feel free to fork the repository, make changes, and submit a pull request to contribute to the project. Contributions are always welcome!

License
This project is licensed under the MIT License - see the LICENSE file for more details.
