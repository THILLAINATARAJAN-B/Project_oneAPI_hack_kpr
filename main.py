import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import os
import sys
import subprocess
import csv
import threading
import time
import logging
from command_executor import execute_command
from command_Checker import NLP
from paths import paths

# Global variable to track executed commands
executed_commands = set()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def restart_app():
    """Restart the application."""
    try:
        print("Restarting application...")  # Debugging output
        root.destroy()
        python = sys.executable
        script = os.path.abspath(__file__)
        subprocess.Popen([python, script] + sys.argv)
        sys.exit()
    except Exception as e:
        print(f"Failed to restart the application: {e}")
        messagebox.showerror("Error", f"Failed to restart the application: {e}")

def focus_input_entry():
    """Focus the input entry widget by clicking at its center."""
    widget_x = input_entry.winfo_rootx()
    widget_y = input_entry.winfo_rooty()
    widget_width = input_entry.winfo_width()
    widget_height = input_entry.winfo_height()

    center_x = widget_x + widget_width // 2
    center_y = widget_y + widget_height // 2
    pyautogui.click(center_x, center_y)

def process_input(event=None):
    """Process the user input."""
    user_input = input_entry.get()
    
    try:
        task_list = NLP(user_input)
        print("task to execute:",task_list)
        for command, value in task_list:
            result_label.config(text=f"Executing: {command} with value: {value}")
            root.update()  # Update the GUI to show the current command being executed
            execute_command(command, value, result_label, root, input_entry)
            
            # Log the executed command
            send_message_to_csv(f"Executed: {command} with value: {value}")
            
            time.sleep(1)  # Add a small delay between commands for better visibility
        
        result_label.config(text="All commands executed.")
    except Exception as e:
        error_message = f"Error: {e}"
        result_label.config(text=error_message)
        send_message_to_csv(error_message)
    
    # Clear the input field and refocus
    input_entry.delete(0, tk.END)
    root.after(100, lambda: input_entry.focus_force())
    root.after(100, focus_input_entry)

def disable_close(event=None):
    """Prevent the user from closing the window."""
    messagebox.showinfo("Information", "This window cannot be closed. Please use the input box.")
    return "break"

def read_commands_from_csv(file_path):
    """Read commands from the CSV file."""
    commands = []
    try:
        with open(file_path, mode='r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if len(row) < 3:
                    continue
                command_text, command_type, source = row
                logging.info(f"Read row: {row}")
                if command_type == 'message' and source == 'android' and command_text not in executed_commands:
                    commands.append(command_text)
                    logging.info(f"Added command: {command_text}")
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
    except Exception as e:
        logging.error(f"Failed to read commands from CSV: {e}")

    return commands

def execute_commands_from_csv(file_path):
    global executed_commands
    commands = read_commands_from_csv(file_path)
    
    if not commands:
        print("No new commands to execute.")  # Debugging output
        return
    
    # Read all rows to filter out executed commands later
    with open(file_path, mode='r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        all_rows = list(reader)

    updated_rows = []  # To store remaining rows after executionco

    for command in commands:
        print(f"Executing command: {command}")  # Debugging output
        
        # Check if the command has already been executed
        if command in executed_commands:
            print(f"Command already executed: {command}")
            continue  # Skip already executed commands

        try:
            process_command(command, result_label, root, input_entry)  # Pass additional arguments
            executed_commands.add(command)  # Mark command as executed
            
            # Log the executed command
            send_message_to_csv(f"Executed from CSV: {command}")
        except Exception as e:
            error_message = f"Error executing '{command}': {e}"
            print(error_message)
            send_message_to_csv(error_message)
        
        # Filter out the executed command
        for row in all_rows:
            if row and row[0] == command:  # Check if the first column matches the command
                continue  # Skip this row
            updated_rows.append(row)  # Keep this row

    # Write the updated rows back to the CSV
    with open(file_path, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(updated_rows)  # Write remaining commands back to CSV      

def send_message_to_csv(content):
    """Send a message to the send.csv file."""
    send_csv_file = paths['to_send_message_path']
    try:
        with open(send_csv_file, mode='a', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([content, 'message', 'application'])  # Log the message
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write to send.csv: {e}")

def process_command(command, result_label, root, input_entry):
    """Process a set of commands and execute them one by one."""
    task_list = NLP(command)
    
    print(task_list)
    # Convert to a list to allow modifications during iteration
    task_list = list(task_list)
    
    for parsed_command, value in task_list:
        result_label.config(text=f"Executing: {parsed_command} with value: {value}")
        root.update()  # Update the GUI to show the current command being executed
        print("Executing", parsed_command)
        execute_command(parsed_command, value, result_label, root, input_entry)
        
        time.sleep(1)  # Add a small delay between commands for better visibility
    
    result_label.config(text="All commands executed.")
    
    # After all commands are processed, mark the command as executed
    executed_commands.add(command)  # Mark command as executed

    # Remove the command from the task list after execution
    for cmd in task_list:

        pass  # Remove logic can be implemented here

def poll_csv_commands(file_path):
    """Continuously poll the CSV file for new commands."""
    while True:
        execute_commands_from_csv(file_path)
        time.sleep(5)  # Poll every 5 seconds

def create_sticky_window():
    global input_entry, result_label, root

    # Create the main window
    root = tk.Tk()
    
    # Set the window title
    root.title("Command Input")

    # Set the window size and remove padding
    window_width = 400
    window_height = 40
    root.geometry(f"{window_width}x{window_height}")

    # Make the window always on top
    root.attributes('-topmost', True)

    # Remove the title bar and window decorations to integrate with Windows
    root.overrideredirect(True)

    # Style configuration
    style = ttk.Style()
    style.configure('TEntry', padding='2', relief='flat', background='#F0F0F0', foreground='#000000', font=('Arial', 8))
    style.configure('TLabel', padding='0', background='#FFFFFF', foreground='#000000', font=('Arial', 8))
    style.configure('TButton', padding='2', relief='flat', background='#0078D4', foreground='#FFFFFF', font=('Arial', 8))

    # Create a frame for better layout management
    frame = ttk.Frame(root, padding="5")
    frame.pack(expand=True, fill=tk.BOTH)

    # Create a frame for side-by-side widgets
    button_entry_frame = ttk.Frame(frame)
    button_entry_frame.pack(fill=tk.X, expand=True)

    # Create an Entry widget for user input
    input_entry = ttk.Entry(button_entry_frame, width=30)
    input_entry.pack(side=tk.LEFT, padx=5, pady=1, fill=tk.X, expand=True)

    # Create a Restart button
    restart_button = ttk.Button(button_entry_frame, text="Restart", command=restart_app)
    restart_button.pack(side=tk.LEFT, padx=5, pady=1)

    # Create a label to display results or error messages
    result_label = ttk.Label(frame, text="", wraplength=window_width)
    result_label.pack(pady=5, padx=5)

    # Bind the Enter key to the process_input function
    root.bind('<Return>', process_input)
    input_entry.focus_force()
    
    # Bind the window close event to a function to prevent closing
    root.protocol("WM_DELETE_WINDOW", disable_close)

    # Position the window at the top of the screen
    screen_width = root.winfo_screenwidth()
    x_position = (screen_width - window_width) // 2
    y_position = 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Start a thread to poll for CSV commands
    threading.Thread(target=poll_csv_commands, args=(paths['task_path'],), daemon=True).start()

    # Start the Tkinter event loop
    root.mainloop()

# Run the application
if __name__ == "__main__":
    create_sticky_window()
