import spacy

# Load the spaCy model
nlp = spacy.load("en_core_web_sm")

applications = {
            'settings': 'ms-settings:',
            'whatsapp': 'C:\\Program Files\\WindowsApps\\Microsoft.WhatsApp_...',
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'word': 'winword.exe'
        }
def parse_command(command):
    
    """
    Parses the user command to determine the action to perform.
    """
    try:
        # Lowercase and process the command
        doc = nlp(command.lower())

        # Check for brightness command
        if 'brightness' in command:
            for ent in doc.ents:
                if ent.label_ == 'PERCENT':
                    try:
                        brightness = int(ent.text.replace('%', ''))
                        return 'set_brightness', brightness
                    except ValueError:
                        return 'unknown_command', "Invalid brightness value."

        command = command.lower()
        
        # Check for opening settings
        if 'open settings' in command:
            return 'open_settings', None
        
        if 'windows security' in command:
            return 'open_windows_security',None
        
        if 'app' in command:
            for app_name in applications:
                if app_name in command:
                    return 'open_application', app_name
        
        # Check for closing settings
        if 'close settings' in command:
            return 'close', 'settings'

        # Check for creating a folder
        if 'create folder' in command:
            try:
                folder_name = command.split('create folder', 1)[1].strip()
                return 'create_folder', folder_name
            except IndexError:
                return 'unknown_command', "Folder name not provided."

        # Check for renaming a folder
        if 'rename folder' in command:
            try:
                parts = command.split('rename folder', 1)[1].split('to', 1)
                if len(parts) == 2:
                    old_name = parts[0].strip()
                    new_name = parts[1].strip()
                    return 'rename_folder', (old_name, new_name)
                else:
                    return 'unknown_command', "Invalid rename format."
            except IndexError:
                return 'unknown_command', "Folder names not provided."

        # Check for deleting a file or folder
        if 'delete' in command:
            try:
                item_name = command.split('delete', 1)[1].strip()
                return 'delete_file_or_folder', item_name
            except IndexError:
                return 'unknown_command', "Item name not provided."

        # Check for opening a file
        
        if 'open' in command and ('my pc' in command or 'this pc' in command):
            return 'open', 'my pc'
        
        # Check for opening a specific drive
        if 'open' in command and any(drive in command for drive in ['c:', 'd:', 'e:', 'f:', 'a:', 'b:']):
            for drive in ['c:', 'd:', 'e:', 'f:', 'a:', 'b:']:
                if drive in command:
                    return 'open', drive.upper()
        
        # Check for opening a folder or file by name
        if 'open' in command:
            path = command.split('open', 1)[1].strip()
            return 'open', path

        # Check for searching a file
        if 'search file' in command:
            try:
                file_name = command.split('search file', 1)[1].strip()
                return 'search_file', file_name
            except IndexError:
                return 'unknown_command', "File name not provided."

        # Check for closing an application
        if 'close application' in command:
            try:
                app_name = command.split('close application', 1)[1].strip()
                return 'close_application', app_name
            except IndexError:
                return 'unknown_command', "Application name not provided."

        # Check for 'close' command
        if 'close' in command:
            if 'file' in command:
                try:
                    # Extract file path to close
                    file_path = command.split('close file', 1)[1].strip()
                    return 'close_file', file_path
                except IndexError:
                    return 'unknown_command', "File path not provided."
                    
            elif 'settings' in command:
                return 'close', 'settings'
            
            elif 'my pc' in command:
                return 'close', 'my pc'
            
            else:
                # Extract the target of the 'close' command
                try:
                    target = command.split('close', 1)[1].strip()
                    
                    # Check if the target is a disk drive
                    if target in ['c:', 'd:', 'e:', 'f:']:
                        return 'close', target.upper()
                    
                    # If not a disk drive, return the target as is
                    return 'close', target
                
                except IndexError:
                    # In case no specific target is provided
                    return 'close', None

        # Check for opening the task manager
        if 'task manager' in command:
            return 'open_task_manager', None

        # Check for taking a screenshot
        if 'take screenshot' in command:
            try:
                file_name = command.split('take screenshot', 1)[1].strip()
                return 'take_screenshot', file_name
            except IndexError:
                return 'unknown_command', "File name for screenshot not provided."

        # Check for search bar, refresh, maximize, and minimize commands
        # Check for search command
        if 'search' in command:
            try:
                # Extract the search query after 'search'
                search_query = command.split('search', 1)[1].strip()
                return 'search', search_query
            except IndexError:
                return 'unknown_command', "Search query not provided."
        if 'refresh' in command:
            return 'refresh', None
        if 'maximize' in command:
            return 'maximize', None
        if 'minimize' in command:
            return 'minimize', None
        if 'wordcontent' in command:
            return 'word_content','Ferrari'
        if 'create google account' in command:
            return 'create_google_account',None

    except Exception as e:
        return 'unknown_command', f"Error parsing command: {e}"

    # Return unknown command if no matches are found
    return 'unknown_command', "Command not recognized."
