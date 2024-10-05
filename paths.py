# paths.py

import os

# Get the directory of the current script (assumed to be inside AI-PC-Assist)
base_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the paths relative to AI-PC-Assist
nlp_dir = os.path.join(base_dir, 'NLP')
command_csv_path = os.path.join(nlp_dir, 'command.csv')
model_dir = os.path.join(nlp_dir, 'Models')
tokenizer_path = os.path.join(model_dir, 'tokenizer.pkl')
intent_model_path = os.path.join(model_dir, 'intent_model.h5')
summarization_model_path = os.path.join(model_dir, 'summarization_model.h5')
dataset_dir = os.path.join(base_dir, 'dataset')
button_dataset_dir = os.path.join(dataset_dir, 'button')
task_path = os.path.join(base_dir, 'task.csv')
to_send_message_path = os.path.join(base_dir, 'to_send_message.csv')
config_path = os.path.join(base_dir, 'config.json')
backend_dir=os.path.join(base_dir,'backend')
file_storage_path = os.path.join(backend_dir, 'Files_From_Mobile')

# You can also define a dictionary for better organization
paths = {
    "command_csv_path": command_csv_path,
    "intent_model_path": intent_model_path,
    "tokenizer_path": tokenizer_path,
    "summarization_model_path": summarization_model_path,
    "task_path": task_path,
    "to_send_message_path": to_send_message_path,
    "config_path": config_path,
    "button_dataset_dir": button_dataset_dir,
    "file_storage_dir":file_storage_path
}
