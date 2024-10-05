import json  
from groq import Groq
import spacy
from typing import Dict, Any
from paths import config_path

# Load configuration from JSON
with open(config_path) as config_file:
    config = json.load(config_file)
    api_key = config.get('api_key')

class CommandParser:
    def __init__(self):
        self.nlp = spacy.load("en_core_web_sm")

    def parse_command(self, user_input: str) -> Dict[str, Any]:
        doc = self.nlp(user_input.lower())
        
        result = {
            "original_input": user_input,
            "action": self._extract_action(doc),
            "folder_name": self._extract_folder_name(doc),
            "file_name": self._extract_file_name(doc),
            "source_folder": self._extract_source_folder(doc),
            "destination_folder": self._extract_destination_folder(doc),
            "file_type": self._extract_file_type(doc),
        }
        
        return {k: v for k, v in result.items() if v}  # Remove empty values

    def _extract_action(self, doc):
        action_verbs = {"create": "Create", "open": "Open", "move": "Move", "install": "Install", "search": "Search"}
        for token in doc:
            if token.lemma_ in action_verbs:
                return action_verbs[token.lemma_]
        return None

    def _extract_folder_name(self, doc):
        folders = ["documents", "desktop", "downloads", "pictures", "music"]
        for token in doc:
            if token.text.lower() in folders:
                return token.text.capitalize()
        return None

    def _extract_file_name(self, doc):
        for token in doc:
            if token.like_num or (token.pos_ == "NOUN" and "." in token.text):
                return token.text
        return None

    def _extract_source_folder(self, doc):
        for i, token in enumerate(doc):
            if token.text.lower() == "from" and i + 1 < len(doc):
                return doc[i + 1].text.capitalize()
        return None

    def _extract_destination_folder(self, doc):
        for i, token in enumerate(doc):
            if token.text.lower() == "to" and i + 1 < len(doc):
                return doc[i + 1].text.capitalize()
        return None

    def _extract_file_type(self, doc):
        file_types = {"pdf": "PDF", "jpg": "JPG", "jpeg": "JPG", "png": "PNG", "txt": "TXT", "doc": "DOC", "docx": "DOCX"}
        for token in doc:
            if token.text.lower() in file_types:
                return file_types[token.text.lower()]
        return None

class ContentGenerator:
    def __init__(self, api_key):
        self.client = Groq(api_key=api_key)
        self.command_parser = CommandParser()

    def generate_content(self, prompt):
        try:
            completion = self.client.chat.completions.create(
                model="llama3-groq-70b-8192-tool-use-preview",
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.5,
                max_tokens=1024,
                top_p=0.65,
                stream=False,
                stop=None,
            )
            
            # Extract content from the response
            if completion.choices:
                content = completion.choices[0].message.content if hasattr(completion.choices[0].message, 'content') else ''
                return content.strip()
            else:
                print("No choices found in response.")
                return ''
        except Exception as e:
            print(f'Failed to generate content: {e}')
            return ''

    def process_command(self, user_input: str):
        parsed_result = self.command_parser.parse_command(user_input)
        prompt = f"Based on the command: {user_input}, can you provide detailed instructions?"
        generated_content = self.generate_content(prompt)
        return {
            "parsed_result": parsed_result,
            "generated_content": generated_content
        }

def main():
    generator = ContentGenerator(api_key)
    
    # Example command input
    user_command = "send mail to vikash@gmail.com about welcoming"
    result = generator.process_command(user_command)

    print("Parsed Result:")
    for key, value in result['parsed_result'].items():
        print(f"{key.replace('_', ' ').capitalize()}: {value}")

    print("\nGenerated Content:")
    print(result['generated_content'])

if __name__ == "__main__":
    main()
