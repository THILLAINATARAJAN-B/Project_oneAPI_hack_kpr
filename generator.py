import json  
from groq import Groq

# Load configuration from JSON
with open('config.json') as config_file:
    config = json.load(config_file)
    api_key = config.get('api_key')


class content_generator:
    def __init__(self,api_key):
        self.client = Groq(api_key=api_key)

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
                stream=False,  # Set to False for complete response
                stop=None,
            )
            
            # Extract content from the response
            if completion.choices:
                # Assuming the content is in the first choice
                content = completion.choices[0].message.content if hasattr(completion.choices[0].message, 'content') else ''
                return content.strip()
            else:
                print("No choices found in response.")
                return ''
        except Exception as e:
            print(f'Failed to generate content: {e}')
            return ''