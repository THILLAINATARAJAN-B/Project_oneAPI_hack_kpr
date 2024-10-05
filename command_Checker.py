from NLP.NLP_ORG import NATURAL

def format_for_task_performer(output):
    task_list = []  # Use a list to maintain the order of tasks
    
    for subtask in output['subtasks']:
        # Remove quotes and leading/trailing whitespace from the intent
        intent = subtask['predicted_intent'].strip().strip('"')
        
        # Extract values if present, otherwise use an empty string
        extracted_value = ''
        if 'extracted_values' in subtask and subtask['extracted_values']:
            # For simplicity, we're just taking the first value found
            first_key = next(iter(subtask['extracted_values']))
            first_value = subtask['extracted_values'][first_key]
            if isinstance(first_value, list):
                extracted_value = first_value[0] if first_value else ''
            else:
                extracted_value = str(first_value)
        
        # Add the intent and extracted value as a tuple to the list
        task_list.append((intent, extracted_value))
    
    return task_list

def print_all(output):
    print(output)
    print("\nProcessed Result:")
    print(f"Original Input: {output['original_input']}")
    print(f"Overall Intent: {output['overall_intent']}")
    print(f"Overall Similarity Score: {output['overall_similarity_score']:.4f}")
    
    if output['extracted_values']:
        print("\nExtracted Values:")
        for key, value in output['extracted_values'].items():
            print(f"  {key.title()}: {value}")
    
    print("\nSubtasks:")
    for i, subtask in enumerate(output['subtasks'], 1):
        print(f"\n  Subtask {i}:")
        print(f"    Text: {subtask['subtask']}")
        print(f"    Predicted Intent: {subtask['predicted_intent']}")
        print(f"    Similarity Score: {subtask['similarity_score']:.4f}")
        print(f"    Best Matching Command: {subtask['best_matching_command']}")
        
        if subtask['extracted_values']:
            print(f"    Extracted Values:")
            for key, value in subtask['extracted_values'].items():
                print(f"      {key.title()}: {value}")


def NLP(user_input):
    natural = NATURAL()  # Create the NATURAL instance outside the loop
    output = natural.run(user_input)
    print_all(output)
    task_list = format_for_task_performer(output)
    return task_list

# a=input("Enter your:")
# b=NLP(a)

# def NLP(user_input):
#     natural = NATURAL()  # Create the NATURAL instance outside the loop
#     output = natural.run(user_input)
#     print_all(output)
#     task_list = format_for_task_performer(output)
#     return task_list