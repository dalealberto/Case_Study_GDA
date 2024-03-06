# I already installed package 'docx' but if you haven't yet, please do so
import os  # To change directory
from docx import Document  # To read .docx files
import sqlite3  # To create .sqlite file
import re  # To use regex

def glossary_sorted_by_course_week(glossary_dict):
    # Sort dictionary items by course week
    sorted_items = sorted(glossary_dict.items(), key = lambda x: x[1][1])

    # Create empty string variable
    glossary_string = ''

    # Iterate over the sorted items and print them
    for phrase, (definition, course_week) in sorted_items:
        glossary_string += f'{phrase}: {definition} {course_week}\n'

    return f'{glossary_string}\nGlossary Items: {len(glossary_dict)}'

# def glossary_sorted_by_phrase(glossary_dict):
#     # Sort dictionary items by phrase
#     sorted_items = sorted(glossary_dict.items())
#
#     # Create empty string variable
#     glossary_string = ''
#
#     # Iterate over the sorted items and print them
#     for phrase, (definition, course_week) in sorted_items:
#         glossary_string += f'{phrase}: {definition} {course_week}\n'
#
#     return f'{glossary_string}\nGlossary Items: {len(glossary_dict)}'

# Go to directory where glossary files are located
directory = r'C:\Users\dalea\Documents\Coursera\Google Data Analytics\Glossary'

# Create a dictionary to store the glossary items
glossary_dict = {}

# Iterate through each file in directory
for filename in os.listdir(directory):
    # Check if the file is a .docx file
    if filename.endswith('.docx'):
        # Split filename with underscore as delimiter (variable to be used for extracting file name later)
        file_split = filename.split('_')
        # Create the full path to the .docx file
        filepath = os.path.join(directory, filename)
        # Open the .docx file
        doc = Document(filepath)

        # Iterate through each paragraph in the file
        for paragraph in doc.paragraphs:
            # Filter only for text containing ':' (delimiter to separate phrase from definition)
            if ':' in paragraph.text:
                phrase, definition = re.split(r':\s*', paragraph.text, maxsplit=1)  # Split phrase and definition
                phrase = phrase.strip()  # Remove leading and trailing white spaces from the phrase
                definition = definition.strip()  # Remove leading and trailing white spaces from the definition
                course_week = f'({file_split[1]} {file_split[2]} {file_split[3]} {file_split[4]})'
                key = phrase  # Use phrase as the key for the dictionary

                # Check if phrase already exists in the dictionary
                if key in glossary_dict:
                    # If phrase exists, update its definition while retaining the old course_week
                    old_definition, old_course_week = glossary_dict[key]
                    glossary_dict[key] = (definition, old_course_week)
                else:
                    # If phrase doesn't exist, add it to the dictionary with its definition and course_week
                    glossary_dict[key] = (definition, course_week)

# Call function to print the glossary directly (optional)
#glossary_sorted_by_course_week(glossary_dict)

# Create string variable to be used for writing to .txt file
glossary_sorted_by_course_week_string = glossary_sorted_by_course_week(glossary_dict)

# Check if properly printing
#print(glossary_sorted_by_course_week_string)

# Write to a text file
with open('GDA_Glossary_Sorted_By_Course_Week.txt', 'w') as file:
    file.write(glossary_sorted_by_course_week_string)

