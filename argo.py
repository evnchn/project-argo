import os
import glob
import collections
import collections.abc
from pptx import Presentation

def search_pptx_files(word):
    current_dir = os.getcwd()
    pptx_files = glob.glob(os.path.join(current_dir, "*.pptx"))

    word = word.lower()  # Convert the word to lowercase for case-insensitive matching

    for file_path in pptx_files:
        presentation = Presentation(file_path)
        total_pages = len(presentation.slides)

        found_pages = []  # List to store the page numbers where the word is found

        for page_number, slide in enumerate(presentation.slides, start=1):
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    if word in shape.text.lower():
                        found_pages.append(page_number)

        if found_pages:
            print(f"Word '{word}' found in '{os.path.basename(file_path)}', on page(s): {','.join(map(str, found_pages))}")
        else:
            print(f"Word '{word}' not found in '{os.path.basename(file_path)}'.")

search_word = input("Enter the word to search: ")
search_pptx_files(search_word)