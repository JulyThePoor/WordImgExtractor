import os
from docx import Document
import zipfile

def extract_images_from_word(file_path, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    # Load the Word document
    doc = Document(file_path)

    # Extract images from the document
    for i, rel in enumerate(doc.part.rels.values()):
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            # Get the image format from the content type
            image_format = rel.target_part.content_type.split('/')[-1]
            image_name = f"image_{i}.{image_format}"
            image_path = os.path.join(output_folder, image_name)

            # Save the image to the output folder
            with open(image_path, "wb") as f:
                f.write(image_data)

    # Create a zip file and add the images to it
    zip_file_path = os.path.join(output_folder, "images.zip")
    with zipfile.ZipFile(zip_file_path, "w") as zip_file:
        for root, _, files in os.walk(output_folder):
            for file in files:
                file_path = os.path.join(root, file)
                zip_file.write(file_path, os.path.relpath(file_path, output_folder))

    print("Images extracted and saved in a zip file.")

# Usage example
word_file_path = r"C:\Users\Administrator\Desktop\vs\Doc1.docx"
output_folder_path = r"C:\Users\Administrator\Desktop\vs\img"
extract_images_from_word(word_file_path, output_folder_path)