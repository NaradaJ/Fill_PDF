import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import cv2
import os
import numpy as np
from tempfile import NamedTemporaryFile



# Specify the path to your Excel file
excel_file_path = r"C:\Users\narad\Downloads\1Q april to June 23 sample.xlsm"

# Read the Excel file into a Pandas DataFrame, skipping the first row (header row)
df = pd.read_excel(excel_file_path, header=0)

# Remove rows with NaN values
df = df.dropna()

# Reset the index of the DataFrame
df.reset_index(drop=True, inplace=True)

# Remove special characters from column headers
df.columns = df.columns.str.replace(r'[^\w\s]', '', regex=True)

# Identify merged cells (empty cells in this example)
merged_cells = df.applymap(lambda x: x == '')

# Clean up by replacing original merged cells with NaN
df = df.mask(merged_cells)

# Reset the index
df.reset_index(drop=True, inplace=True)

# Specify the indices of the columns to keep
columns_to_keep = [0, 1, 2, 6, 7, 11, 12, 13]

# Select the desired columns and exclude merged columns
desired_frame = df.iloc[:, columns_to_keep]

# Convert the DataFrame into a list of dictionaries
list_of_dicts = [row.to_dict() for _, row in desired_frame.iterrows()]

# Extract rows as separate datasets without headers
row_datasets = [row for row in desired_frame.values]

# Specify the path to the PDF file
pdf_file_path = r'C:\Users\narad\Downloads\APIT_T10_2324_EST.pdf'

# Open the PDF and select the first page
pdf_document = fitz.open(pdf_file_path)
page = pdf_document.load_page(0)

# Convert the page to an image
pix = page.get_pixmap()
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

# Load the image using OpenCV
image = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

# Define the order of text boxes and the corresponding row values
textbox_order = ["Employee Number", "Full Name", "NIC", "Total Gross Remuneration", "Cash Benefits",
                 "Deducted Tax1", "Deducted Tax2", "Remitted to the Inland Revenue Department"]

# Create a directory to store output images
output_directory = r"C:\Users\narad\Downloads\output_images"
os.makedirs(output_directory, exist_ok=True)

# Create a list to store the image paths
image_paths = []

# Define a function to save an image with text boxes
def save_image_with_text_boxes(image, row_data, output_directory, index):
    # Create a copy of the image to draw text boxes with the same resolution
    image_with_text_boxes = image.copy()

    # Define coordinates for the text boxes
    textbox_coordinates = [
        ((225, 282), (341, 299)),
        ((191, 236), (523, 254)),
        ((419, 283), (540, 299)),
        ((398, 362), (557, 377)),
        ((201, 404), (563, 419)),
        ((366, 518), (556, 530)),
        ((421, 530), (561, 543)),
        ((377, 570), (562, 586))
    ]

    # Define the order of text boxes and fill with corresponding row values
    for label, ((x1, y1), (x2, y2)), value in zip(textbox_order, textbox_coordinates, row_data):
        # Subtract 2 pixels from y1 and y2 to move the box 2 pixels up
        y1 -= 5
        y2 -= 5

        # Convert the value to a string
        if isinstance(value, float):
            value = f'{value:.2f}'
        else:
            value = str(value)

        # Add the value text to the copied image
        cv2.putText(image_with_text_boxes, value, (x1, y1 + 18), cv2.FONT_HERSHEY_SIMPLEX, 0.3, (0, 0, 0), 1)

    # Define a unique filename for each image
    image_filename = f"image_{index}.png"
    image_path = os.path.join(output_directory, image_filename)

    # Save the image with text boxes
    cv2.imwrite(image_path, image_with_text_boxes)

    return image_path

# Define a separate font size for "Deducted Tax2" value
deducted_tax2_font_size = 0.2  # Adjust the font size as needed

# Iterate through each row dataset and save each image
for index, row_data in enumerate(row_datasets):
    image_path = save_image_with_text_boxes(image, row_data, output_directory, index)
    image_paths.append(image_path)


import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import cv2
import os
import numpy as np
from tempfile import NamedTemporaryFile


# Create a PDF document
pdf_document = fitz.open()

# Iterate through the image paths and convert each image to a temporary PDF
for image_path in image_paths:
    with NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf_file:
        temp_pdf_path = temp_pdf_file.name
        img = Image.open(image_path)
        img.save(temp_pdf_path, "PDF")

    # Open the temporary PDF and insert it into the final PDF document
    temp_pdf_document = fitz.open(temp_pdf_path)
    pdf_document.insert_pdf(temp_pdf_document)
    temp_pdf_document.close()

    # Clean up: Remove the temporary PDF file
    os.remove(temp_pdf_path)

# Define the PDF output path
pdf_output_path = r"C:\Users\narad\Downloads\output.pdf"

# Save the PDF document with standard resolution
pdf_document.save(pdf_output_path)

# Close the PDF document
pdf_document.close()

# Clean up: Remove the temporary image files
for image_path in image_paths:
    os.remove(image_path)

print("PDF created successfully.")
