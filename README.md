# Data Processing and PDF Generation Script

## Overview

This Python script reads data from an Excel file, processes and formats it, adds the data to a PDF template, and generates a final PDF document. It also allows for custom font sizes for specific values.

## Prerequisites

Before running the script, ensure you have the following installed:

- Python 3.x
- Required Python libraries (see `requirements.txt`)

You can install the required Python libraries using the following command:

```bash
pip install -r requirements.txt
Usage
Place your Excel file at the following location:

Input Excel file: data/input.xlsx
Place your PDF template at the following location:

PDF template: data/template.pdf
Modify the script as needed for custom data processing or font size adjustments (e.g., for "Deducted Tax2").

Run the script:

bash
Copy code
python your_script.py
The generated PDF will be saved as:
PDF output: output.pdf in the script directory
Custom Font Size
You can adjust the font size for specific values by modifying the deducted_tax2_font_size variable in the script.

Output Images
Intermediate PNG images with text boxes are saved in the following directory:

Output images directory: output_images/
Clean-Up
The script automatically cleans up temporary image files and temporary PDFs.

License
This project is licensed under the [License Name] License - see the LICENSE file for details.

Author
[Your Name]
Contact: [Your Email]
markdown
Copy code

In this template:

- **Overview**: Provide a brief overview of what the script does and its main purpose.

- **Prerequisites**: List the prerequisites or dependencies required to run the script, including Python version and any specific libraries.

- **Usage**: Explain how to use the script, including steps for setting up input files, running the script, and where to find the output.

- **Custom Font Size**: If there are any specific customization options in your script (like adjusting font sizes), explain how users can make those changes.

- **Output Images**: Mention any intermediate output files or directories created by the script.

- **Clean-Up**: Describe how the script handles cleaning up temporary files.

- **License**: Specify the project's license and provide a link to the full license details in a separate LICENSE file.

- **Author**: Narada Kasun  

Feel free to modify and expand the content to fit your specific project and provi