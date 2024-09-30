import pandas as pd
from PIL import Image, ImageDraw, ImageFont

def extract_names_from_excel(excel_path):
    # Read the Excel file and extract names from the "Full Name" column
    df = pd.read_excel(excel_path)
    names = df['Full Name'].tolist()
    return names

def create_certificate(name, template_path, output_path, font_path, name_position, font_size):
    # Load the template
    template = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(template)

    # Load the font
    font = ImageFont.truetype(font_path, font_size)

    # Calculate text size to center it
    text_bbox = draw.textbbox((0, 0), name, font=font)  # Get bounding box of the text
    text_width = text_bbox[2] - text_bbox[0]  # width = right - left
    center_x = ((template.width + 25) - text_width) / 2  # Calculate the x position to center the text
    center_y = name_position[1]  # Keep the y position as defined in the name_position

    # Draw the name on the certificate
    draw.text((center_x, center_y), name, fill="black", font=font)

    # Save the certificate with the name
    output_file = f"{output_path}/{name.replace(' ', '_')}_certificate.png"
    template.save(output_file)

def generate_certificates(excel_path, template_path, output_path, font_path, name_position, font_size):
    # Extract names from the Excel file
    names = extract_names_from_excel(excel_path)

    # Generate a certificate for each name
    for name in names:
        create_certificate(name, template_path, output_path, font_path, name_position, font_size)

if __name__ == "__main__":
    # Configuration
    excel_path = r"C:\Users\apurb\LegendsOfTheDecade\Legends of the Decade - Final Selected Volunteers (Responses).xlsx"  # Update with your Excel file path
    template_path = r"C:\Users\apurb\LegendsOfTheDecade\template.png"
    output_path = r"C:\Users\apurb\LegendsOfTheDecade\output"  # Folder where certificates will be saved
    font_path = r"C:\\Windows\\Fonts\\Georgia.ttf"  # Ensure this path is correct
    name_position = (0, 1570)  # X is now set to 0 for centering; adjust Y as needed
    font_size = 310  # Font size

    # Generate certificates
    generate_certificates(excel_path, template_path, output_path, font_path, name_position, font_size)
