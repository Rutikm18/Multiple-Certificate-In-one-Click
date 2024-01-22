# Multiple-Certificate-In-one-Click

from PIL import Image, ImageDraw, ImageFont
import pandas as pd

def generate_image(name, number, template_path, output_path):
    # Load the template image
    template = Image.open(template_path)

    # Create a drawing object
    draw = ImageDraw.Draw(template)

    # Define font and size
    font_size = 40
    font = ImageFont.load_default()  # You may need to adjust this to a specific font and path on your system
    font = ImageFont.truetype("arial.ttf", font_size)  # Example using Arial font

    # Define the position to place the name and number
    name_position = (400, 305)
    number_position = (450, 440)

    # Define text color
    text_color = (0, 0, 0)

    # Draw the name and number on the image
    draw.text(name_position, f"{name}", font=font, fill=text_color)
    draw.text(number_position, f"{number}", font=font, fill=text_color)

    # Save the modified image
    template.save(output_path)

    # Uncomment the line below if you want to show the image
    template.show()

if __name__ == "__main__":
    # Replace 'D:\\1.xlsx' with the correct path to your Excel file
    excel_file_path = r'D:\\1.xlsx'

    # Read Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file_path) 

    # Replace these values with your desired template and output paths
    template_image_path = "D:/certificate.jpg"
    output_image_path_prefix = "D:/Cert_"  # Prefix for output image names

    # Iterate through rows in the DataFrame
    for index, row in df.iterrows():
        # Accessing data in each row
        name = row['Name']
        number = row['Number']

        # Generate output path for each row
        output_image_path = f"{output_image_path_prefix}{name}-{number}.jpg"
        
        # Generate the image for each row
        generate_image(name, number, template_image_path, output_image_path)
