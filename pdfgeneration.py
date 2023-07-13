from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import fonts

# Register the Times New Roman font
fonts.addMapping('Times-Roman', 0, 0, 'Times-Roman')
fonts.addMapping('Arial', 0, 0, 'Arial')

# Create a PDF canvas with letter page size
pdf = canvas.Canvas("test.pdf", pagesize=letter)

# Add image with a 1-inch margin from the top and left
pdf.drawImage("UCCS Logo Signature.png", x=inch, y=9 * inch, width=4 * 72, height=0.5 * inch)

# Set font properties
pdf.setFont("Times-Roman", 12)

name = "John Doe"  # Replace with actual name
address = "123 Test St, Colorado Springs, CO 80907"  # Replace with actual address

# Add congratulatory message
message_x = 1 * inch
message_y = letter[1] - 2.4 * inch  # Adjusted position for the paragraph
message_text = """July 13, 2023,<br/><br/><br/>
John Doe<br/>
123 Test St<br/>
Colorado Springs, CO 80907<br/><br/><br/>
Dear John,<br/><br/>
Congratulations! I am most pleased to announce that you have been named to the Dean's List for the Spring 2023 semester.
In order to receive this honor, a student must have earned between a 3.75-3.99 GPA, while completing at least 12 credit hours.
Your GPA from Spring 2023 was 3.93.<br/><br/>
Your outstanding academic performance for this semester is a source of considerable pride for the College of Letters, Arts, and Sciences at the University of Colorado Colorado Springs.<br/><br/>
Please accept my sincere congratulations for this well-deserved honor and my hope for your continued academic success. Keep up the great work.<br/><br/><br/>
Yours Truly,"""

signature_text = """K. Alex Ilyasova<br/>
Associate Dean<br/>
College of Letters, Arts, and Sciences"""

styles = getSampleStyleSheet()
paragraph_style = styles["Normal"]
paragraph_style.fontName = "Times-Roman"  # Set font to Times New Roman
paragraph_style.fontSize = 12  # Set font size to 12
paragraph_style.spaceBefore = 12
paragraph_style.leading = 14.4  # Set line spacing to 1.15
paragraph_style.spaceAfter = 13.86  # Set paragraph spacing after

# Footer text
footer_text = """College of Letters, Arts & Sciences • Office of the Dean<br/>
Columbine Hall 2025 • 1420 Austin Bluffs Pkwy • Colorado Springs, CO 80918<br/> 
t 719-255-4550 • f 719-255-4200"""

# Calculate the text area width and height
text_width = letter[0] - 2 * inch
text_height = 4 * inch

# Create a Paragraph object for the message text
paragraph = Paragraph(message_text, style=paragraph_style)

# Calculate the height required for the wrapped text
wrapped_text_height = paragraph.wrap(text_width, text_height)[1]

# Calculate the starting y-coordinate for the wrapped text
start_y = message_y - wrapped_text_height

# Draw the wrapped text within the text area
paragraph.drawOn(pdf, message_x, start_y)

# insert signature
signature_image_y = 3.33 * inch  # Adjusted position for the signature image
signature_text_y = signature_image_y - 0.01 * inch  # Adjusted position for the signature text

# Draw the signature image
pdf.drawImage("deanslistlettersignature.jpg", x=inch, y=signature_image_y, width=2.07 * 72, height=0.82 * inch)

# Create a Paragraph object for the signature text
signature_paragraph = Paragraph(signature_text, style=paragraph_style)

# Calculate the height required for the signature text
signature_text_height = signature_paragraph.wrap(text_width, text_height)[1]

# Calculate the starting y-coordinate for the signature text
signature_start_y = signature_text_y - signature_text_height

# Draw the signature text
signature_paragraph.drawOn(pdf, message_x, signature_start_y)

# Create a Paragraph style for the footer text
paragraph_style.fontSize = 9.96
paragraph_style.alignment = 1  # Center-justified

# Create a Paragraph object for the footer text
footer_paragraph = Paragraph(footer_text, paragraph_style)

# Calculate the height required for the footer text
footer_text_height = footer_paragraph.wrap(text_width, text_height)[1]

# Calculate the starting y-coordinate for the footer text
footer_text_y = 0.75 * inch

# Draw the centered, center-justified footer text
footer_paragraph.drawOn(pdf, 0.5 * (letter[0] - text_width), footer_text_y)

# Save the PDF
pdf.save()
