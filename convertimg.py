from PIL import Image

# Load the PNG file
logo = Image.open("C:\\Users\\j84319062\\Documents\\GitHub\\Code_SAP_Tool\\img\\x.icon - Copy.png")

# Convert the PNG file to ICO and save it
logo.save("C:\\Users\\j84319062\\Documents\\GitHub\\Code_SAP_Tool\\x.icon - Copy.ico", format='ICO')
