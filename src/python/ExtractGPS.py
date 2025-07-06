#pip install geopy Pillow gmplot simplekml

import os
import PIL.Image
from PIL.ExifTags import TAGS
from geopy.geocoders import Nominatim
import gmplot
import simplekml
import pathlib
import shutil
from zipfile import ZipFile

# Define paths and title
base_path = r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\Desktop\1\New folder"
title = 'Visita VGR 06-2025.kmz'
kml = simplekml.Kml()
temp_dir = os.path.join(base_path, "temp_images")

# Create a temporary directory to store images for embedding
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

# Function to find files
def find_files(extension, search_path):
    result = []
    for root, dirs, files in os.walk(search_path):
        for file in files:
            if file.endswith(extension):
                result.append(
                    {
                        'InputFile': file,
                        'InputDir': root,
                        'OutputFile': file.split(".")[0] + ".pdf",
                        'OutputDir': pathlib.Path(base_path).as_uri(),
                    }
                )
    return result

# Helper function to convert GPS coordinates to degrees
def convert_to_degress(Hours, Minutes, Seconds, Direction):
    if Direction == "S" or Direction == "W":
        return -(float(Hours) + (float(Minutes) / 60.0) + (float(Seconds) / 3600.0))
    else:
        return +(float(Hours) + (float(Minutes) / 60.0) + (float(Seconds) / 3600.0))

# Process files and add points to KML with image references
Files = find_files(".jpg", base_path)

for file in Files:
    image_path = os.path.join(file['InputDir'], file['InputFile'])
    image = PIL.Image.open(image_path)
    metadata = image._getexif()
    
    for tag, value in metadata.items():
        if TAGS.get(tag) == "GPSInfo":
            gps_info = value
            latitude = convert_to_degress(gps_info[2][0], gps_info[2][1], gps_info[2][2], gps_info[1])
            longitude = convert_to_degress(gps_info[4][0], gps_info[4][1], gps_info[4][2], gps_info[3])
            
            # Copy the image to the temp directory
            temp_image_path = os.path.join(temp_dir, file['InputFile'])
            image.save(temp_image_path)
            
            # Add the KML point with image reference
            point = kml.newpoint(name=file['InputFile'], coords=[(longitude, latitude)])
            point.description = f'<img src="{file["InputFile"]}" alt="{file["InputFile"]}" width="400" height="300" align="left" />'
            
            print(f"{file['InputFile']},{latitude},{longitude}")
            break

# Save the KML file to the temp directory
kml_path = os.path.join(temp_dir, "doc.kml")
kml.save(kml_path)

# Create KMZ file by zipping the temp directory contents
kmz_path = os.path.join(base_path, title)
with ZipFile(kmz_path, 'w') as kmz:
    for root, _, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            kmz.write(file_path, os.path.relpath(file_path, temp_dir))

# Clean up temp directory after creating KMZ
shutil.rmtree(temp_dir)

print(f"KMZ file created at: {kmz_path}")
