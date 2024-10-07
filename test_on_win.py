import os
import cv2
import face_recognition
import numpy as np
import pandas as pd
from openpyxl import Workbook
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Configuration
KNOWN_FACES_DIR = "J:\projects\pratibha_project\images"
CAPTURED_IMAGE = "captured_image.jpg"
ATTENDANCE_FILE = r"J:\projects\pratibha_project\attendance\attendance.xlsx"
FOLDER_ID = "16r1DeF6bFDkrrZGth2bp-Qtk5yLLlBDX"  # Replace with your Google Drive folder ID
SCOPES = ['https://www.googleapis.com/auth/drive.file']
# Initialize webcam
video_capture = cv2.VideoCapture(0)

# Load known faces
known_faces = []
known_names = []

for filename in os.listdir(KNOWN_FACES_DIR):
    image = face_recognition.load_image_file(f"{KNOWN_FACES_DIR}/{filename}")
    encoding = face_recognition.face_encodings(image)[0]
    known_faces.append(encoding)
    known_names.append(os.path.splitext(filename)[0])

# Capture an image from the webcam
ret, frame = video_capture.read()
if ret:
    cv2.imwrite(CAPTURED_IMAGE, frame)

# Release the webcam
video_capture.release()

# Load captured image and find faces
image = face_recognition.load_image_file(CAPTURED_IMAGE)
face_locations = face_recognition.face_locations(image)
face_encodings = face_recognition.face_encodings(image, face_locations)

# Initialize attendance list
attendance = {name: "Absent" for name in known_names}

# Mark attendance
for face_encoding in face_encodings:
    matches = face_recognition.compare_faces(known_faces, face_encoding)
    face_distances = face_recognition.face_distance(known_faces, face_encoding)
    best_match_index = np.argmin(face_distances)
    if matches[best_match_index]:
        name = known_names[best_match_index]
        attendance[name] = "Present"

# Generate Excel file
df = pd.DataFrame(list(attendance.items()), columns=["Name", "Attendance"])
df.to_excel(ATTENDANCE_FILE, index=False)

# Cleanup
if os.path.exists(CAPTURED_IMAGE):
    os.remove(CAPTURED_IMAGE)

# Google Drive upload
def upload_to_drive(file_name, folder_id):
    credentials = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    service = build("drive", "v3", credentials=credentials)
    
    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }
    media = MediaFileUpload(file_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().create(body=file_metadata, media_body=media, fields="id").execute()
# Upload attendance file to Google Drive
upload_to_drive(ATTENDANCE_FILE, FOLDER_ID)
print("Attendance Uploaded to Drive")