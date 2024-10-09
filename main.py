import os
import win32com.client
import cv2
import shutil
import numpy as np
from cvzone.HandTrackingModule import HandDetector
from moviepy.editor import VideoFileClip

# PowerPoint to PNG conversion function remains the same
def ppt_to_png(ppt_file, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        presentation = powerpoint.Presentations.Open(ppt_file)
    except Exception as e:
        print(f"Error opening PowerPoint file: {e}")
        powerpoint.Quit()
        return

    for i, slide in enumerate(presentation.Slides):
        slide.Export(os.path.join(output_folder, f"slide_{i + 1}.png"), "PNG")

    presentation.Close()
    powerpoint.Quit()
    print(f"Slides saved as PNGs in '{output_folder}'")

def delete_presentation_images(output_folder):
    """Delete all PNG images in the output folder."""
    try:
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
            print(f"All slides deleted from '{output_folder}'")
    except Exception as e:
        print(f"Error deleting presentation images: {e}")

# Convert PowerPoint slides to PNG
ppt_file = r"C:\Users\kavya\Desktop\Hand-Gesture-Controlled-Presentation\Blue and Yellow Playful Doodle Digital Brainstorm Presentation.pptx"
output_folder = r"C:\Users\kavya\Desktop\Hand-Gesture-Controlled-Presentation\Presentation"

ppt_to_png(ppt_file, output_folder)

# Variables
cv2.namedWindow("Slides", cv2.WINDOW_NORMAL)
cv2.setWindowProperty("Slides", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
(x, y, screenWidth, screenHeight) = cv2.getWindowImageRect("Slides")
screenWidth -= 150
screenHeight -= 150
folderPath = output_folder  # Use the correct output folder
imgNumber = 0
cameraHeight, cameraWidth = int(120 * 1.2), int(213 * 1.2)
gestureThresholdHeight = 350
buttonPressed = False
buttonCounter = 0
buttonDelay = 15
annotations = [[]]
annotationNumber = 0
annotationStart = False
zoomFactor = 1.0
minZoom = 1.0
maxZoom = 4.0
record_video = False  # Flag to track video recording state
output_video_path = r"C:\Users\kavya\Desktop\Hand-Gesture-Controlled-Presentation\presentation_recording.avi"  # Video output path
output_audio_path = r"C:\Users\kavya\Desktop\Hand-Gesture-Controlled-Presentation\presentation_audio.mp3"  # Audio output path
fourcc = cv2.VideoWriter_fourcc(*"XVID")
video_writer = None  # Video writer object

# Gestures
leftSide = [1, 0, 0, 0, 0]
rightSide = [0, 0, 0, 0, 1]
pointerSide = [0, 1, 1, 0, 0]
drawSide = [0, 1, 0, 0, 0]
undoSide = [0, 1, 1, 1, 0]
zoominSide = [0, 1, 0, 0, 1]
zoomoutSide = [0, 1, 1, 1, 1]

# Hand Detector
detector = HandDetector(detectionCon=0.5, maxHands=1)

# Get list of presentation images
try:
    pathImages = sorted(os.listdir(folderPath), key=len)
except FileNotFoundError:
    print(f"Error: Folder '{folderPath}' not found. Please check the path.")
    exit()
slidesCount = len(pathImages)

# Camera Setup
cap = cv2.VideoCapture(0)
if not cap.isOpened():
    print("Error: Failed to open camera.")
    exit()
cap.set(3, screenWidth)
cap.set(4, screenHeight)

print("Press 'r' to start/stop recording. Press 'q' to quit.")

while True:
    success, camera = cap.read()
    if not success:
        print("Error: Failed to read camera frame.")
        break
    camera = cv2.flip(camera, 1)
    pathFullImage = os.path.join(folderPath, pathImages[imgNumber])
    imgCurrent = cv2.imread(pathFullImage)

    # Apply zoom
    newWidth, newHeight = int(screenWidth * zoomFactor), int(screenHeight * zoomFactor)
    imgCurrent = cv2.resize(imgCurrent, (newWidth, newHeight), interpolation=cv2.INTER_AREA)

    # Center zoomed image
    xOffset = (newWidth - screenWidth) // 2
    yOffset = (newHeight - screenHeight) // 2
    imgCurrent = imgCurrent[yOffset:yOffset + screenHeight, xOffset:xOffset + screenWidth]

    # Gesture handling
    hands, camera = detector.findHands(camera)
    cv2.line(camera, (0, gestureThresholdHeight), (screenWidth, gestureThresholdHeight), (0, 200, 200), 2)

    if hands and not buttonPressed:
        hand = hands[0]
        fingers = detector.fingersUp(hand)
        cx, cy = hand["center"]

        # Index finger landmarks
        landmarks = hand['lmList']
        xIndex = int(np.interp(landmarks[8][0], [screenWidth // 2, screenWidth], [0, screenWidth]))
        yIndex = int(np.interp(landmarks[8][1], [50, screenHeight - 50], [0, screenHeight]))
        indexFinger = xIndex, yIndex

        if cy <= gestureThresholdHeight:  # Hand on top right quarter
            # Gesture 1 - Left
            if fingers == leftSide:
                annotations = [[]]
                annotationNumber = 0
                annotationStart = False
                buttonPressed = True
                imgNumber = (imgNumber - 1) % slidesCount

            # Gesture 2 - Right
            if fingers == rightSide:
                annotations = [[]]
                annotationNumber = 0
                annotationStart = False
                buttonPressed = True
                imgNumber = (imgNumber + 1) % slidesCount

        # Gesture 3 - Pointer
        if fingers == pointerSide:
            cv2.circle(imgCurrent, indexFinger, 8, (0, 0, 0), cv2.FILLED)

        # Gesture 4 - Draw Pointer
        if fingers == drawSide:
            if not annotationStart:
                annotationStart = True
                annotationNumber += 1
                annotations.append([])
            cv2.circle(imgCurrent, indexFinger, 8, (0, 0, 255), cv2.FILLED)
            annotations[annotationNumber].append(indexFinger)
        else:
            annotationStart = False

        # Gesture 5 - Undo
        if fingers == undoSide:
            if annotationNumber > 0:
                annotationNumber -= 1
                annotations.pop(-1)
                buttonPressed = True

        # Gesture 6 - Zoom In
        if fingers == zoominSide and zoomFactor < maxZoom:
            zoomFactor += 0.25
            buttonPressed = True

        # Gesture 7 - Zoom Out
        if fingers == zoomoutSide and zoomFactor > minZoom:
            zoomFactor -= 0.25
            buttonPressed = True

    # Button presses delay
    if buttonPressed:
        buttonCounter += 1
        if buttonCounter > buttonDelay:
            buttonPressed = False
            buttonCounter = 0

    # Draw annotations
    for i in range(len(annotations)):
        for j in range(1, len(annotations[i])):
            cv2.line(imgCurrent, annotations[i][j - 1], annotations[i][j], (0, 0, 200), 8)

    # Adding camera footage on slides
    imgSmall = cv2.resize(camera, (cameraWidth, cameraHeight))
    h, w, _ = imgCurrent.shape
    imgCurrent[0:cameraHeight, w - cameraWidth:w] = imgSmall

    # Write the frame to video file if recording
    if record_video:
        if video_writer is None:
            video_writer = cv2.VideoWriter(output_video_path, fourcc, 20.0, (screenWidth, screenHeight))
        video_writer.write(imgCurrent)

    cv2.imshow("Slides", imgCurrent)

    key = cv2.waitKey(1)
    if key == ord('r'):
        record_video = not record_video  # Toggle recording
        if not record_video and video_writer is not None:
            video_writer.release()  # Stop recording
            video_writer = None
            print(f"Recording stopped. Video saved as '{output_video_path}'")
            # Convert video to audio
            video_clip = VideoFileClip(output_video_path)
            video_clip.audio.write_audiofile(output_audio_path)
            print(f"Audio extracted and saved as '{output_audio_path}'")
        elif record_video:
            print("Recording started.")

    if key == ord('q'):
        break

# Clean up
if video_writer is not None:
    video_writer.release()  # Ensure video writer is closed
cv2.destroyAllWindows()
cap.release()

# Auto-delete slides after presentation
delete_presentation_images(output_folder)
