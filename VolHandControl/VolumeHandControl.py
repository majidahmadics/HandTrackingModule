# Description: This file is used to control the volume of the system using the hand gestures.

# Import the necessary libraries
import cv2
import time
import numpy as np
import HandTrackingModule as htm
import math
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


# Define the width and height of the camera
wCam, hCam = 640, 480


# Check The Camera and successfully open it and capture the image
cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
pTime = 0

detector = htm.HandDetector(detectionCon=0.7) # Create an object of HandDetector class


# Using pycaw library to control the volume of the system
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = interface.QueryInterface(IAudioEndpointVolume)

# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
# End of pycaw library Block


vol = 0
volBar = 400
volPer = 0

# Main Loop
while True:
    #Capture the image from the camera and find the hands Landmarks
    success, img = cap.read()
    img = detector.findHands(img)
    lmList = detector.findPosition(img, draw=False)
    if len(lmList) != 0:
        
        
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        
        cv2.circle(img, (x1, y1), 10, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 10, (255, 0, 255), cv2.FILLED)
        
        # Draw a line between the two points and a circle at the center of the line
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
        cv2.circle(img, (cx, cy), 10, (255, 0, 255), cv2.FILLED)
        
        # Calculate the length of the line
        length = math.hypot(x2 - x1, y2 - y1)
        
        
        # Min hand range is 50 and Max hand range is 250
        # Volume range is from -65.25 to 0
        # Using np.interp to convert the length of the line to the volume range
        
        vol = np.interp(length, [50, 250], [minVol, maxVol])
        volBar = np.interp(length, [50, 250], [400, 150])
        volPer = np.interp(length, [50, 250], [0, 100])
        
        
        # Set the volume of the system equal to the volume calculated based on the length of the line
        volume.SetMasterVolumeLevel(vol, None)
        
        
    cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
    cv2.putText(img, f"Vol: {int(volPer)}%", (40, 450), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)
    
    
    cTime = time.time()
    fps = 1/(cTime - pTime)
    pTime = cTime
    
    cv2.putText(img, f"FPS: {int(fps)}", (40, 50), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)
    cv2.imshow("Image", img)
    cv2.waitKey(1)