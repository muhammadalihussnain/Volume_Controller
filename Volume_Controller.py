#################################################################################################
#                       Imports                                                                 #
#################################################################################################
import cv2 as cv
import mediapipe as mp
import numpy as np
from cvzone.HandTrackingModule import HandDetector
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
devices = AudioUtilities.GetSpeakers()
#################################################################################################
#                       Functions                                                               #
#################################################################################################
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

volume_range    =   volume.GetVolumeRange()
volume.SetMasterVolumeLevel(-20.0, None)
vol_min, vol_max = volume_range[0],volume_range[1]
print(vol_min, vol_max)
volume_bar  =   400
#################################################################################################
#                       Initialization                                                               #
#################################################################################################
cap             =   cv.VideoCapture(0)
image_width     =   620
image_height    =   480
cTime = 0
pTime = 0
cap.set(3, image_width)
cap.set(4, image_height)
detector = HandDetector(detectionCon=0.8, maxHands=1)
while True:
    success, image =   cap.read()
    hands, image     =   detector.findHands(image, flipType=False)
    if hands:
        if hands[0]["lmList"]:
            lmList = hands[0]["lmList"]
            fingure_pos =   lmList[8]
            thumb___pos =   lmList[8]
            #print(fingure_pos, thumb___pos)
            distance_bw_thumb_fing, _, _ = detector.findDistance(lmList[4], lmList[8], image)
            #print(distance_bw_thumb_fing)
            vol= np.interp(distance_bw_thumb_fing,[30,170],[vol_min,vol_max])
            volume_bar  =   np.interp(distance_bw_thumb_fing,[30,170],[400,150])
            print("value of volume bar is ",volume_bar)
            volume.SetMasterVolumeLevel(vol, None)

    else:
        print("Nothing Found")
    cv.rectangle(image, [50, 150], [85, 400], (0, 255, 0), 3)
    cv.rectangle(image, [50, int(volume_bar)], [85, 400], (0, 255, 0), cv.FILLED)
    cTime = time.time()
    fps = 1/(cTime - pTime)
    pTime = cTime
    cv.putText(image, f"FPS:{int(fps)}", (10, 40), cv.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 3)
    cv.imshow("preview", image)
    cv.waitKey(1)