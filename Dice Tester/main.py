import keyboard
import xlwt
import cv2
import numpy as np
from collections import deque

style0 = xlwt.easyxf('font: name Arial, color-index black, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf ('font: name Arial, color-index black, bold off')

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')

ws.write(0, 0, "#", style0)
ws.write(0, 1, "Numero 1", style0)
ws.write(0, 2, "Numero 2", style0)
ws.write(0, 3, "Numero 3", style0)
ws.write(0, 4, "Numero 4", style0)
ws.write(0, 5, "Numero 5", style0)
ws.write(0, 6, "Numero 6", style0)

n1 = 0
n2 = 0
n3 = 0 
n4 = 0
n5 = 0
n6 = 0

min_threshold = 10                      # these values are used to filter our detector.
max_threshold = 200                     # they can be tweaked depending on the camera distance, camera angle, ...
min_area = 100                          # ... focus, brightness, etc.
min_circularity = 0.3
min_inertia_ratio = 0.5
 
cap = cv2.VideoCapture(0)               # '0' is the webcam's ID. usually it's 0/1/2/3/etc. 'cap' is the video object.
cap.set(15, -4)                         # '15' references video's exposure. '-4' sets it.
 
counter = 0                             # script will use a counter to handle FPS.
readings = deque([0, 0], maxlen=10)     # lists are used to track the number of pips.
display = deque([0, 0], maxlen=10)
 
while True:
    ret, im = cap.read()                                    # 'im' will be a frame from the video.
 
    params = cv2.SimpleBlobDetector_Params()                # declare filter parameters.
    params.filterByArea = True
    params.filterByCircularity = True
    params.filterByInertia = True
    params.minThreshold = min_threshold
    params.maxThreshold = max_threshold
    params.minArea = min_area
    params.minCircularity = min_circularity
    params.minInertiaRatio = min_inertia_ratio
 
    detector = cv2.SimpleBlobDetector_create(params)        # create a blob detector object.

    keypoints = detector.detect(im)                         # keypoints is a list containing the detected blobs.
 
    # here we draw keypoints on the frame.
    im_with_keypoints = cv2.drawKeypoints(im, keypoints, np.array([]), (0, 0, 255),
                                          cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
 
    cv2.imshow("Dice Reader", im_with_keypoints)            # display the frame with keypoints added.
 
    if counter % 10 == 0:                                   # enter this block every 10 frames.
        reading = len(keypoints)                            # 'reading' counts the number of keypoints (pips).
        readings.append(reading)                            # record the reading from this frame.
 
        if readings[-1] == readings[-2] == readings[-3]:    # if the last 3 readings are the same...
            display.append(readings[-1])                    # ... then we have a valid reading.
 
        # if the most recent valid reading has changed, and it's something other than zero, then print it.
        if display[-1] != display[-2] and display[-1] != 0:
            msg = f"{display[-1]}\n****"
            print(msg)

            if display[-1] == 1:
                n1 = n1 + 1
                print (n1)
            if msg == 2:
                n2 = n2 + 1
                print (n2)
            if msg == "3":
                n3 = n3 + 1
                print (n3)
            if msg == "4":
                n4 = n4 + 1
                print (n4)
            if msg == "5":
                n5 = n5 + 1
                print (n5)
            if msg == "6":
                n6 = n6 + 1
                print (n6)
            
    counter += 1
 
    if cv2.waitKey(1) & 0xff == 27: 
        ws.write (1, 1, n1, style1)
        ws.write (1, 2, n2, style1)
        ws.write (1, 3, n3, style1)
        ws.write (1, 4, n4, style1)
        ws.write (1, 5, n5, style1)
        ws.write (1, 6, n6, style1)
        wb.save('example.xls')
        break                            # press [Esc] to exit.

cv2.destroyAllWindows()
