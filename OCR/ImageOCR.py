import cv2
import numpy as np

def imageToText(image):
    """
    Takes an image and returns the text in it.
    """
    # Convert image to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # Apply edge detection to image
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    # Find contours in image
    contours, hierarchy = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # Sort contours by area
    contours = sorted(contours, key=cv2.contourArea, reverse=True)
    # Get largest contour
    cnt = contours[0]
    # Get bounding rectangle for largest contour
    x, y, w, h = cv2.boundingRect(cnt)
    # Create empty image
    empty = np.zeros_like(image)
    # Draw bounding rectangle around largest contour
    cv2.rectangle(empty, (x, y), (x+w, y+h), (255, 255, 255), 2)
    # Get contour of bounding rectangle
    cnt = contours[0]

