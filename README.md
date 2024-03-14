This code monitors a specific region of the screen for changes, and if a significant change is detected,
it performs a series of actions including capturing a screenshot, comparing it with the original image,
extracting text from the changed region (optional), notifying the user with text-to-speech, and sending an email notification.

Before running the code, ensure that Tesseract OCR is installed and its executable path is correctly set.
Also, make sure to have the necessary libraries installed (pyautogui, keyboard, pytesseract, etc.).

The code workflow is as follows:
1. Set up Tesseract OCR executable path and obtain available languages.
2. Prompt the user to input matching words to search for in the changed region or continue without filtering.
3. Prompt the user to select a region of the screen to monitor.
4. Capture the original reference image of the selected region.
5. Enter a loop to continuously monitor the screen until the 'esc' key is pressed or a change is detected.
6. If a change is detected, extract text from the changed region (if provided), perform further actions based on text matching,
   such as text-to-speech and email notification with screenshots.
7. If no change is detected or the text doesn't match, continue monitoring.
8. After notification, enter another loop to keep the program running until the 'esc' key is pressed.
