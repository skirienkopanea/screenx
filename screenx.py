"""
Author: Sergio Kirienko (GitHub: skirienkopanea)

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
"""

import win32gui, win32api, win32con
from win32api import GetSystemMetrics
import pyautogui
from PIL import ImageChops, Image, ImageDraw
import time
import mouse
import keyboard
import os
import win32com.client as win32
import datetime
import pytesseract
import win32com.client as wincl

def sendmail(to,subject,body,attachments):
    """
    Send an email using Microsoft Outlook.

    Parameters:
    - to (str): Email address of the recipient.
    - subject (str): Subject of the email.
    - body (str): Body/content of the email.
    - attachments (list): List of file paths for attachments.

    Note:
    - Microsoft Outlook must be installed and configured on your machine.
    - This function uses the win32com library which is only available on Windows.

    Returns:
    - None

    Raises:
    - FileNotFoundError: If any of the attachment files are not found.

    Example:
    >>> sendmail('example@example.com', 'Test Email', 'This is a test email.', ['attachment1.pdf', 'attachment2.docx'])
    Email sent
    """

    #outlook has to be available and configured on your machine
    olApp = win32.Dispatch('Outlook.Application')

    # construct the email item object
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = subject
    mailItem.BodyFormat = 1
    mailItem.Body = body
    mailItem.To = to

    for a in attachments:
        mailItem.Attachments.Add(os.path.join(os.getcwd(), a))

    mailItem.Display()

    mailItem.Save()
    mailItem.Send()
    print("Email sent")

def drawbox(region):
    """
    Draw a red box around a specified region on the screen.

    Parameters:
    - region (tuple): A tuple representing the region to draw the box around.
        It should be in the format (x_coordinate, y_coordinate, width, height).

    Note:
    - This function utilizes the win32gui library, which is only available on Windows.
    - The box will be drawn using red color.
    - The region coordinates are relative to the top-left corner of the screen.
    - It may not work in additional monitors since their coordinates might have negative numbers or be offsetted by some amount. Only use un default monitor.

    Returns:
    - None

    Example:
    >>> drawbox((100, 100, 200, 150))

    This will draw a red box starting at coordinates (100, 100) with a width of 200 pixels
    and a height of 150 pixels.
    """
    #region,None,False
    dc = win32gui.GetDC(0)
    hwnd = win32gui.WindowFromPoint((0,0))
    monitor = (0, 0, GetSystemMetrics(0), GetSystemMetrics(1))

    red = win32api.RGB(255, 0, 0) # Red

    past_coordinates = monitor
    x_coordinate = region[0]
    y_coordinate = region[1]
    width = region[2]
    height = region[3]
    past_coordinates = (past_coordinates[0]+x_coordinate, past_coordinates[1]+y_coordinate, past_coordinates[2], past_coordinates[3])

    #refreshed square in some apps: Chrome, clicking in another window often works too.
    rect = win32gui.CreateRoundRectRgn(*past_coordinates, 2 , 2)
    win32gui.RedrawWindow(hwnd, past_coordinates, rect, win32con.RDW_INVALIDATE)      
    for x in range(width):
        win32gui.SetPixel(dc, past_coordinates[0]+x, past_coordinates[1], red)
        win32gui.SetPixel(dc, past_coordinates[0]+x, past_coordinates[1]+height, red)
    for y in range(height):        
        win32gui.SetPixel(dc, past_coordinates[0], past_coordinates[1]+y, red)
        win32gui.SetPixel(dc, past_coordinates[0]+width, past_coordinates[1]+y, red)

def highlight_different_pixels(image1_path, image2_path, output_path):
    """
    Highlight the differing pixels between two images and save the result as a new image.

    Parameters:
    - image1_path (str): Path to the first image file.
    - image2_path (str): Path to the second image file.
    - output_path (str): Path to save the resulting image with highlighted differences.

    Returns:
    - float: Percentage of differing pixels between the two images.

    Raises:
    - ValueError: If the images have different dimensions.

    Note:
    - This function requires the PIL (Python Imaging Library) package to be installed.
    - The resulting image will highlight differing pixels in red.

    Example:
    >>> percentage_diff = highlight_different_pixels("image1.png", "image2.png", "output.png")
    >>> print(f"Percentage of differing pixels: {percentage_diff}%")
    Percentage of differing pixels: 3.2%
    """    
    # Open the images
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path)

    # Convert images to RGB mode (in case they are in different modes)
    image1 = image1.convert("RGB")
    image2 = image2.convert("RGB")

    # Check if images have the same dimensions
    if image1.size != image2.size:
        raise ValueError("Images must have the same dimensions")

    # Create a copy of original image for highlighting differences
    diff_image = image1.copy()
    draw = ImageDraw.Draw(diff_image)

    # Compare pixels and highlight differences
    # Get the width and height of the images
    width, height = image1.size
    differing_pixels = 0
    for x in range(image1.width):
        for y in range(image1.height):
            pixel1 = image1.getpixel((x, y))
            pixel2 = image2.getpixel((x, y))
            if pixel1 != pixel2:
                draw.point((x, y), fill="red")
                differing_pixels += 1
    
    total_pixels = width * height
    percentage_different = (differing_pixels / total_pixels) * 100

    diff_image.save(output_path)
    return percentage_different

def detect_screen_changes(region,original_path,diff_threshold_percentage,output_path,diff_path):
    """
    Detect changes on the screen within a specified region and save the differences.

    Parameters:
    - region (tuple): A tuple representing the region of the screen to capture changes.
        It should be in the format (x, y, width, height).
    - original_path (str): Path to the original reference image.
    - diff_threshold_percentage (float): Threshold percentage for considering a screen change significant.
    - output_path (str): Path to save the current screenshot.
    - diff_path (str): Path to save the image highlighting the differences.

    Returns:
    - bool: True if a significant screen change is detected, otherwise False.

    Note:
    - This function utilizes the pyautogui, PIL (Python Imaging Library), and datetime libraries.
    - The function captures the current screen within the specified region and compares it to the original image.
    - If a significant difference is detected (exceeding the provided threshold), it saves the current screenshot and highlights the differences.
    """    
    print("check " + str(datetime.datetime.now()))
    current_image = pyautogui.screenshot(region=region)
    original_image = Image.open(original_path)
    diff = ImageChops.difference(original_image, current_image)
    if diff.getbbox() is not None:
        current_image.save(output_path)
        diff_percentage = highlight_different_pixels(original_path, output_path, diff_path)
        print(f'{round(diff_percentage,2)}% screen change detected!') 
        if(diff_percentage>=diff_threshold_percentage): return True
    return False

def getRegion():
    """
    Capture a region of the screen defined by the user through mouse clicks.

    Returns:
    - tuple: A tuple representing the captured region in the format (x, y, width, height).

    Note:
    - This function requires the pyautogui, mouse, and time libraries.
    - The user is prompted to select a region by clicking and dragging the mouse.
    - Once the region is selected, the function returns a tuple representing the region.

    Example:
    >>> region = getRegion()
    [(100, 200, 300, 400)]
    """
    click = None
    click2 = None
    x_coordinate = None
    y_coordinate = None
    width = None
    height = None
    region = None
    
    while True:
        if mouse.is_pressed(button='left') and click is None:
            click = pyautogui.position()
            print(click)
        if click is not None and mouse.is_pressed(button='left'):
            click2 = pyautogui.position()
            x_coordinate = min(click[0],click2[0])
            y_coordinate = min(click[1],click2[1])
            width = abs(click2[0] - click[0])
            height = abs(click2[1] - click[1])
            region = (x_coordinate, y_coordinate, width, height)
            drawbox(region)
            
        if click is not None and click2 is not None and not mouse.is_pressed(button='left'):
            print(click2)
            print('Region saved')
            
            break
        
    
    return region

def image_to_text(path):
    """
    Convert an image to text using OCR (Optical Character Recognition).

    Parameters:
    - path (str): Path to the image file.

    Returns:
    - str: The extracted text from the image.

    Note:
    - This function requires the pytesseract and PIL (Python Imaging Library) libraries.
    - The function uses pytesseract to perform OCR on the image and extract text.
    - The extracted text is returned after removing leading/trailing whitespaces and collapsing multiple spaces into single spaces.

    Example:
    >>> text = image_to_text("image.png")
    >>> print(text)
    "This is a sample text extracted from an image."
    """    
    # Simple image to string
    return ' '.join(pytesseract.image_to_string(Image.open(path)).split())

def text_to_speech(text,language):
    #This depends on your pc
    speaker_number = 1    
    if language =='es':
        speaker_number = 0
    spk = wincl.Dispatch("SAPI.SpVoice")
    vcs = spk.GetVoices()
    spk.Voice
    spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)
    spk.Speak(text)

def loop_click(x,y):
    """
    Move the mouse cursor to a specified position and perform a left-click.

    Parameters:
    - x (int): The x-coordinate of the target position.
    - y (int): The y-coordinate of the target position.

    Returns:
    - None

    Note:
    - This function requires the pyautogui library.
    - The function moves the mouse cursor to the specified position (x, y) and performs a left-click.

    Example:
    >>> loop_click(100, 200)
    """    
    pyautogui.moveTo(x,y)    #teams
    pyautogui.click()

# If you don't have tesseract executable in your PATH, include the following:
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
#only use the languages you have available as shown from here
lang = pytesseract.get_languages(config='')[0]
print(lang)

matching_text = input("Enter matching words or just hit enter key to contiune for pixel change above 5%\n")
print("\nClick and drag to select region to watch")
region = getRegion() 
adjusted_region = (region[0]+1,region[1]+1,region[2]-2,region[3]-2) #excludes 1px border
original_path = "original.png"
output_path = "output.png"
diff_path = "diff.png"
original = pyautogui.screenshot(region=adjusted_region) #keep original as reference only once
original.save("original.png")
change = False
clickloop = input("Loop click through teams? (Y/any)\n").upper() == "Y"
sendmailb = input("Send mail? (Y/any)\n").upper() == "Y"


#program watch loop
while not keyboard.is_pressed('esc') and not change:
    drawbox(region)
    change = detect_screen_changes(adjusted_region,original_path,5,output_path,diff_path)
    if change:
        text = None
        if len(matching_text) > 0:
            text = image_to_text(output_path)
            print(text)
        if len(matching_text) == 0 or matching_text.lower() in text.lower():
            if text is not None:
                text_to_speech(text,lang)
            if sendmailb: sendmail('kirienkosergio@gmail.com','Check Screen ' + str(datetime.datetime.now()),'Screen change detected',[output_path,original_path,diff_path])
        else:
            change = False
    time.sleep(1)
    if clickloop: loop_click(763, 1051)    #teams  coordinates

# keep loop after notification
while not keyboard.is_pressed('esc') and clickloop:
     
    time.sleep(1)
    loop_click(763, 1051)    #teams  coordinates

print("End")
