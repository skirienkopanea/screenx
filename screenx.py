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
            print('Region saved in 3...')
            time.sleep(1)
            print('Region saved in 2...')
            time.sleep(1)
            print('Region saved in 1...')
            time.sleep(1)
            print('Region saved')
            
            break
        
    
    return region

def image_to_text(path):
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
            sendmail('email@example.com','Check Screen ' + str(datetime.datetime.now()),'Screen change detected',[output_path,original_path,diff_path])
        else:
            change = False
    time.sleep(1)
