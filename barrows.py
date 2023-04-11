import pyautogui
from pyautogui import keyDown, keyUp
import time
import win32gui
import pytesseract
import win32com.client
import sys
import random

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'  
# your path may be different

def find_window():
    #Look through all your windows
    def search(handle, window):
        text = win32gui.GetWindowText(handle)
        rect = win32gui.GetWindowRect(handle)
        #Get coordinates
        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
        #If the window has RuneWild, it saves the coordinates
        if "runelite" in text.lower():
            print(f"Window {win32gui.GetWindowText(handle)}:")
            print(f"\tLocation: ({x},{y})")
            print(f"\tSize: ({w},{h})")
            window.append({'handle':handle,'x':x,'y':y,'w':w,'h':h})
    #Store all windows that match and output found window, will be first in list (and only)
    window = []
    win32gui.EnumWindows(search, window)
    return window[0]

#Get window dimensions
window = find_window()
print(window)
x,y,w,h = window['x'], window['y'], window['w'], window['h']
print(x,y,w,h)

#Set window into foreground
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')

win32gui.SetForegroundWindow(window['handle'])
time.sleep(1)

def long_click(input_x,input_y,button="left"):
    #Sometimes click happens too fast to register, manual for better responsiveness
    pyautogui.moveTo(input_x,input_y)
    pyautogui.mouseDown(button=button)
    time.sleep(0.4)
    pyautogui.mouseUp(button=button)

def healing():
    while True:
        # Healing
        hp = pyautogui.pixelMatchesColor(hp_x, hp_y, (19, 19, 19))
        if hp:
            print('need hp')
            monkfish = pyautogui.locateCenterOnScreen('monkfish.png', confidence=0.9)
            if monkfish:
                long_click(monkfish, None)
                long_click(monkfish, None)
            else:
                break
        else:
            break

#Initialized PyAutoGui    
pyautogui.FAILSAFE = True
  
# ---------- VARIABLES SECTION ----------

#Display tab coordinates
display_tab = x+574, y+256

#Brightness level coordinates
brightness = x+711, y+333

# Coordinates for Hitpoint healer
hp_x = x + 550
hp_y = y + 85

offset_x = random.randint(-5, 5)
offset_y = random.randint(-5, 5)

#Countdown timer
print("Starting", end="")
for i in range(0,3):
    print(".", end="")
    time.sleep(1)
print("Go")


# ---------- SETTING THE BASICS START ----------

#Click on compass
long_click(x+564, y+48)

#Move Mouse in center of window
pyautogui.moveTo(x+w/2, y+h/2)

#Move camera up
keyDown("up")
time.sleep(1.5)
keyUp("up")

# Click on setting tab
keyDown("F10")
time.sleep(0.5)

#Click on display setting
long_click(display_tab, None)
time.sleep(0.5)

#Set the brightness level
long_click(brightness, None)
time.sleep(0.5)

# click on inventory tab
keyDown("F1")
time.sleep(0.5)

#Move Mouse in center of window
pyautogui.moveTo(x+w/2, y+h/2)
time.sleep(0.2)

#Scroll up a lot
for i in range(10):
    pyautogui.scroll(-550)
    time.sleep(0.1)

#Scroll up a lot
for i in range(10):
    pyautogui.scroll(250)
    time.sleep(0.1)

# ---------- SETTING THE BASICS END ----------

# INITIAL LOOP
num = 1
for i in range(20):
    print("Master Loop", num)
    num += 1

    # Go home
    pyautogui.write('::home')
    time.sleep(0.2)
    keyDown('enter')
    time.sleep(0.5)
    pyautogui.write('::home')
    time.sleep(0.2)
    keyDown('enter')
    time.sleep(3)
    
    # From home, go bank from minimap
    long_click(x + 658, y + 83)
    time.sleep(3.5)
    
    # Open bank
    long_click(x+ 321, y + 178)
    time.sleep(0.5)
    deposit = pyautogui.locateCenterOnScreen('deposit.png', confidence=0.9)
    if deposit:
        long_click(deposit, None)
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        
    # Click minimap reset pool
    long_click(x + 625, y + 128)
    time.sleep(3.5)
    
    # Click on pool
    long_click(x + 181, y + 212)
    time.sleep(0.5)
    
    # Click on quest tab
    long_click(x + 616, y + 212)
    time.sleep(0.5)
    
    # Click on load out tab
    long_click(x + 606, y + 242)
    time.sleep(0.5)
    
    # Click on barrows load out
    barrows = pyautogui.locateCenterOnScreen('barrows.png', confidence=0.9)
    if barrows:
        long_click(barrows, None)
        time.sleep(0.5)
          
    pyautogui.press('F1')
    time.sleep(0.5)
    
    rune_bolt = pyautogui.locateOnScreen('rune_bolt.png', confidence=0.9)
    if not rune_bolt:
        print('Rune Bolt not found in inv')
        sys.exit()
        
    fire_rune = pyautogui.locateOnScreen('fire_rune.png', confidence=0.9)
    if not fire_rune:
        print('Fire Rune not found in inv')
        sys.exit()
        
    death_rune = pyautogui.locateOnScreen('death_rune.png', confidence=0.9)
    if not death_rune:
        print('Death Rune not found in inv')
        sys.exit()
        
    monkfish = pyautogui.locateOnScreen('monkfish.png', confidence=0.9)
    if not monkfish:
        print('Monkfish not found in inv')
        sys.exit()
    
    # Open mage tab
    pyautogui.press('F3')
    time.sleep(0.2)
    
    # Click on teleport
    long_click(x + 567, y + 247)
    time.sleep(0.5)
    
    # Click on minigames
    minigames = pyautogui.locateCenterOnScreen('minigames.png', confidence=0.9)
    if minigames:
        long_click(minigames, None)
        time.sleep(0.2)
        barrows_tele = pyautogui.locateCenterOnScreen('barrows_tele.png', confidence=0.9)
        if barrows_tele:
            long_click(barrows_tele, None)
            time.sleep(3)
            
    # Open inv tab    
    pyautogui.press('F1')
    time.sleep(0.5)
    
    # Make sure magic is set
    
    # Open combat tab    
    pyautogui.press('F5')
    time.sleep(0.5)
    
    # Click on spell
    long_click(x + 687, y + 300) # remove 50 for spell
    time.sleep(1)
    
    iban = pyautogui.locateCenterOnScreen('iban_spell.png', confidence=0.9)
    if iban:
        long_click(iban, None)
        time.sleep(0.5)
        
    # Open inv tab    
    pyautogui.press('F1')
    time.sleep(0.2)
        
    
    # ---------- BARROWS START ----------
    
    # ---------- VERAC ----------
    print('Vera Started')
    # Move to VERAC
    long_click(x + 616, y + 152)
    time.sleep(4)
    
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2)
    
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click VERAC sarcophagus
        long_click(x + 34, y + 188, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(0.5)
            break
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Climp up VERAC staircases
        long_click(x + 422, y + 292, button='right')
        time.sleep(0.2)
        staircases = pyautogui.locateCenterOnScreen('staircases.png', confidence=0.9)
        if staircases:
            long_click(staircases, None)
            time.sleep(5)
            break
    
    # ---------- DHAROK ----------
    print('Dharok Started')
    # Move to DHAROK
    long_click(x + 717, y + 112)
    time.sleep(6)
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2)
          
    # Activate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click DHAROK sarcophagus
        long_click(x + 192, y + 352, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(0.5)
            break
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    # Deactivate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Climp up DHAROK staircases
        long_click(x + 374, y + 110, button='right')
        time.sleep(0.5)
        staircases = pyautogui.locateCenterOnScreen('staircases.png', confidence=0.9)
        if staircases:
            long_click(staircases, None)
            time.sleep(5)
            break
    
    
    # ---------- AHRIM ----------
    print('Ahrim Started')
    # Move to AHRIM
    long_click(x + 615, y + 147)
    time.sleep(6)
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2)
    
    #Move Mouse in center of window
    pyautogui.moveTo(x+w/2, y+h/2)
    time.sleep(0.2)    
    
    # Switch gear to range
    rune_crossbow = pyautogui.locateCenterOnScreen('rune_crossbow.png', confidence=0.70)
    if rune_crossbow:
        long_click(rune_crossbow, None)
        time.sleep(0.5)
    else:
        print(' Rune Crossbow not found')
    
    #Move Mouse in center of window
    pyautogui.moveTo(x+w/2, y+h/2)
    time.sleep(0.2)
    
    # Switch gear to range
    dragon_crossbow = pyautogui.locateCenterOnScreen('dragon.png', confidence=0.70)
    if dragon_crossbow:
        long_click(dragon_crossbow, None)
        time.sleep(0.5)
    else:
        print('Dragon Crossbow not found')
    
    #Move Mouse in center of window
    pyautogui.moveTo(x+w/2, y+h/2)
    time.sleep(0.2)
        
    rune_bolt = pyautogui.locateCenterOnScreen('rune_bolt.png', confidence=0.8)
    if rune_bolt:
        long_click(rune_bolt, None)
        time.sleep(0.5)
    
    # Activate Mage Prayers
    pyautogui.press('F2')
    time.sleep(0.5)
    # Click on anti mage
    long_click(x + 609, y + 371)
    time.sleep(0.5)
    pyautogui.press('F1')
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click AHRIM sarcophagus
        long_click(x + 192, y + 352, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(0.5)
            break
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    # Deactivate Mage Prayers
    pyautogui.press('F2')
    time.sleep(0.5)
    # Click on anti mage
    long_click(x + 609, y + 371)
    time.sleep(0.5)
    pyautogui.press('F1')
     
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Climp up AHRIM staircases
        long_click(x + 374, y + 110, button='right')
        time.sleep(0.2)
        staircases = pyautogui.locateCenterOnScreen('staircases.png', confidence=0.9)
        if staircases:
            long_click(staircases, None)
            time.sleep(5)
            break
    
    iban_staff = pyautogui.locateCenterOnScreen('iban_staff.png', confidence=0.8)
    if iban_staff:
        long_click(iban_staff, None)
        time.sleep(0.5)
    
    #Move Mouse in center of window
    pyautogui.moveTo(x+w/2, y+h/2)
    time.sleep(0.2)
    
    blessing = pyautogui.locateCenterOnScreen('blessing.png', confidence=0.8)
    if blessing:
        long_click(blessing, None)
        time.sleep(0.5)
    
    
    # ---------- TORAG ----------
    print('Torag Started')
    # Move to TORAG
    long_click(x + 595, y + 134)
    time.sleep(6)
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2)
       
    # Activate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click TORAG sarcophagus
        long_click(x + 337, y + 52, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(0.5)
            break
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    # Deactivate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Climp up TORAG staircases
        long_click(x + 112, y + 247, button='right')
        time.sleep(0.2)
        staircases = pyautogui.locateCenterOnScreen('staircases.png', confidence=0.9)
        if staircases:
            long_click(staircases, None)
            time.sleep(4)
            break
        
    
    # ---------- KARIL ----------
    print('Karil Started')
    # Move to KARIL
    long_click(x + 695, y + 137)
    time.sleep(6)
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2)
    
    # Activate Range Prayers
    pyautogui.press('F2')
    time.sleep(0.5)
    # Click on anti range
    long_click(x + 647, y + 372)
    time.sleep(0.5)
    pyautogui.press('F1')
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click KARIL sarcophagus
        long_click(x + 497, y + 232, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(0.5)
            break
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    # Deactivate Range Prayers
    pyautogui.press('F2')
    time.sleep(0.5)
    # Click on anti range
    long_click(x + 647, y + 372)
    time.sleep(0.5)
    pyautogui.press('F1') 
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Climp up KARIL staircases
        long_click(x + 127, y + 102, button='right')
        time.sleep(0.2)
        staircases = pyautogui.locateCenterOnScreen('staircases.png', confidence=0.9)
        if staircases:
            long_click(staircases, None)
            time.sleep(5)
            break
       
    
    # ---------- GUTHAN ----------
    print('Guthan Started')
    # Move to GUTHAN
    long_click(x + 697, y + 92)
    time.sleep(6)
    # Click on spade
    spade = pyautogui.locateCenterOnScreen('spade.png', confidence=0.9)
    if spade:
        long_click(spade, None)
        time.sleep(2) 
    
    for i in range(10):
        #Move Mouse in center of window
        pyautogui.moveTo(x+w/2, y+h/2)
        time.sleep(0.2)  
        # Right Click GUTHAN sarcophagus
        long_click(x + 502, y + 209, button='right')
        time.sleep(0.5)
        search = pyautogui.locateCenterOnScreen('search_sarcophagus.png', confidence=0.9)
        if search:
            long_click(search, None)
            time.sleep(3)
            break
    
    time.sleep(0.5)
    #Move camera down
    keyDown("down")
    time.sleep(1.5)
    keyUp("down")
    time.sleep(0.5)
    
    while True:
        time.sleep(0.2)
        # Click on chest
        pyautogui.moveTo(x + 292 + offset_x, y + 48 + offset_y)
        open_chest = pyautogui.locateOnScreen('open_chest.png', confidence=0.8)
        if open_chest:
            pyautogui.click(clicks=2, interval=0.25)
            time.sleep(0.2)
            someone = pyautogui.locateOnScreen('someone.png', confidence=0.9)
            if not someone:
                print('Someone not found, continuing...')
                break
    
    # Activate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    #Move camera up
    keyDown("up")
    time.sleep(1.5)
    keyUp("up")
    
    num = 1
    # Heal during fight
    for i in range(200):
        time.sleep(0.2)
        print('loop', num)
        num += 1
        healing()
        nohp = pyautogui.locateOnScreen('0hp.png', region=(x+10, y+47, 135, 80))
        if nohp:
            print('0 HP Detected')
            break
        
    # Deactivate Quick Prayers
    long_click(x + 560, y + 118)
    time.sleep(0.5)
    
    for i in range(10):
        time.sleep(0.2)
        # Click on chest
        pyautogui.moveTo(x + 287 + offset_x, y + 143 + offset_y)
        open_chest = pyautogui.locateOnScreen('open_chest.png', confidence=0.8)
        if open_chest:
            pyautogui.click(clicks=2, interval=0.25)
            time.sleep(1)
            break

    time.sleep(3)



