# python 3.12
# pip install pyautogui pillow, pyperclip
# davinci keyboard shortcuts:
#         -Console - Ctrl+Shift+`
#         -Reset UI Layout - Shift+P

# for the clicks you might need to make an "if mac > do a -10 or something on the y axis"
# because the fullscreen isn't truly fullscreen
# or maybe force davinci into fullscreen (if on a mac) and see where the clicks are then

import pyautogui
import pyperclip
import time
import sys
import os
import re

def wait():
    time.sleep(0.3)

screen_width, screen_height = pyautogui.size()

import pygetwindow as gw

def get_resolve_window():
    resolve_windows = gw.getWindowsWithTitle('DaVinci Resolve')
    return resolve_windows[0]

def click(x, y):

    x_percent = float(x / screen_width)
    y_percent = float(y / screen_height)

    resolve_window = get_resolve_window()
    
    x = int(resolve_window.left + x_percent * resolve_window.width)
    y = int(resolve_window.top + y_percent * resolve_window.height)
    
    pyautogui.click(x, y)
    wait()

def mouse_down(x, y):

    x_percent = float(x / screen_width)
    y_percent = float(y / screen_height)

    resolve_window = get_resolve_window()
    
    x = int(resolve_window.left + x_percent * resolve_window.width)
    y = int(resolve_window.top + y_percent * resolve_window.height)
    
    pyautogui.mouseDown(x, y)
    wait()
    return x, y

def move_mouse_to(x, y):
    pyautogui.moveTo(x, y)
    wait()
    return x, y

def mouse_up(x, y):    
    pyautogui.mouseUp(x, y)
    wait()

def press(key):
    pyautogui.press(key)
    wait()

def copy(string):
    pyperclip.copy(str(string))
    wait()

def paste():
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', 'v')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', 'v')
        wait()

def select_all():
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', 'a')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', 'a')
        wait()

def get_clip_duration():
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', 'd')
        wait()
        pyautogui.hotkey('ctrl', 'c')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', 'd')
        wait()
        pyautogui.hotkey('command', 'c')
        wait()

    paste()
    press('enter')

def set_clip_duration():
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', 'd')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', 'd')
        wait()

    paste()
    press('enter')

def move_clip_to_v2():
    press('home')
    x_down, y_down = mouse_down(670, 800)
    x_up, y_up = move_mouse_to(x_down, (y_down - 50))
    mouse_up(x_up, y_up)

def extract_sts_number(filename):
    match = re.search(r'(\d{4})', filename)
    if match:
        return match.group(1)
    return 000

def resetUI():
    resolve.OpenPage('edit')
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', '`')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', '`')
        wait()

################################################

resolve = app.GetResolve()

projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
media_pool = project.GetMediaPool()
media_storage = resolve.GetMediaStorage()

frames = 59.94
width = 1920
height = 1080

project.SetSetting("timelineFrameRate", str(frames))
project.SetSetting("timelineResolutionWidth", str(width))
project.SetSetting("timelineResolutionHeight", str(height))

################################################
# Sequence Below #
################################################

resetUI()

clip_list = media_pool.GetRootFolder().GetClipList()

welcome_bookend = None
closing_bookend = None
english_dvd_slate = None
french_dvd_slate = None
spanish_dvd_slate = None
nature_video = None

for clip in clip_list:
    name = str(clip.GetName()).lower()

    if "welcome" in name:
        welcome_bookend = clip
    elif "closing" in name:
        closing_bookend = clip
    elif "eng" in name and "slate" in name:
        english_dvd_slate = clip
    elif "frn" in name and "slate" in name:
        french_dvd_slate = clip
    elif "spn" in name and "slate" in name:
        spanish_dvd_slate = clip
    elif "mp4" in name or "mov" in name:
        nature_video = clip

def make_video(clip):
    resetUI()
    clip_name = clip.GetName()
    media_pool.CreateEmptyTimeline("Timeline " + clip_name[:-4])
    media_pool.AppendToTimeline(clip)

    time.sleep(1)
    move_clip_to_v2()

    media_pool.AppendToTimeline(nature_video)

    current_timeline = project.GetCurrentTimeline()
    timeline_clips = current_timeline.GetItemsInTrack('video', 1)
    clip_2 = timeline_clips.get(1)
    clip_2.SetProperty("RetimeProcess", 3)

    if "closing" in clip_name:
        clip_2.SetProperty("Pan", -850)

    press('home')
    click(670, 800)
    get_clip_duration()

    click(670, 725)
    set_clip_duration()

    sts_number = extract_sts_number(clip_name)
    target_dir = os.path.join(os.path.expanduser("~/Desktop"), "Rendered Bookends", sts_number)
    os.makedirs(target_dir, exist_ok=True)

    # Add to Render Queue
    project.SetRenderSettings({
        'TargetDir': target_dir,
        'CustomName': clip_name[:-4],
        'FrameRate': 59.94,
        'VideoCodec': 'mp4'
    })
    project.AddRenderJob()
    click(1723, 493)
    
    resetUI()

if nature_video is not None:
    # Check if at least one bookend or slate is present
    if welcome_bookend is not None or english_dvd_slate is not None or \
       french_dvd_slate is not None or spanish_dvd_slate is not None or \
       closing_bookend is not None:

        if welcome_bookend is not None:
            make_video(welcome_bookend)
        if english_dvd_slate is not None:
            make_video(english_dvd_slate)
        if french_dvd_slate is not None:
            make_video(french_dvd_slate)
        if spanish_dvd_slate is not None:
            make_video(spanish_dvd_slate)
        if closing_bookend is not None:
            make_video(closing_bookend)

        project.StartRendering()
    
    else:
        print('No bookend or slate picture present')
        print('OR')
        print("Bookend / slates don't have proper naming")
        print("Bookend / slates needs to be named as the following:")
        print("'#### Welcome/Closing bookend'")
        print("'#### ENG/SPN/FRN DVD Slate'")
        print("Where '####' refers to the correct STS number")
        print('')

        time.sleep(1)
        pyautogui.hotkey('shift', 'P')

else:
    print('No nature video present')
    print('OR')
    print('Rename nature video to "nature"')
    print('')

    time.sleep(1)
    pyautogui.hotkey('shift', 'P')

# TODO adjust targetDir
# Open console when fully rendered and say
# "Bookends rendered to ..."
