import time
import sys
import os
import re
import pyautogui
import pyperclip

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

def resetUI():
    resolve.OpenPage('edit')
    wait()
    pyautogui.hotkey('shift', '0')
    wait()

#################################################################

resolve = app.GetResolve()

projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
media_pool = project.GetMediaPool()
media_storage = resolve.GetMediaStorage()

frames = 59.94
width = 1280
height = 720

project.SetSetting("timelineFrameRate", str(frames))
project.SetSetting("timelineResolutionWidth", str(width))
project.SetSetting("timelineResolutionHeight", str(height))

resetUI()	

root_folder = media_pool.GetRootFolder()
sub_folders = root_folder.GetSubFolderList() #list

for folder in sub_folders:
    print(folder.GetName())

#################################################################

media_pool.SetCurrentFolder(sub_folders[0]) #Switch to Timeline Folder

sp_clips = sub_folders[1].GetClipList()
fr_clips = sub_folders[2].GetClipList()
video_clips = sub_folders[3].GetClipList()
wait()

##################### Spanish Timeline Loop

for i, video in enumerate(video_clips):
    video_clip = video_clips[i]
    wait()
    video_clip_name = video_clip.GetName()[:-4]
    wait()
    print(video_clip_name)
    media_pool.CreateEmptyTimeline("SP " + video_clip_name)
    wait()

    spanish_clip = None

    for clip in sp_clips:
        name = str(clip.GetName()).lower()
        substring_name = name[3:name.index(' w ')]
        if str(video_clip_name).lower() == substring_name:
            spanish_clip = clip

    media_pool.AppendToTimeline(spanish_clip)
    wait()
    press('home')
    wait()
    pyautogui.mouseDown(900, 850)
    pyautogui.moveTo(900, 900)
    pyautogui.mouseUp(900, 900)

    media_pool.AppendToTimeline(video_clip)
    wait()
    press('home')

    timeline = project.GetCurrentTimeline()
    timeline.SetTrackName('audio', 1, "English")
    timeline.SetTrackName('audio', 2, "Spanish")
    wait()

##################### French Timeline Loop

for i, video in enumerate(video_clips):
    video_clip = video_clips[i]
    wait()
    video_clip_name = video_clip.GetName()[:-4]
    wait()
    print(video_clip_name)
    media_pool.CreateEmptyTimeline("FR " + video_clip_name)
    wait()

    french_clip = None

    for clip in fr_clips:
        name = str(clip.GetName()).lower()
        substring_name = name[3:name.index(' w ')]
        if str(video_clip_name).lower() == substring_name:
            french_clip = clip

    media_pool.AppendToTimeline(french_clip)
    wait()
    press('home')
    wait()
    pyautogui.mouseDown(900, 850)
    pyautogui.moveTo(900, 900)
    pyautogui.mouseUp(900, 900)

    media_pool.AppendToTimeline(video_clip)
    wait()
    press('home')

    timeline = project.GetCurrentTimeline()
    timeline.SetTrackName('audio', 1, "English")
    timeline.SetTrackName('audio', 2, "French")
    wait()

##################### Testing

# video_clip = video_clips[0]
# wait()
# video_clip_name = video_clip.GetName()[:-4]
# wait()
# print(video_clip_name)
# media_pool.CreateEmptyTimeline("FR " + video_clip_name)
# wait()

# french_clip = None

# for clip in fr_clips:
#     name = str(clip.GetName()).lower()
#     substring_name = name[3:name.index(' w ')]
#     if str(video_clip_name).lower() == substring_name:
#         french_clip = clip

# media_pool.AppendToTimeline(french_clip)
# wait()
# press('home')
# wait()
# pyautogui.mouseDown(900, 850)
# pyautogui.moveTo(900, 900)
# pyautogui.mouseUp(900, 900)

# media_pool.AppendToTimeline(video_clip)
# wait()
# press('home')

# timeline = project.GetCurrentTimeline()
# timeline.SetTrackName('audio', 1, "English")
# timeline.SetTrackName('audio', 2, "French")
# wait()

# output_dir = os.path.join(os.path.expanduser("~/Desktop"), "AC French MOV Renders")
# os.makedirs(output_dir, exist_ok=True)

# project.SetRenderSettings({
#      'TargetDir': output_dir,
#      'CustomName': 'FR ' + video_clip_name,
#      'FrameRate': 59.94,
#      'VideoCodec': 'mov',
#      'AudioCodec': 'aac'
# })









