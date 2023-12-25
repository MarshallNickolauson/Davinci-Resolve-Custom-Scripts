# python 3.12
# pip install pyautogui pillow, pyperclip

import pyautogui
import pyperclip
import time
import sys

def wait():
    time.sleep(0.3)

screen_width, screen_height = pyautogui.size()

import pygetwindow as gw

def get_resolve_window():
    resolve_windows = gw.getWindowsWithTitle('DaVinci Resolve')

    if not resolve_windows:
        raise Exception("DaVinci Resolve window not found")

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
    
# Reset UI Layout
def resetUI():
    resolve.OpenPage('edit')
    if sys.platform.startswith('win'):  # Windows
        pyautogui.hotkey('ctrl', '`')
        wait()
    elif sys.platform.startswith('darwin'):  # macOS
        pyautogui.hotkey('command', '`')
        wait()

resetUI()

clip_list = media_pool.GetRootFolder().GetClipList()

for clip in clip_list:

    name = str(clip.GetName())

    if name.lower().__contains__("welcome"):
        welcome_bookend = clip
    elif name.lower().__contains__("closing"):
        closing_bookend = clip
    elif name.lower().__contains__("eng") and name.lower().__contains__("slate"):
        english_dvd_slate = clip
    elif name.lower().__contains__("frn") and name.lower().__contains__("slate"):
        french_dvd_slate = clip
    elif name.lower().__contains__("spn") and name.lower().__contains__("slate"):
        spanish_dvd_slate = clip
    elif name.lower().__contains__("mp4") or name.lower().__contains__("mov"):
        nature_video = clip

def make_opening(clip):
    resolve.OpenPage('edit')
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

    press('home')
    click(670, 800)
    get_clip_duration()

    click(670, 725)
    set_clip_duration()

    # Add to Render Queue
    project.SetRenderSettings({
        'TargetDir': 'C:\\Users\\marsh\\Desktop',
        'CustomName': clip_name[:-4],
        'FrameRate': 59.94,
        'VideoCodec': 'mp4'
    })
    project.AddRenderJob()
    click(1723, 493)
    
    resetUI()

def make_closing(clip):
    resolve.OpenPage('edit')
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
    clip_2.SetProperty("Pan", -850)
    
    press('home')
    click(670, 800)
    get_clip_duration()

    click(670, 725)
    set_clip_duration()

    # Add to Render Queue
    project.SetRenderSettings({
        'TargetDir': 'C:\\Users\\marsh\\Desktop',
        'CustomName': clip_name[:-4],
        'FrameRate': 59.94,
        'VideoCodec': 'mp4'
    })
    project.AddRenderJob()
    click(1723, 493)
    
    resetUI()

# TODO if one of these is not None? (try running without a clip): run

make_opening(welcome_bookend)
# make_opening(english_dvd_slate)
# make_opening(french_dvd_slate)
# make_opening(spanish_dvd_slate)
# make_closing(closing_bookend)

project.StartRendering()