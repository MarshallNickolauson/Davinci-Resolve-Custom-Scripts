# Davinci Resolve v20.0 Beta

# python 3.12
# pip install pyautogui opencv-python pillow

# Customize variables below before running

frames = 59.94
width = 1920
height = 1080

# Output folder is your Desktop/Davinci Script Renders

####################

import os
import time
import pyautogui
import pyperclip
import sys
import pygetwindow as gw

target_dir = os.path.join(os.path.expanduser("~/Desktop"), "Davinci Script Renders")
os.makedirs(target_dir, exist_ok=True)

script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
assets_dir = os.path.join(script_dir, "assets")

screen_width, screen_height = pyautogui.size()

def wait():
    time.sleep(0.3)
    
def get_resolve_window():
    resolve_windows = gw.getWindowsWithTitle('DaVinci Resolve')
    return resolve_windows[0]

def click(x, y):

    x_percent = float(x / screen_width)
    y_percent = float(y / screen_height)

    resolve_window = get_resolve_window()
    
    x = int(resolve_window.left + x_percent * resolve_window.width)
    y = int(resolve_window.top + y_percent * resolve_window.height)
    
    pyautogui.click(x, y, duration=0.2)
    wait()
    
def mouse_down(x, y):

    x_percent = float(x / screen_width)
    y_percent = float(y / screen_height)

    resolve_window = get_resolve_window()
    
    x = int(resolve_window.left + x_percent * resolve_window.width)
    y = int(resolve_window.top + y_percent * resolve_window.height)
    
    pyautogui.mouseDown(x, y, duration=0.2)
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
    pyautogui.hotkey('ctrl', 'v')
    wait()
    
def select_all():
    pyautogui.hotkey('ctrl', 'a')
    wait()
    
def get_clip_duration():
    pyautogui.hotkey('ctrl', 'd')
    wait()
    pyautogui.hotkey('ctrl', 'c')
    wait()

    paste()
    press('enter')
    

def set_clip_duration():
    pyautogui.hotkey('ctrl', 'd')
    wait()

    paste()
    press('enter')
    
def move_clip_to_v2():
    press('home')
    x_down, y_down = mouse_down(670, 800)
    x_up, y_up = move_mouse_to(x_down, (y_down - 50))
    mouse_up(x_up, y_up)
    
def find_image_on_screen(image_name, confidence=0.9, timeout=5):
    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            image_path = os.path.join(assets_dir, image_name)
            location = pyautogui.locateOnScreen(
                image_path, confidence=confidence, grayscale=True
            )
            if location:
                center = pyautogui.center(location)
                return center
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)

    print(f"Image not found within {timeout} seconds.")
    return None

def resetUI():
    resolve.OpenPage('edit')
    w = find_image_on_screen("workspace.png")
    click(w[0], w[1])
    r = find_image_on_screen("reset_ui_layout.png")
    click(r[0], (r[1] + 20))
    wait()

################################################
# Setup #
################################################

resolve = app.GetResolve()
projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
media_pool = project.GetMediaPool()
media_storage = resolve.GetMediaStorage()

project.SetSetting("timelineFrameRate", str(frames))
project.SetSetting("timelineResolutionWidth", str(width))
project.SetSetting("timelineResolutionHeight", str(height))

################################################
# Sequence Below #
################################################

resetUI()

clip_list = media_pool.GetRootFolder().GetClipList()

video_list = []
audio_list = []

for clip in clip_list:
    name = clip.GetName().lower()
    if name.endswith((".mp3", ".wav")):
        audio_list.append((name, clip))
    elif name.endswith((".mp4", ".mov")):
        video_list.append((name, clip))
        
video_list.sort()
audio_list.sort()

print("Video Clips (sorted):")
for name, clip in video_list:
    print(name)

print("\nAudio Clips (sorted):")
for name, clip in audio_list:
    print(name)

# media_pool.CreateEmptyTimeline("Timeline Test")
# media_pool.AppendToTimeline(video_list[0][1])
# time.sleep(2)
# g = find_image_on_screen("green_audio.png")
# mouse_down((g[0] + 100), g[1] + 60)
# move_mouse_to((g[0] + 100), (g[1] + 90))
# mouse_up((g[0] + 100), (g[1] + 90))
# press('home')
# media_pool.AppendToTimeline(audio_list[0][1])

for idx, (name, clip) in enumerate(video_list, start=1):
    media_pool.CreateEmptyTimeline(str(f"Timeline {idx}"))
    media_pool.AppendToTimeline(clip)
    time.sleep(0.5)
    g = find_image_on_screen("green_audio.png")
    mouse_down((g[0] + 100), g[1] + 60)
    move_mouse_to((g[0] + 100), (g[1] + 90))
    mouse_up((g[0] + 100), (g[1] + 90))
    press("home")
    media_pool.AppendToTimeline(audio_list[idx - 1][1])
    time.sleep(0.5)
    press("home")