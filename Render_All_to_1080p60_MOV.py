# Davinci Resolve v20.0 Beta

# python 3.12
# pip install pyautogui opencv-python pillow

# Customize variables below before running:

video_codec = 'mov' # mov or mp4 only

import os

target_dir = os.path.join(os.path.expanduser("~/Desktop"), "Davinci Script Renders")   #adjust
os.makedirs(target_dir, exist_ok=True)

################################################
# Setup #
################################################

resolve = app.GetResolve()

projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
media_pool = project.GetMediaPool()
media_storage = resolve.GetMediaStorage()

################################################
# Sequence Below #
################################################

# print(project.GetRenderCodecs('mp4'))
# print(project.GetSetting("timelineFrameRate"))

resolve.OpenPage('deliver')
timeline_count = project.GetTimelineCount()

for i in range(timeline_count):
    timeline = project.GetTimelineByIndex(i + 1)
    project.SetCurrentTimeline(timeline)
    
    timeline = project.GetCurrentTimeline()
    timeline_name = timeline.GetName()
    
    project.SetRenderSettings({
        'TargetDir': target_dir,
        'CustomName': timeline_name,
        'FrameRate': project.GetSetting("timelineFrameRate"),
        'VideoCodec': video_codec,
    })
    
    if video_codec == 'mov':
        project.SetCurrentRenderFormatAndCodec('mov', 'Apple ProRes 422')
    elif video_codec == 'mp4':
        project.SetCurrentRenderFormatAndCodec('mp4', 'H.264')
    project.AddRenderJob()

project.StartRendering()