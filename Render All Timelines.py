# If you want to use this script you need to customize the following:

frames = 59.94  #adjust
width = 1280    #adjust
height = 720    #adjust

import os

target_dir = os.path.join(os.path.expanduser("~/Desktop"), "Rendered MOVs")   #adjust
os.makedirs(target_dir, exist_ok=True)

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

resolve.OpenPage('deliver')
timeline_count = project.GetTimelineCount()

for i in range(timeline_count):
    timeline = project.GetTimelineByIndex(i + 1)
    project.SetCurrentTimeline(timeline)
    
    timeline = project.GetCurrentTimeline()
    timeline_name = timeline.GetName()
    print(timeline_name)
    
    project.SetRenderSettings({
        'TargetDir': target_dir,
        'CustomName': timeline_name,
        'FrameRate': 59.94,
        'VideoCodec': 'mov',
    })
    project.SetCurrentRenderFormatAndCodec('mov', 'DNxHD720p220')
    project.AddRenderJob()

project.StartRendering()