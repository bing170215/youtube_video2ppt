#@Author:   Casserole fish
#@Time:    2022/4/21 20:39
from __future__ import unicode_literals
import yt_dlp
import os
from save_slide import save_slide

print("      _____________________________________________________   ")
print("     /                                                     \‎  ")
print("     |    _____________________________________________     | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |      Welcome To Video to PPT Command        |    | ")
print("     |   |                  Utility                    |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("     |   |                                             |    | ")
print("      \_____________________________________________________/ ")
print("           \_______________________________________/          ")
print("         _______________________________________________ ")
print("      _-'    .-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.  --- '-_ ")
print("    _-'.-.-. .---.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.--.  .-.-.'-_ ")
print(" _-'.-.-.-. .---.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.- __ . .-.-.-.'-_ ")
print("_-'.-.-.-.-. .-----.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-----. .-.-.-.-.'-_ ")
print("_-'.-.-.-.-.-. .---.-. .-----------------------------. .-.---. .---.-.-.-.'-_ ")



input_exit = False
def isExit(input_str):
    if input_str.lower()=='exit':
        return True

vid_url = input('请输入需要转换的video链接：')
if isExit(vid_url):
    exit()
ppt_save_dir = input('请输入ppt保存目录：')
if isExit(ppt_save_dir):
    exit()
ppt_name = input('请输入ppt名称：')
if isExit(ppt_name):
    exit()
if ppt_name.split('.')[1] not in ['ppt','pptx']:
    raise Exception('请输入正确的ppt名称，以.ppt或.pptx结尾！！！')
start = input('请输入需要从视频多少秒开始转换ppt：')
if isExit(start):
    exit()
try:
    int(start) and int(start) > 0
except ValueError:
    raise Exception('请输入正确的开始秒数，开始秒数必须为整数！！！')
start = int(start)




if not os.path.exists(ppt_save_dir):
    os.makedirs(ppt_save_dir)
if os.path.exists(os.path.join(ppt_save_dir,ppt_name)):
    raise Exception ('该ppt名字已存在，请换个名字！！！')
mp4_name = ppt_name.split('.')[0]+'.mp4'
vid_save_path = os.path.join(ppt_save_dir,mp4_name)
ppt_save_path = os.path.join(ppt_save_dir,ppt_name)

class MyLogger(object):
    def debug(self, msg):
        pass

    def warning(self, msg):
        pass

    def error(self, msg):
        print(msg)

def finished_hook(d):
    if d['status'] == 'finished':
        print('\n视频下载完成！！！')
        # 接下来处理转ppt
        print('开始转换ppt！！！')
        save_slide(start=start, dir_path=ppt_save_dir, vid_path=vid_save_path, ppt_path=ppt_save_path)

def downloading_hook(d):
    if d['status']=='downloading':
        print("\r预计还需"+str(d['eta'])+"秒完成下载！！！".ljust(10), end='')


ydl_opts={
    'logger': MyLogger(),
    'progress_hooks': [downloading_hook, finished_hook],
    'outtmpl':vid_save_path
}

with yt_dlp.YoutubeDL(ydl_opts) as ydl:
    ydl.download([vid_url])