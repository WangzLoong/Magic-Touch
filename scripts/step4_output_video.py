import os
import gc
import random
from concurrent.futures import ThreadPoolExecutor
from PIL import Image, ImageFilter
from moviepy.editor import (
    ImageSequenceClip,
    AudioFileClip,
    CompositeVideoClip,
    concatenate_videoclips,
    VideoFileClip,
    vfx
)
import json
from datetime import datetime
import chardet
import concurrent.futures
from tqdm import tqdm
import numpy as np

def get_config():
    config_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'config.json')

    with open(config_file, 'rb') as f:
        encoding_result = chardet.detect(raw_data := f.read())
        encoding = encoding_result['encoding']

    return json.loads(raw_data.decode(encoding))

def transform_image(img, t, x_speed, y_speed, move_on_x, move_positive):
    original_size = img.size

    crop_width = img.width * 0.8
    crop_height = img.height * 0.8
    if move_on_x:
        left = min(x_speed * t, img.width - crop_width) if move_positive else max(img.width - crop_width - x_speed * t, 0)
        upper = (img.height - crop_height) / 2
    else:
        upper = min(y_speed * t, img.height - crop_height) if move_positive else max(img.height - crop_height - y_speed * t, 0)
        left = (img.width - crop_width) / 2

    right = left + crop_width
    lower = upper + crop_height

    cropped_img = img.crop((left, upper, right, lower))
    
    return cropped_img.resize(original_size)

print("软件作者：西装革律")
print("禁止商用、倒卖、反向编译，违者必究！")
print("交流群：797579852")

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
env_dir = os.path.join(parent_dir, 'env')

env_image_magick_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'env', 'ImageMagick'))
os.environ["MAGICK_HOME"] = env_image_magick_dir
os.environ["PATH"] += os.pathsep + os.path.join(env_image_magick_dir, 'bin')

os.add_dll_directory(env_dir)

config = get_config()

image_dir = os.path.join(parent_dir, 'image')
voice_dir = os.path.join(parent_dir, 'voice')
video_dir = os.path.join(parent_dir, 'video')
temp_dir = os.path.join(parent_dir, 'temp')

os.makedirs(temp_dir, exist_ok=True)

total_files = len(os.listdir(image_dir))
fps = config['fps']
enlarge_background = config['enlarge_background']
enable_effect = config['enable_effect']
effect_type = config['effect_type']

extensions = ['.png', '.jpg', '.jpeg']
for i in tqdm(range(total_files), ncols=None, desc="正在生成视频"):
    filename = f'output_{i+1}'
    audio_filename = f'output_{i+1}'
    temp_filename = os.path.join(temp_dir, f'output_{i+1}.mp4')

    for ext in extensions:
        try:
            im = Image.open(os.path.join(image_dir, filename + ext))
            break
        except FileNotFoundError:
            continue

    audio = AudioFileClip(os.path.join(voice_dir, audio_filename + '.wav'))

    effect_type = random.choice([0, 1])

    if effect_type == 0:
        x_speed = (im.width - im.width * 0.8) / audio.duration
        y_speed = 0
        move_on_x = True
        move_positive = random.choice([True, False])
    elif effect_type == 1:
        x_speed = 0
        y_speed = (im.height - im.height * 0.8) / audio.duration
        move_on_x = False
        move_positive = random.choice([True, False])

    n_frames = int(fps * audio.duration)
    frames_foreground = [np.array(transform_image(im, t / fps, x_speed, y_speed, move_on_x, move_positive)) for t in range(n_frames)]
    img_foreground = ImageSequenceClip(frames_foreground, fps=fps)

    img_blur = im.filter(ImageFilter.GaussianBlur(radius=30))
    if enlarge_background:
        img_blur = img_blur.resize((int(im.width * 1.1), int(im.height * 1.1)), Image.ANTIALIAS)

    frames_background = [np.array(img_blur)] * n_frames
    img_background = ImageSequenceClip(frames_background, fps=fps)

    img_foreground = img_foreground.set_audio(audio)

    final_clip = CompositeVideoClip([img_background.set_position("center"), img_foreground.set_position("center")], size=img_blur.size)

    final_clip = {
        'fade': final_clip.fadein(1).fadeout(1),
        'slide': final_clip.crossfadein(1).crossfadeout(1),
        'rotate': final_clip.rotate(lambda t: 360*t/10),
        'scroll': final_clip.fx(vfx.scroll, y_speed=50),
        'flip_horizontal': final_clip.fx(vfx.mirror_x),
        'flip_vertical': final_clip.fx(vfx.mirror_y)
    }.get(effect_type, final_clip)


    final_clip.write_videofile(temp_filename)
    gc.collect()

temp_filenames = [os.path.join(temp_dir, f'output_{i+1}.mp4') for i in range(total_files)]
temp_filenames.sort(key=lambda x: int(x.split('_')[-1].split('.')[0]))
final_video = concatenate_videoclips([VideoFileClip(filename) for filename in temp_filenames], method="compose")
final_video.write_videofile(os.path.join(video_dir, f'output_{datetime.now().strftime("%Y%m%d%H%M%S")}.mp4'))
