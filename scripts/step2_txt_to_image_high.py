import os
import json
import base64
import requests
import openpyxl
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
import chardet
from tqdm import tqdm

session = requests.Session()
retries = Retry(total=5, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(max_retries=retries)
session.mount('http://', adapter)
session.mount('https://', adapter)

current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def post(url, data):
    return session.post(url, json=data)

def save_img(b64_image, path):
    with open(path, "wb") as file:
        file.write(base64.b64decode(b64_image))

def get_prompts(path):
    prompts_file = os.path.join(current_dir, path)
    wb = openpyxl.load_workbook(prompts_file)
    sheet = wb.active
    prompts = [cell.value for cell in sheet['C'] if cell.value]
    wb.close()
    return prompts
    
def get_cloud_address():
    config_file = os.path.join(current_dir, 'config.json')
    cloud_address = None
    more_details = None
    data = None

    if os.path.exists(config_file):
        with open(config_file, 'rb') as f:
            raw_data = f.read()
            detected_encoding = chardet.detect(raw_data)['encoding']
        with open(config_file, 'r', encoding=detected_encoding) as f:
            config = json.load(f)
            cloud_address = config.get('cloud_address')
            more_details = config.get('more_details')
            data = config.get('data')

    if not cloud_address:
        cloud_address = None
    if more_details is None:
        more_details = {}
    if data is None:
        data = {}

    return cloud_address, more_details, data

def run_program(cloud_address, prompts_to_redraw=None, data=None):
    fixed_data = {
        "alwayson_scripts": {
            "ADetailer": {
                "args": [
                    {"ad_model": "face_yolov8n.pt"},
                    {"ad_model": "hand_yolov8n.pt"}
                ]
            }
        }
    }

    if data:
        fixed_data.update(data)

    url = cloud_address.rstrip('/') + '/sdapi/v1/txt2img' if cloud_address else ""

    if not url:
        print("未提供云端Stable Diffusion地址")
        return

    prompts = get_prompts(os.path.join('txt', 'txt.xlsx'))

    image_dir = os.path.join(current_dir, 'image')
    os.makedirs(image_dir, exist_ok=True)

    prompts_to_process = list(enumerate(prompts))
    if prompts_to_redraw is not None:
        prompts_to_process = [(i, prompt) for i, prompt in prompts_to_process if i in prompts_to_redraw]

    total_images = len(prompts_to_process)
    existing_files = set(os.listdir(image_dir))

    for i, prompt_b in tqdm(prompts_to_process, desc='绘图进度', unit='image'):
        prompt = f"{prompt_b},{more_details}"

        output_file = f'output_{i+1}.png'
        if output_file in existing_files and prompts_to_redraw is None:
            continue
        fixed_data["prompt"] = prompt
        response = post(url, fixed_data)
        if response.status_code == 200:
            save_img(response.json()['images'][0], os.path.join(image_dir, output_file))
            temp_dir = os.path.join(current_dir, 'temp')
            os.makedirs(temp_dir, exist_ok=True)
            with open(os.path.join(temp_dir, 'params.json'), 'a') as f:
                json.dump({output_file: fixed_data}, f)
                f.write('\n')
        else:
            print(f'错误：{response.status_code}')

        if not prompt_b:
            break

if __name__ == '__main__':
    print("软件作者：西装革律")
    print("禁止商用、倒卖、反向编译，违者必究！")
    print("交流群：797579852")

    cloud_address, more_details, data = get_cloud_address()

    if cloud_address is None:
        cloud_address = "http://127.0.0.1:7860"
        print("使用本地Stable Diffusion")
    else:
        print("使用云端Stable Diffusion")

    print("Stable Diffusion正在绘图，请稍后...")
    run_program(cloud_address, data=data)
    print("Stable Diffusion绘图完成，请检查图片，不想踩缝纫机就赶紧把小黄图删了~")

    while True:
        user_input = input("请输入需要重绘的图片对应的数字（多个数字用空格隔开，输入N退出程序）: ")
        if user_input == "N":
            break

        file_numbers_to_redraw = []
        for s in user_input.split():
            try:
                file_number = int(s.strip()) - 1
                file_name = f"output_{file_number+1}.png"
                file_path = os.path.join('image', file_name)

                if os.path.exists(os.path.join(current_dir, file_path)):
                    file_numbers_to_redraw.append(file_number)
                    os.remove(os.path.join(current_dir, file_path))
                    print(f"重绘图片: {file_name}")
                else:
                    print(f"无效图片: {file_name}")
            except ValueError:
                print(f"无效输入: {s.strip()}，跳过")

        if file_numbers_to_redraw:
            print("Stable Diffusion正在重绘，请稍后...")
            run_program(cloud_address, prompts_to_redraw=file_numbers_to_redraw, data=data)
            print("Stable Diffusion重绘完成，请检查图片，你确定还不删小黄图么？")
        else:
            print("没有需要重绘的图片")
