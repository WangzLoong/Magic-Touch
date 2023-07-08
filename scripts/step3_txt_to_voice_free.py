import os
import json
import openpyxl
import asyncio
from tqdm import tqdm
import argparse
import aiofiles
import edge_tts
import edge_tts.exceptions
import chardet

def get_encoding(file_path):
    with open(file_path, 'rb') as f:
        return chardet.detect(f.read())['encoding']

class SpeechProvider:
    def __init__(self, config_file):
        encoding = get_encoding(config_file)
        with open(config_file, 'r', encoding=encoding) as f:
            self._config = json.load(f)

    async def async_get_tts_audio(self, message, language):
        voice = self._config.get('voice')
        if voice is None:
            return None, None

        wav = b''
        service_data = {
            "voice": voice,
            "rate": self._config.get('rate'),
            "volume": self._config.get('volume')
        }

        tts = edge_tts.Communicate(message, **service_data)
        try:
            async for chunk in tts.stream():
                if chunk["type"] == "audio":
                    wav += chunk["data"]
        except edge_tts.exceptions.NoAudioReceived:
            return None, None

        return 'wav', wav

async def convert_text_to_audio(text, language, output_path, row_index, config_file):
    if not text:
        return False

    provider = SpeechProvider(config_file)
    attempt = 0  
    wait_time = 1  

    while True:
        try:
            audio_format, audio_data = await provider.async_get_tts_audio(text, language)
            if audio_data is not None:
                wav_file_path = os.path.join(output_path, f"output_{row_index}.wav")

                async with aiofiles.open(wav_file_path, "wb") as f:
                    await f.write(audio_data)

                if not os.path.exists(wav_file_path) or os.path.getsize(wav_file_path) == 0:
                    raise Exception(f"音频文件生成失败或文件大小为0: {wav_file_path}")

                return True
        except Exception as e:
            print(f"尝试 {attempt + 1} 失败，原因：{str(e)}。将在 {wait_time} 秒后重试。")
            await asyncio.sleep(wait_time)  
            attempt += 1  
            wait_time *= 2   
        except edge_tts.exceptions.RateLimitException:  
            print("超过速率限制。将在 60 秒后重试。")
            await asyncio.sleep(60)
    return False

async def process_text_files(input_file, output_dir, language, config_file):
    print("软件作者：西装革律")
    print("禁止倒卖，违者必究！")
    print("交流群：797579852")  
    
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    column = sheet["D"]
    tasks = []
    row_index = 1

    for cell in column:
        text = cell.value
        if text:
            tasks.append(asyncio.create_task(convert_text_to_audio(text, language, output_dir, row_index, config_file)))
            row_index += 1

    progress_bar = tqdm(desc="正在生成配音音频", total=len(tasks), unit="files")

    while tasks:
        done, tasks = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)
        for task in done:
            try:
                if task.result():
                    progress_bar.update(1)
            except edge_tts.exceptions.EdgeTTSException as e:
                progress_bar.write(f"发生错误：{str(e)}")
    progress_bar.close()

parser = argparse.ArgumentParser(description='Text to Speech Converter')
script_dir = os.path.dirname(os.path.realpath(__file__))

default_input_file = os.path.join(script_dir, "..", "txt", "txt.xlsx")
default_output_dir = os.path.join(script_dir, "..", "voice")

default_config_file = os.path.join(script_dir, "..", "config.json")

parser.add_argument('--input_file', type=str, default=default_input_file, help='输入文本文件的路径')
parser.add_argument('--output_dir', type=str, default=default_output_dir, help='输出目录的路径')
parser.add_argument('--config_file', type=str, default=default_config_file, help='配置文件的路径')

parser.add_argument('--language', type=str, default="zh-CN", help='文本的语言')

args = parser.parse_args()
provider = SpeechProvider(args.config_file)

asyncio.run(process_text_files(args.input_file, args.output_dir, args.language, args.config_file))
