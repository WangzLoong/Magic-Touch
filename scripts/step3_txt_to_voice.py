import os
import openpyxl
import asyncio
from tqdm.asyncio import tqdm as async_tqdm
import argparse
import azure.cognitiveservices.speech as speechsdk
from azure.cognitiveservices.speech import SpeechConfig, SpeechSynthesizer, AudioDataStream, ResultReason
from io import BytesIO
import html
import json
import chardet


def load_config():
    script_directory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_file = os.path.join(script_directory, 'config.json')

    with open(config_file, 'rb') as f:
        raw_data = f.read()
        encoding_result = chardet.detect(raw_data)
        encoding = encoding_result['encoding']

    with open(config_file, 'r', encoding=encoding) as f:
        config = json.load(f)

    return config

config = load_config()

subscription = config.get('subscription')
region = config.get('region')
voice_name = config.get('voice_name')
style = config.get('style')
role = config.get('role')
prosody_rate = config.get('prosody_rate')
prosody_pitch = config.get('prosody_pitch')
prosody_volume = config.get('prosody_volume')
emphasis_level = config.get('emphasis_level')
style_degree = config.get('style_degree')

class SpeechProvider:
    def __init__(self, config=None):
        self._config = config or {}

    async def get_tts_audio(self, message, language, index):
        while True:
            try:
                speech_config = SpeechConfig(subscription=subscription, region=region)
                speech_config.speech_synthesis_voice_name = voice_name
                synthesizer = SpeechSynthesizer(speech_config=speech_config, audio_config=None)
                escaped_message = html.escape(message)
                ssml_text = f"""
                <speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='http://www.w3.org/2001/mstts' xml:lang='{language}'>
                  <voice name='{voice_name}'>
                    <mstts:express-as style='{style}' role='{role}' styledegree='{style_degree}'>
                      <prosody rate='{prosody_rate}' pitch='{prosody_pitch}' volume='{prosody_volume}'>
                        {escaped_message}
                      </prosody>
                    </mstts:express-as>
                  </voice>
                </speak>
                """
                loop = asyncio.get_running_loop()
                result = await loop.run_in_executor(None, lambda: synthesizer.speak_ssml_async(ssml_text).get())
                if result.reason == ResultReason.SynthesizingAudioCompleted:
                    audio_data = BytesIO(result.audio_data)
                    return {"index": index, "audio_data": audio_data, "error": None}
                elif result.reason == ResultReason.Canceled:
                    cancellation_details = speechsdk.SpeechSynthesisCancellationDetails(result)
                    print(f"序号 {index} 的语音合成出错，错误信息：{str(cancellation_details.reason)} {str(cancellation_details.error_details)}，正在进行下一次尝试...")
            except Exception as e:
                print(f"序号 {index} 的语音合成出错，错误信息：{str(e)}，正在进行下一次尝试...")

async def process_text_files(input_file, output_dir, language):
    print("软件作者：西装革律")
    print("禁止倒卖，违者必究！")
    print("交流群：797579852")

    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    column = sheet["D"]
    provider = SpeechProvider()
    tasks = [provider.get_tts_audio(cell.value, language, i) for i, cell in enumerate(column, 1)]
    results = []
    for f in async_tqdm(asyncio.as_completed(tasks), total=len(tasks), desc="正在合成配音"):
        result = await f
        audio_data = result['audio_data']
        output_path = os.path.join(script_directory, output_dir, f"output_{result['index']}.wav")
        with open(output_path, 'wb') as f:
            f.write(audio_data.getbuffer())
        results.append(result)
    return results

parser = argparse.ArgumentParser(description='文本转语音转换器')
parser.add_argument('--input_file', type=str, default="txt/txt.xlsx", help='输入文本文件的路径')
parser.add_argument('--output_dir', type=str, default="voice", help='输出目录的路径')
parser.add_argument('--language', type=str, default="zh-CN", help='文本的语言')

args = parser.parse_args()

script_directory = os.path.dirname(os.path.abspath(__file__))

args.input_file = os.path.join(script_directory, '..', args.input_file.replace('/', '\\'))
args.output_dir = os.path.join(script_directory, '..', args.output_dir)

asyncio.run(process_text_files(args.input_file, args.output_dir, args.language))
