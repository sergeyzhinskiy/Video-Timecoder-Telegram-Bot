# -*- coding: utf-8 -*-
import asyncio
import aiohttp
import configparser
import logging
import os
import re
import json
import subprocess
import tempfile
import sys
from docx import Document
import openpyxl
import zipfile
from telethon import TelegramClient, events
from deepgram import Deepgram
import yt_dlp
import time
from tenacity import retry, stop_after_attempt, wait_exponential
from urllib.parse import urlparse
from typing import List, Tuple

# Указание путей к FFmpeg 
ffmpeg_path = "c:/bot/factory/choc/tools/chocolateyInstall/lib/ffmpeg/tools/ffmpeg/bin/ffmpeg.exe"
ffprobe_path = "c:/bot/factory/choc/tools/chocolateyInstall/lib/ffmpeg/tools/ffmpeg/bin/ffprobe.exe"

# Настройка логгирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("telegram_bot.log", encoding='utf-8')
    ]
)
logger = logging.getLogger("TelegramBot")

# Загрузка конфигурации
try:
    config = configparser.ConfigParser()
    if not config.read("config.ini"):
        raise FileNotFoundError("Файл config.ini не найден")
    
    api_id = config['Telegram']['api_id']
    api_hash = config['Telegram']['api_hash']
    bot_token = config['Telegram']['bot_token']
    yandex_api_key = config['Yandex']['yandexgpt_api']
    
    # Загрузка всех Deepgram API ключей
    deepgram_api_keys = []
    for key in config['Deepgram']:
        if key.startswith('api_key_'):
            deepgram_api_keys.append(config['Deepgram'][key])
    
    if not deepgram_api_keys:
        raise ValueError("Не найдены Deepgram API ключи в конфигурации")
    
    logger.info(f"Загружено {len(deepgram_api_keys)} Deepgram API ключей")
    logger.info("Конфигурация успешно загружена")
except Exception as e:
    logger.critical(f"Ошибка загрузки конфигурации: {str(e)}")
    sys.exit(1)

# Инициализация клиента Telegram
client = TelegramClient("bot", int(api_id), api_hash).start(bot_token=bot_token)

# Менеджер для работы с Deepgram API ключами
class DeepgramKeyManager:
    def __init__(self, api_keys: List[str]):
        self.api_keys = api_keys
        self.current_key_index = 0
        self.failed_keys = set()
        self.lock = asyncio.Lock()
    
    def get_current_key(self) -> str:
        return self.api_keys[self.current_key_index]
    
    async def rotate_key(self, failed_key: str = None):
        async with self.lock:
            if failed_key:
                self.failed_keys.add(failed_key)
            
            # Ищем следующий рабочий ключ
            original_index = self.current_key_index
            while True:
                self.current_key_index = (self.current_key_index + 1) % len(self.api_keys)
                if self.api_keys[self.current_key_index] not in self.failed_keys:
                    break
                if self.current_key_index == original_index:
                    # Все ключи нерабочие
                    raise Exception("Все Deepgram API ключи исчерпаны или нерабочие")
            
            logger.warning(f"Переключение на Deepgram API ключ #{self.current_key_index + 1}")
    
    def get_available_keys_count(self) -> int:
        return len(self.api_keys) - len(self.failed_keys)

# Инициализация менеджера ключей Deepgram
deepgram_key_manager = DeepgramKeyManager(deepgram_api_keys)

# Настройки YandexGPT
MAX_TOKENS = 2000
SYSTEM_PROMPT = """
Ты – сервис для создания тайм-кодов. Получаешь текстовую расшифровка видео.
Сделай список тайм-кодов с краткими описаниями содержания.
Формат: 00:00 – Введение
             # основная мысль 1
             # основная мысль 2
             # основная мысль 3
        01:25 – Основная тема
             # основная мысль 1
             # основная мысль 2
             # основная мысль 3
		"""

# ---------- Вспомогательные функции ----------

def is_valid_url(url: str) -> bool:
    """Проверяет, является ли строка валидным URL"""
    try:
        result = urlparse(url)
        return all([result.scheme, result.netloc])
    except:
        return False

async def download_audio(url: str, user_id: int) -> str:
    """Скачивает аудиодорожку из видео с оптимизацией для длинных видео"""
    try:
        # Создаем временную директорию для аудио
        audio_dir = os.path.join("user_data", str(user_id), "audio")
        os.makedirs(audio_dir, exist_ok=True)
        
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': os.path.join(audio_dir, '%(title)s.%(ext)s'),
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '128',  # Понижаем качество для уменьшения размера
            }],
            'ffmpeg_location': ffmpeg_path,
            'quiet': True,
            'no_warnings': True,
            'extractaudio': True,
            'audioformat': 'mp3',
            'noplaylist': True,
            'socket_timeout': 30,
            'retries': 3,
        }
        
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            downloaded_file_path = ydl.prepare_filename(info)
            base, _ = os.path.splitext(downloaded_file_path)
            audio_file_path = base + '.mp3'
            
            # Проверяем размер файла
            file_size = os.path.getsize(audio_file_path) / (1024 * 1024)  # в МБ
            logger.info(f"[{user_id}] Размер аудиофайла: {file_size:.2f} МБ")
            
            return audio_file_path
            
    except Exception as e:
        logger.error(f"Ошибка скачивания аудио {url}: {e}")
        return None

async def transcribe_audio(audio_path: str) -> str:
    """Преобразует аудио в текст с помощью Deepgram через асинхронное API"""
    max_retries = deepgram_key_manager.get_available_keys_count()
    
    for retry_count in range(max_retries):
        current_key = deepgram_key_manager.get_current_key()
        
        try:
            # Читаем аудиофайл
            with open(audio_path, 'rb') as audio:
                audio_data = audio.read()
            
            # Используем асинхронное API для длинных файлов
            url = "https://api.deepgram.com/v1/listen/async"
            params = {
                'model': 'whisper-large',
                'language': 'ru',
                'punctuate': 'true',
                'paragraphs': 'true',
                'diarize': 'true'
            }
            
            headers = {
                'Authorization': f'Token {current_key}',
                'Content-Type': 'audio/mpeg'  # Меняем на audio/mpeg для MP3
            }
            
            # Шаг 1: Загрузка файла и получение request_id
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=300)) as session:
                async with session.post(
                    url,
                    params=params,
                    headers=headers,
                    data=audio_data
                ) as response:
                    
                    if response.status in [401, 403]:
                        error_text = await response.text()
                        logger.error(f"Ошибка аутентификации Deepgram API: {response.status} - {error_text}")
                        await deepgram_key_manager.rotate_key(current_key)
                        continue
                    elif response.status in [402, 429]:
                        error_text = await response.text()
                        logger.error(f"Ошибка баланса/лимита Deepgram API: {response.status} - {error_text}")
                        await deepgram_key_manager.rotate_key(current_key)
                        continue
                    elif response.status != 200:
                        error_text = await response.text()
                        logger.error(f"Ошибка Deepgram API при загрузке: {response.status} - {error_text}")
                        return None
                    
                    result = await response.json()
                    request_id = result.get('request_id')
                    
                    if not request_id:
                        logger.error("Не получили request_id от Deepgram")
                        return None
            
            # Шаг 2: Ожидание обработки с экспоненциальной задержкой
            status_url = f"https://api.deepgram.com/v1/listen/async/{request_id}"
            transcript_url = f"{status_url}/transcript"
            
            max_wait_time = 3600  # Максимальное время ожидания 1 час
            start_time = time.time()
            wait_time = 5  # Начальная задержка
            
            while time.time() - start_time < max_wait_time:
                await asyncio.sleep(wait_time)
                
                async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=60)) as session:
                    async with session.get(
                        status_url,
                        headers={'Authorization': f'Token {current_key}'}
                    ) as status_response:
                        
                        if status_response.status != 200:
                            await asyncio.sleep(wait_time)
                            wait_time = min(wait_time * 1.5, 30)  # Экспоненциальная задержка до 30 сек
                            continue
                        
                        status_result = await status_response.json()
                        
                        if status_result.get('status') == 'finished':
                            # Получаем готовую транскрипцию
                            async with session.get(
                                transcript_url,
                                headers={'Authorization': f'Token {current_key}'}
                            ) as transcript_response:
                                
                                if transcript_response.status == 200:
                                    transcript_result = await transcript_response.json()
                                    if 'results' in transcript_result and 'channels' in transcript_result['results']:
                                        transcript = transcript_result['results']['channels'][0]['alternatives'][0]['transcript']
                                        return transcript
                                else:
                                    logger.error(f"Ошибка получения транскрипции: {transcript_response.status}")
                                    return None
                        
                        elif status_result.get('status') == 'error':
                            logger.error(f"Ошибка обработки Deepgram: {status_result.get('error', 'Unknown error')}")
                            return None
                        
                        # Увеличиваем задержку для следующей проверки
                        wait_time = min(wait_time * 1.5, 30)
            
            logger.error(f"Превышено время ожидания обработки для {audio_path}")
            return None
                        
        except asyncio.TimeoutError:
            logger.error("Таймаут при обращении к Deepgram API")
            return None
        except Exception as e:
            logger.error(f"Ошибка транскрибации через Deepgram: {e}")
            return None
    
    logger.error("Все попытки транскрибации с разными ключами провалились")
    return None

async def split_audio_if_needed(audio_path: str, max_duration: int = 1800) -> List[str]:
    """Разделяет длинные аудиофайлы на части по 30 минут"""
    try:
        # Получаем длительность аудио
        cmd = [
            ffprobe_path,
            '-v', 'error',
            '-show_entries', 'format=duration',
            '-of', 'default=noprint_wrappers=1:nokey=1',
            audio_path
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        duration = float(result.stdout.strip())
        
        logger.info(f"Длительность аудио: {duration} секунд")
        
        # Если аудио короче max_duration, возвращаем как есть
        if duration <= max_duration:
            return [audio_path]
        
        # Разделяем аудио на части
        parts = []
        temp_dir = os.path.join(os.path.dirname(audio_path), "parts")
        os.makedirs(temp_dir, exist_ok=True)
        
        base_name = os.path.splitext(os.path.basename(audio_path))[0]
        
        for i, start_time in enumerate(range(0, int(duration), max_duration)):
            part_path = os.path.join(temp_dir, f"{base_name}_part_{i+1}.mp3")
            
            cmd = [
                ffmpeg_path,
                '-i', audio_path,
                '-ss', str(start_time),
                '-t', str(max_duration),
                '-c', 'copy',
                '-y',
                part_path
            ]
            
            subprocess.run(cmd, check=True, capture_output=True)
            parts.append(part_path)
            logger.info(f"Создана часть {i+1}: {part_path}")
        
        return parts
        
    except Exception as e:
        logger.error(f"Ошибка разделения аудио: {e}")
        return [audio_path]

async def transcribe_long_audio(audio_path: str) -> str:
    """Транскрибация длинных аудио через разделение на части"""
    parts = await split_audio_if_needed(audio_path)
    full_transcript = ""
    
    for i, part_path in enumerate(parts):
        logger.info(f"Обработка части {i+1}/{len(parts)}")
        
        # Используем обычную транскрибацию для каждой части
        part_transcript = await transcribe_audio(part_path)
        
        if part_transcript:
            full_transcript += f"\n\n[Часть {i+1}]\n{part_transcript}"
        
        # Очищаем временные файлы частей (кроме оригинального)
        if part_path != audio_path:
            try:
                os.remove(part_path)
            except:
                pass
        
        # Задержка между частями
        await asyncio.sleep(1)
    
    # Очищаем папку с частями если она пустая
    temp_dir = os.path.join(os.path.dirname(audio_path), "parts")
    try:
        if os.path.exists(temp_dir) and not os.listdir(temp_dir):
            os.rmdir(temp_dir)
    except:
        pass
    
    return full_transcript if full_transcript else None

async def generate_timestamps_with_yandexgpt(transcript: str, video_url: str) -> str:
    """Генерация тайм-кодов через YandexGPT"""
    headers = {
        'Authorization': f'Api-Key {yandex_api_key}',
        'Content-Type': 'application/json'
    }
    
    prompt = {
        "modelUri": "gpt://ваш-номер-папки/yandexgpt/latest",
        "completionOptions": {
            'stream': False,
            'temperature': 0.3,
            'maxTokens': MAX_TOKENS
        },
        "messages": [
            {"role": "system", "text": SYSTEM_PROMPT},
            {"role": "user", "text": f"Вот текстовая расшифровка видео: {transcript}. Создай тайм-коды для этого видео."}
        ]
    }
    
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(
                "https://llm.api.cloud.yandex.net/foundationModels/v1/completion",
                headers=headers,
                json=prompt,
                timeout=60
            ) as response:
                response.raise_for_status()
                result = await response.json()
                return result['result']['alternatives'][0]['message']['text']
                
    except Exception as e:
        logger.error(f"Ошибка YandexGPT: {e}")
        return None

def sanitize_filename(name: str) -> str:
    """Очищает строку для использования в имени файла"""
    return re.sub(r'[\\/*?:"<>|]', "", name)[:80]

async def process_single_url(url: str, user_id: int) -> str:
    """Обрабатывает одиночную ссылку и возвращает текст с тайм-кодами"""
    try:
        logger.info(f"[{user_id}] Обработка одиночной ссылки: {url}")
        
        # Шаг 1: Скачиваем аудио
        audio_path = await download_audio(url, user_id)
        if not audio_path:
            logger.error(f"[{user_id}] Не удалось скачать аудио: {url}")
            return None

        # Шаг 2: Преобразуем аудио в текст (с поддержкой длинных файлов)
        transcript = await transcribe_long_audio(audio_path)  # Используем новую функцию
        if not transcript:
            logger.error(f"[{user_id}] Не удалось преобразовать аудио в текст: {url}")
            # Удаляем временный аудиофайл
            try:
                os.remove(audio_path)
            except:
                pass
            return None

        # Шаг 3: Генерируем тайм-коды с помощью YandexGPT
        timestamps = await generate_timestamps_with_yandexgpt(transcript, url)
        if not timestamps:
            logger.error(f"[{user_id}] Не удалось сгенерировать тайм-коды: {url}")
            # Удаляем временный аудиофайл
            try:
                os.remove(audio_path)
            except:
                pass
            return None

        # Удаляем временный аудиофайл
        try:
            os.remove(audio_path)
        except:
            pass

        return timestamps
        
    except Exception as e:
        logger.error(f"[{user_id}] Ошибка обработки ссылки {url}: {e}")
        return None

async def process_excel(user_id: int, file_path: str, output_dir: str) -> Tuple[int, int]:
    """Обработка Excel-файла со ссылками"""
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        os.makedirs(output_dir, exist_ok=True)
        
        processed_count = 0
        failed_count = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
                
            url = str(row[0]).strip()
            
            # Пропускаем невалидные URL
            if not is_valid_url(url):
                continue
                
            logger.info(f"[{user_id}] Обработка: {url}")

            # Обрабатываем ссылку
            timestamps = await process_single_url(url, user_id)
            if not timestamps:
                logger.error(f"[{user_id}] Не удалось обработать: {url}")
                failed_count += 1
                continue

            # Создаем документ
            title = f"Видео_{processed_count + 1}"
            doc = Document()
            doc.add_heading(f"Тайм-коды для: {title}", level=1)
            doc.add_paragraph(f"Ссылка на видео: {url}")
            doc.add_paragraph()
            
            for line in timestamps.splitlines():
                if line.strip():
                    doc.add_paragraph(line)

            # Сохраняем файл
            clean_title = sanitize_filename(title)
            file_name = os.path.join(output_dir, f"{clean_title}.docx")
            doc.save(file_name)
            logger.info(f"[{user_id}] Сохранён: {file_name}")
            processed_count += 1
            
            # Задержка между обработкой видео
            await asyncio.sleep(2)

        return processed_count, failed_count
    except Exception as e:
        logger.error(f"[{user_id}] Ошибка обработки Excel: {e}")
        return 0, 0

def make_zip(output_dir: str, zip_path: str):
    """Создает ZIP-архив с документами"""
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                if file.endswith('.docx'):
                    file_path = os.path.join(root, file)
                    arcname = os.path.basename(file_path)
                    zf.write(file_path, arcname)
    return zip_path

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
async def download_audio_with_retry(url: str, user_id: int) -> str:
    return await download_audio(url, user_id)

# ---------- Обработчик сообщений ----------

@client.on(events.NewMessage(pattern='/start'))
async def start_handler(event):
    """Обработчик команды /start"""
    await event.reply(
        "Привет! Я бот для создания тайм-кодов видео.\n\n"
        "Отправьте мне:\n"
        "1. Excel-файл со списком ссылок на видео (ссылками в первом столбце), затем используйте команду /timecod для обработки\n"
        "2. Прямую ссылку на видео для мгновенной обработки\n\n"
        "Я скачаю аудио, преобразую его в текст и сгенерирую тайм-коды с помощью YandexGPT."
    )

@client.on(events.NewMessage(pattern='/status'))
async def status_handler(event):
    """Обработчик команды /status - показывает статус API ключей"""
    available_keys = deepgram_key_manager.get_available_keys_count()
    total_keys = len(deepgram_key_manager.api_keys)
    current_key_index = deepgram_key_manager.current_key_index
    
    status_message = (
        f"📊 Статус API ключей:\n"
        f"• Всего ключей: {total_keys}\n"
        f"• Доступно ключей: {available_keys}\n"
        f"• Текущий ключ: #{current_key_index + 1}\n"
        f"• Заблокировано ключей: {total_keys - available_keys}"
    )
    
    await event.reply(status_message)

@client.on(events.NewMessage(incoming=True))
async def handler(event):
    """Основной обработчик сообщений"""
    msg = event.message
    user_id = msg.sender_id
    text = msg.text or ""

    # Приём Excel-файлов
    if msg.file and msg.file.name and msg.file.name.endswith(('.xlsx', '.xls')):
        user_dir = os.path.join("user_data", str(user_id), "in")
        os.makedirs(user_dir, exist_ok=True)
        
        file_name = msg.file.name
        file_path = os.path.join(user_dir, file_name)
        
        await msg.download_media(file_path)
        await event.reply(
            f"📥 Файл сохранён: `{file_name}`.\n"
            f"Теперь отправьте команду `/timecod` для обработки.",
            parse_mode="md"
        )
        return

    # Обработка одиночных URL
    if text and is_valid_url(text.strip()):
        url = text.strip()
        await event.reply(f"🔗 Начинаю обработку ссылки: {url}\n\nЭто может занять несколько минут...")
        
        # Обрабатываем ссылку
        timestamps = await process_single_url(url, user_id)
        
        if timestamps:
            # Если результат слишком длинный, отправляем файлом
            if len(timestamps) > 4000:
                # Создаем временный файл
                with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
                    f.write(f"Тайм-коды для: {url}\n\n")
                    f.write(timestamps)
                    temp_file = f.name
                
                # Отправляем файл
                await event.reply("📝 Результат слишком длинный, отправляю файлом:")
                await client.send_file(user_id, temp_file)
                
                # Удаляем временный файл
                try:
                    os.unlink(temp_file)
                except:
                    pass
            else:
                # Отправляем текстом
                await event.reply(f"📝 Тайм-коды для {url}:\n\n{timestamps}")
        else:
            await event.reply("❌ Не удалось обработать ссылку. Проверьте, что ссылка корректна и ведет к видео.")
        
        return

    # Команда обработки Excel
    if text.strip().lower() == '/timecod':
        user_dir = os.path.join("user_data", str(user_id))
        input_dir = os.path.join(user_dir, "in")
        output_dir = os.path.join(user_dir, "out")
        zip_path = os.path.join(user_dir, f"timecodes_{user_id}.zip")

        # Ищем Excel-файл в папке
        excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))] if os.path.exists(input_dir) else []
        
        if not excel_files:
            await event.reply("⚠️ Сначала загрузите Excel-файл со ссылками!")
            return

        input_file = os.path.join(input_dir, excel_files[0])
        
        await event.reply("⏳ Обрабатываю видео, это может занять длительное время...")
        
        processed, failed = await process_excel(user_id, input_file, output_dir)
        
        if processed > 0:
            # Создаем ZIP-архив
            make_zip(output_dir, zip_path)
            
            # Отправляем архив пользователю
            message = f"✅ Готово! Обработано видео: {processed}"
            if failed > 0:
                message += f", не удалось обработать: {failed}"
            message += "\nВаш архив с тайм-кодами:"
            
            await event.reply(message)
            await client.send_file(user_id, zip_path)
            
            # Очищаем временные файлы
            for file in os.listdir(output_dir):
                if file.endswith('.docx'):
                    os.remove(os.path.join(output_dir, file))
            if os.path.exists(zip_path):
                os.remove(zip_path)
        else:
            await event.reply("❌ Не удалось обработать ни одного видео. Проверьте лог для подробностей.")
        
        return

    # Ответ на неизвестные сообщения
    if text and not text.startswith('/'):
        await event.reply(
            "Отправьте мне:\n"
            "1. Excel-файл со ссылками на видео\n"
            "2. Прямую ссылку на видео\n"
            "Затем используйте команду /timecod для обработки Excel-файла."
        )

# ---------- Запуск ----------

async def main():
    """Основная функция запуска бота"""
    try:
        await client.run_until_disconnected()
    except Exception as e:
        logger.exception(f"Ошибка в основном цикле: {e}")

if __name__ == '__main__':
    try:
        # Проверяем наличие необходимых зависимостей
        try:
            import yt_dlp
            from deepgram import Deepgram
        except ImportError as e:
            logger.critical(f"Не установлены необходимые зависимости: {e}")
            logger.critical("Установите их: pip install yt-dlp deepgram-sdk")
            sys.exit(1)
            
        # Проверяем наличие FFmpeg
        try:
            subprocess.run([ffmpeg_path, '-version'], capture_output=True, check=True)
            logger.info("FFmpeg найден и работает корректно")
        except (subprocess.CalledProcessError, FileNotFoundError):
            logger.critical(f"FFmpeg не найден по указанному пути: {ffmpeg_path}")
            logger.critical("Убедитесь, что FFmpeg установлен и путь указан правильно")
            sys.exit(1)
            
        loop = asyncio.get_event_loop()
        loop.run_until_complete(main())
    except KeyboardInterrupt:
        logger.info("Бот остановлен пользователем")
    except Exception as e:
        logger.critical(f"Фатальная ошибка при запуске: {str(e)}")
