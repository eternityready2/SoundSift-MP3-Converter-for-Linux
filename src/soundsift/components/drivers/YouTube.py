
import os
import uuid
import yt_dlp
import logging
import shutil
import requests
from soundsift.components.drivers.metadata_mp3 import MetadataMP3

class Ytube:
    FFMPEG_PATH = shutil.which("ffmpeg") or os.getenv("FFMPEG_PATH")
    if not FFMPEG_PATH:
        raise FileNotFoundError("FFmpeg not found. Install it or set FFMPEG_PATH.")

    @classmethod
    def download_thumbnail(cls, thumbnail_url, output_path, filename):
        """Download thumbnail image."""
        logger = logging.getLogger("soundsift")
        if not thumbnail_url:
            return
        try:
            response = requests.get(thumbnail_url, stream=True, timeout=10)
            response.raise_for_status()
            thumbnail_path = os.path.join(output_path, f"{filename}.jpg")
            with open(thumbnail_path, "wb") as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            logger.info(f"Thumbnail downloaded to {thumbnail_path}")
            print(f"Thumbnail downloaded to {thumbnail_path}")
            return thumbnail_path
        except requests.RequestException as e:
            logger.error(f"Error downloading thumbnail: {e}")
            print(f"Error downloading thumbnail: {e}")
            return None

    @classmethod
    def download_audio_yt_dlp(cls, url, output_path, metadata=None, callback=None):
        """Download audio from YouTube with callbacks, returning the final MP3 path."""
        logger = logging.getLogger("soundsift")
        if not url or not isinstance(url, str):
            logger.error("Invalid URL provided.")
            return "Failed", "Invalid URL", None

        os.makedirs(output_path, exist_ok=True)

        ydl_opts = {
            'format': 'bestaudio/best',
            'ffmpeg_location': cls.FFMPEG_PATH,
            'noplaylist': True,
            'outtmpl': f"{output_path}/%(title)s.%(ext)s",
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '320',
            }],
        }

        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                if callback:
                    callback("start_download")
                logger.info(f"Downloading: {url}")
                print(f"Downloading: {url}")
                info = ydl.extract_info(url, download=True)
                # Get the actual downloaded file path after extraction
                downloaded_file = ydl.prepare_filename(info)
                base_ext = os.path.splitext(downloaded_file)[1]  # e.g., .webm, .m4a
                mp3_path = downloaded_file.replace(base_ext, ".mp3")

                if callback:
                    callback("start_conversion")
                # Conversion happens here via postprocessor
                if callback:
                    callback("end_conversion")

                # Ensure the file exists before proceeding
                if not os.path.exists(mp3_path):
                    logger.error(f"MP3 file not found after download: {mp3_path}")
                    return "Failed", "MP3 file not found after download", None

                if metadata and metadata.get("thumbnail"):
                    thumbnail_path = cls.download_thumbnail(metadata["thumbnail"], output_path, os.path.splitext(os.path.basename(mp3_path))[0])

                if metadata:
                    MetadataMP3.apply_metadata(mp3_path, metadata)
                    mp3_path = MetadataMP3.rename_file(mp3_path, metadata)

            return "Success", "Download completed", mp3_path
        except yt_dlp.utils.DownloadError as e:
            logger.error(f"Download error: {e}")
            print(f"Download error: {e}")
            return "Failed", str(e), None
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            print(f"Unexpected error: {e}")
            return "Failed", f"Unexpected error: {e}", None

    @classmethod
    def download_playlist(cls, url, output_path, callback=None):
        """Download YouTube playlist."""
        logger = logging.getLogger("soundsift")
        temp_dir = os.path.join(output_path, f"playlist_{uuid.uuid4().hex}")
        os.makedirs(temp_dir, exist_ok=True)

        ydl_opts = {
            'format': 'bestaudio/best',
            'ffmpeg_location': cls.FFMPEG_PATH,
            'outtmpl': f"{temp_dir}/%(title)s.%(ext)s",
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '320',
            }],
            'noplaylist': False,
        }

        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                logger.info(f"Downloading playlist: {url}")
                print(f"Downloading playlist: {url}")
                info = ydl.extract_info(url, download=False)
                downloaded_files = []
                for entry in info['entries']:
                    video_url = f"https://www.youtube.com/watch?v={entry['id']}"
                    status, msg, mp3_path = cls.download_audio_yt_dlp(video_url, temp_dir, callback=callback)
                    if status == "Success" and mp3_path and os.path.exists(mp3_path):
                        downloaded_files.append(mp3_path)
            return "Success", "Playlist download completed", downloaded_files, temp_dir
        except yt_dlp.utils.DownloadError as e:
            logger.error(f"Playlist download error: {e}")
            print(f"Playlist download error: {e}")
            return "Failed", str(e), [], temp_dir
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            print(f"Unexpected error: {e}")
            return "Failed", f"Unexpected error: {e}", [], temp_dir

