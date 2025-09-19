import os
import yt_dlp
from soundsift.components.services.ConfigHandler import Config

class Ytube:
    lso_instance = []
    b_initcls_flg = False
    str_ffmpeg_path = ''

    @classmethod
    def classinit(cls):
        cls.b_initcls_flg = True
        cls.str_ffmpeg_path = os.path.join(
            Config.get_Root_Path(),
            'thired_party', 'FFmpeg', 'bin', 'FFmpeg.exe'
        )
        # If you need pydub or similar, you can set:
        # AudioSegment.converter = cls.str_ffmpeg_path

    @classmethod
    def download_audio_yt_dlp(cls, url, output_path=Config.get_Download_Path()):
        """
        Downloads audio from a YouTube URL using yt-dlp.
        Converts it to mp3 at 192 kbps.
        """

        if not cls.b_initcls_flg:
            cls.classinit()

        print("Using FFmpeg at path:", cls.str_ffmpeg_path)

        # Common download options:
        ydl_opts = {
            'format': 'bestaudio/best',
            #'cookiesfrombrowser': ('chrome',),  # or ('firefox',), etc.
            #'ffmpeg_location': cls.str_ffmpeg_path,  # Path to the FFmpeg binary
            'noplaylist': True,                     # Set True if you only want single videos
            'outtmpl': str(output_path)+'/%(title)s.%(ext)s',
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],

            # OPTIONAL: Clear the cache to avoid stale data issues
            # 'rm_cachedir': True,
            # OPTIONAL: Add custom user-agent to avoid 403 if YT is blocking default agents
        }

        try:
            #with yt_dlp.YoutubeDL() as ydl:
            #    ydl.download([url])
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                print(f"Downloading: {url}")
                ydl.download([url])

        except yt_dlp.utils.DownloadError as e:
            # Handle download errors (e.g. HTTP 403, signature extraction, etc.)
            print(f"An error occurred while downloading {url}: {e}")
