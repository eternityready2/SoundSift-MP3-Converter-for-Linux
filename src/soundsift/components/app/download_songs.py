import re
import logging
import shutil
import os
import yt_dlp
import spotipy
from soundsift.config import CREDENTIALS
from soundsift.components.drivers.archive import ArchiveManager
from soundsift.components.drivers.YouTube import Ytube
from soundsift.components.drivers.Spotify import SpotifyDownloader

logger = logging.getLogger(__name__)

class MusicDownloader:
    instances = []
    YOUTUBE_PATTERN = r'youtube\.com|youtu\.be'
    SPOTIFY_PATTERN = r'spotify\.com'

    def __init__(self, status, url, source):
        self.status = status
        self.url = url
        self.source = source
        MusicDownloader.instances.append(self)

    @classmethod
    def identify_source(cls, url):
        """Determine the source and type of a URL."""
        url = str(url).lower()
        if re.search(cls.YOUTUBE_PATTERN, url):
            if 'playlist' in url:
                return 'youtube_playlist'
            else:
                return 'youtube_video'
        elif re.search(cls.SPOTIFY_PATTERN, url):
            if 'track' in url:
                return 'spotify_track'
            elif 'album' in url:
                return 'spotify_album'
            elif 'playlist' in url:
                return 'spotify_playlist'
        return 'unknown'

    @classmethod
    def get_sub_item_count(cls, url, source):
        """Get the number of sub-items for playlists and albums."""
        if source == 'youtube_playlist':
            ydl_opts = {'quiet': True, 'extract_flat': True}
            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(url, download=False)
                    return len(info['entries']) if 'entries' in info else 0
            except Exception as e:
                logger.error(f"Error fetching YouTube playlist length: {e}")
                return 0
        elif source in ['spotify_album', 'spotify_playlist']:
            for sp_idx, (client_id, client_secret, *_) in enumerate(CREDENTIALS['spotify']):
                try:
                    sp = SpotifyDownloader.authenticate_spotify(client_id, client_secret)
                    link_type = SpotifyDownloader.get_spotify_link_type(url)
                    item_id = SpotifyDownloader.extract_item_id(url)
                    if link_type == "playlist":
                        playlist = sp.playlist(item_id)
                        return playlist['tracks']['total']
                    elif link_type == "album":
                        album = sp.album(item_id)
                        return len(album['tracks']['items'])
                    return 0

                except spotipy.SpotifyBaseException as e:
                    if sp_idx == len(CREDENTIALS['spotify']) - 1:
                        logger.error(f"Spotify API count error: {e}, this was the last key.")
                        CREDENTIALS['spotify'][sp_idx][2] = "failed"
                        return 0

                    logger.error(f"Spotify API count error: {e}, changing to a new key of the pool...")
                    CREDENTIALS['spotify'][sp_idx][2] = "failed"
                    continue

                except Exception as e:
                    logger.error(f"Error fetching Spotify count.")
                    return 0

                return 0
            return 0
        return 1  # Single items

    @classmethod
    def download_music(cls, url, output_path, callback=None):
        """Download music based on URL type."""
        source = cls.identify_source(url)

        try:
            if source == 'youtube_playlist':
                status, msg, downloaded_files, temp_dir = Ytube.download_playlist(url, output_path, callback)
                if status == 'Success':
                    zip_name = f"youtube_playlist_{url.split('=')[-1]}.zip"
                    zip_path = ArchiveManager.create_zip(temp_dir, downloaded_files, zip_name)
                    if zip_path:
                        final_zip_path = os.path.join(output_path, zip_name)
                        shutil.move(zip_path, final_zip_path)
                        shutil.rmtree(temp_dir)
                        logger.info(f"Created ZIP archive: {final_zip_path}")
                    else:
                        logger.error("Failed to create ZIP archive.")
                return status, msg

            elif source == 'youtube_video':
                # Handle three-value return from download_audio_yt_dlp
                result = Ytube.download_audio_yt_dlp(url, output_path, callback=callback)
                if len(result) == 3:
                    status, msg, _ = result  # Ignore mp3_path for youtube_video
                else:
                    status, msg = result  # Fallback for unexpected return
                return status, msg

            elif source == 'spotify_track':
                status, msg = SpotifyDownloader.download_spotify_tracks(url, output_path, callback)
                return status, msg

            elif source in ['spotify_album', 'spotify_playlist']:
                status, msg = SpotifyDownloader.download_spotify_tracks(url, output_path, callback)
                return status, msg

            else:
                logger.error(f"Unknown source for URL: {url}")
                return 'Failed', 'Invalid or unsupported URL'

        except Exception as e:
            logger.error(f"Error processing {url}: {e}")
            return 'Failed', str(e)
