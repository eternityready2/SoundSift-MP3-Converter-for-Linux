import logging
import re
import os
import uuid
import shutil
import spotipy
import requests
from spotipy.oauth2 import SpotifyClientCredentials
from soundsift.components.drivers.YouTube import Ytube
from soundsift.components.drivers.archive import ArchiveManager
from soundsift.config import CREDENTIALS
from dotenv import load_dotenv


load_dotenv()

logger = logging.getLogger(__name__)

class SpotifyDownloader:
    SPOTIPY_CLIENT_ID = 'c391704d72a84400a96b718a786aa732'
    SPOTIPY_CLIENT_SECRET = '33e355197c5e4ff49245f53ff6a276b1'
    SPOTIPY_REDIRECT_URI = 'http://localhost:8888/callback'

    # Converted to a list of API keys
    YOUTUBE_API_KEYS = [
        'aAIzaSyDTvJnvE9Q7psr4RosC_w2D2vxg_gnDEyg',
        'bAIzaSyD3JyU15d15GNLdLWqdMTFQBaqwfqkxrP0',# Your current key
        # Add more keys here as needed, e.g.:
        # 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXX',
        # 'AIzaSyYYYYYYYYYYYYYYYYYYYYYYYYYYYYY',
    ]

    @classmethod
    def authenticate_spotify(cls, client_id, client_secret):
        credentials = SpotifyClientCredentials(
            client_id=client_id,
            client_secret=client_secret
        )
        return spotipy.Spotify(client_credentials_manager=credentials)

    @classmethod
    def get_spotify_link_type(cls, url):
        patterns = {
            "playlist": r"open\.spotify\.com/playlist/",
            "track": r"open\.spotify\.com/track/",
            "album": r"open\.spotify\.com/album/",
            "artist": r"open\.spotify\.com/artist/",
            "episode": r"open\.spotify\.com/episode/"
        }
        for link_type, pattern in patterns.items():
            if re.search(pattern, url):
                return link_type
        return "unknown"

    @classmethod
    def extract_item_id(cls, url):
        patterns = {
            "album": r"album/([a-zA-Z0-9]+)",
            "playlist": r"playlist/([a-zA-Z0-9]+)",
            "track": r"track/([a-zA-Z0-9]+)"
        }
        for type_, pattern in patterns.items():
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        raise ValueError("Invalid Spotify URL.")

    @classmethod
    def get_youtube_url(cls, search_query):
        for yt_idx, (api_key, *_) in enumerate(CREDENTIALS['youtube']):
            search_url = f"https://www.googleapis.com/youtube/v3/search?part=snippet&q={search_query}&key={api_key}&type=video"
            try:
                response = requests.get(search_url, timeout=10)
                response.raise_for_status()
                results = response.json().get('items', [])

                if results:
                    youtube_url = f"https://www.youtube.com/watch?v={results[0]['id']['videoId']}"
                    logger.info(f"Fetched YouTube URL with API key {yt_idx+1}/{len(CREDENTIALS['youtube'])}: {api_key[:8]}...")
                    CREDENTIALS['youtube'][yt_idx][1] = "success"
                    return youtube_url
            except requests.HTTPError as e:
                if e.response.status_code in {400, 403}:
                    message = f"API key failed {yt_idx+1}/{len(CREDENTIALS['youtube'])}: {api_key[:8]}... {e.response.json()['error']['message'].strip(".")}"
                    if yt_idx != len(CREDENTIALS['youtube']) - 1:
                        message += ", trying next key..."

                    else:
                        message += ", this wast the last key..."

                    CREDENTIALS['youtube'][yt_idx][1] = "failed"
                    logger.warning(message)
                    continue
                logger.error(f"YouTube API error: {e}")
                CREDENTIALS['youtube'][yt_idx][1] = "failed"
                return None
            except requests.RequestException as e:
                logger.error(f"YouTube API request failed: {e}")
                return None
        logger.error(
            "All YT API keys failed, stop the downloading of tracks..."
        )
        raise ValueError("All YT API keys failed, stop the downloading of tracks...")

    @classmethod
    def fetch_track_metadata(cls, sp, track_id):
        try:
            track_info = sp.track(track_id)
            metadata = {
                "title": track_info["name"],
                "artist": track_info["artists"][0]["name"],
                "album": track_info["album"]["name"],
                "date": track_info["album"]["release_date"],
                "thumbnail": track_info["album"]["images"][0]["url"] if track_info["album"]["images"] else None
            }
            logger.info(f"Track Metadata: {metadata}")
            return metadata
        except Exception as e:
            logger.error(f"Error fetching metadata for track {track_id}: {e}")
            return None

    @classmethod
    def fetch_playlist_or_album_tracks(cls, sp, link_type, item_id):
        tracks = []
        if link_type == "playlist":
            results = sp.playlist_tracks(item_id)
            for item in results["items"]:
                track = item["track"]
                tracks.append({
                    "name": track["name"],
                    "artist": track["artists"][0]["name"],
                    "track_id": track["id"],
                    "url": track["external_urls"]["spotify"]
                })
            while results["next"]:
                results = sp.next(results)
                for item in results["items"]:
                    track = item["track"]
                    tracks.append({
                        "name": track["name"],
                        "artist": track["artists"][0]["name"],
                        "track_id": track["id"],
                        "url": track["external_urls"]["spotify"]
                    })
        elif link_type == "album":
            results = sp.album_tracks(item_id)
            for track in results["items"]:
                tracks.append({
                    "name": track["name"],
                    "artist": track["artists"][0]["name"],
                    "track_id": track["id"],
                    "url": track["external_urls"]["spotify"]
                })
        return tracks

    @classmethod
    def download_spotify_tracks(cls, playlist_url, output_path, callback=None):
        link_type = cls.get_spotify_link_type(playlist_url)
        if link_type in ["unknown", "episode", "artist"]:
            logger.error(f"Unsupported Spotify link type: {link_type}")
            return "Failed", f"Unsupported link type: {link_type}"

        try:
            item_id = cls.extract_item_id(playlist_url)
            logger.info(f"Extracted {link_type} ID: {item_id}")

        except ValueError as e:
            logger.error(f"Invalid URL: {e}")
            return "Failed", str(e)
        try:
            if link_type == "track":
                tracks = [{"track_id": item_id, "url": playlist_url}]
                temp_dir = output_path

            else:
                temp_dir = os.path.join(output_path, f"{link_type}_{uuid.uuid4().hex}")
                os.makedirs(temp_dir, exist_ok=True)
                for sp_idx, (client_id, client_secret, *_) in enumerate(CREDENTIALS['spotify']):
                    try:
                        logger.info(f"Fetching {link_type} tracks from Spotify...")
                        logger.info(f"Spotify credential {sp_idx+1}/{len(CREDENTIALS['spotify'])}")
                        logger.info(f"CLIENT_ID = {client_id}")
                        logger.info(f"CLIENT_SECRET = {client_secret}")
                        sp = cls.authenticate_spotify(client_id, client_secret)
                        tracks = cls.fetch_playlist_or_album_tracks(sp, link_type, item_id)
                        CREDENTIALS['spotify'][sp_idx][2] = "success"
                        break

                    except spotipy.SpotifyException as e:
                        if sp_idx == len(CREDENTIALS['spotify']) - 1:
                            logger.error(f"Spotify API error: {e}, this is the last key, exiting download")
                            CREDENTIALS['spotify'][sp_idx][2] = "failed"
                            return "Failed", str(e)

                        logger.error(f"Spotify API error: {e}, changing to a new key of the pool...")
                        CREDENTIALS['spotify'][sp_idx][2] = "failed"
                        continue

            if not tracks:
                logger.warning("No tracks found.")
                return "Failed", "No tracks found"

            downloaded_files = []

            for track in tracks:
                track_id = track.get("track_id")
                if not track_id and link_type == "track":
                    track_id = item_id

                for sp_idx, (client_id, client_secret, *_) in enumerate(CREDENTIALS['spotify']):
                    try:
                        logger.info(f"Fetching track metadata from spotify...")
                        logger.info(f"Spotify credential {sp_idx+1}/{len(CREDENTIALS['spotify'])}")
                        logger.info(f"CLIENT_ID = {client_id}")
                        logger.info(f"CLIENT_SECRET = {client_secret}")

                        metadata = cls.fetch_track_metadata(sp, track_id) if track_id else {
                            "title": track.get("name"),
                            "artist": track.get("artist"),
                            "album": None,
                            "date": None,
                            "thumbnail": None
                        }
                        CREDENTIALS['spotify'][sp_idx][2] = "success"

                    except spotipy.SpotifyException as e:
                        if sp_idx == len(CREDENTIALS['spotify']) - 1:
                            logger.error(f"Spotify API error: {e}, this is the last key, exiting download")
                            CREDENTIALS['spotify'][sp_idx][2] = "failed"
                            return "Failed", str(e)

                        logger.error(f"Spotify API error: {e}, changing to a new key of the pool...")
                        CREDENTIALS['spotify'][sp_idx][2] = "failed"
                        continue

                query = f"{metadata['title']} {metadata['artist']}"
                try:
                    youtube_url = cls.get_youtube_url(query)
                except ValueError as e:
                    return "Failed", str(e)

                if youtube_url:
                    logger.info(f"Downloading: {metadata['title']} - {metadata['artist']}")
                    if callback:
                        callback("start_download")
                    status, msg, mp3_path = Ytube.download_audio_yt_dlp(youtube_url, temp_dir, metadata=metadata, callback=callback)
                    if status == "Success" and mp3_path and os.path.exists(mp3_path):
                        downloaded_files.append(mp3_path)
                    else:
                        logger.warning(f"Failed to download {metadata['title']} - {metadata['artist']}: {msg}")
                else:
                    logger.warning(f"No YouTube URL for {metadata['title']} - {metadata['artist']}")

            if link_type in ["playlist", "album"] and downloaded_files:
                zip_name = f"{link_type}_{item_id}.zip"
                zip_path = ArchiveManager.create_zip(temp_dir, downloaded_files, zip_name)
                if zip_path:
                    final_zip_path = os.path.join(output_path, zip_name)
                    shutil.move(zip_path, final_zip_path)
                    shutil.rmtree(temp_dir)
                    logger.info(f"Created ZIP archive: {final_zip_path}")
                else:
                    logger.error("Failed to create ZIP archive.")
            return "Success", "Download completed"

        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            return "Failed", str(e)

        logger.error(f"All spotify keys are not working, stopping download...")
        return "Failed", "All Spotify Keys failed"
