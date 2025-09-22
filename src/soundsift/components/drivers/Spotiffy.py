import logging
import re
import spotipy
import requests
from spotify_dl.spotify import fetch_tracks
from spotify_dl.youtube import download_songs
from spotipy.oauth2 import SpotifyClientCredentials
from soundsift.components.services.ConfigHandler import Config as CFG
from soundsift.components.drivers.YouTube import Ytube
from soundsift.config import CREDENTIALS

logger = logging.getLogger(__name__)

class SpotifyDownloader:
    SPOTIPY_CLIENT_ID = '' # Update your Spotify client ID
    SPOTIPY_CLIENT_SECRET = '' # Update your Spotify client secret
    SPOTIPY_REDIRECT_URI = 'http://localhost:8888/callback'  # Update your Spotify redirect link
    YOUTUBE_API_KEY = ''  # Add your YouTube Data API key here. Create an account and access: https://console.cloud.google.com/apis/library/youtube.googleapis.com?inv=1&invt=Ab51TQ&project=my-project-5585-1751544391774

    @classmethod
    def authenticate_spotify(cls, client_id, client_secret):
        client_credentials_manager = SpotifyClientCredentials(
            client_id=client_id,
            client_secret=client_secret
        )
        sp = spotipy.Spotify(
            client_credentials_manager=client_credentials_manager
        )
        return sp

    @classmethod
    def get_spotify_link_type(cls,url):
        """
        Determines the type of Spotify link (playlist, track, album, artist, etc.)
        Args:
            url (str): Spotify URL
        Returns:
            str: The type of Spotify link ('playlist', 'track', 'album', 'artist', or 'unknown')
        """
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
        match_playlist = re.search(r"playlist/([a-zA-Z0-9]+)", url)
        match_track = re.search(r"track/([a-zA-Z0-9]+)", url)
        if match_playlist:
            return match_playlist.group(1)
        elif match_track:
            return match_track.group(1)
        else:
            raise ValueError("Invalid Spotify playlist URL. Could not extract playlist ID.")

    @classmethod
    def get_youtube_url(cls, search_query, youtube_api_key):
        search_url = f"https://www.googleapis.com/youtube/v3/search?part=snippet&q={search_query}&key={youtube_api_key}&type=video"
        try:
            response = requests.get(search_url)
            if response.status_code == 200:
                results = response.json().get('items', [])
                if results:
                    return f"https://www.youtube.com/watch?v={results[0]['id']['videoId']}"
            logger.error(response.json()['error']['message'])

        except Exception as error:
            pass

        return None

    @classmethod
    def download_spotify_tracks(cls,
                                playlist_url,
                                output_path="downloads",
                                ):
        # Set up logging
        #logger = logging.getLogger('spotify_dl')
        #logger.setLevel(logging.DEBUG)
        #ch = logging.StreamHandler()
        #ch.setLevel(logging.DEBUG)
        #formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        #ch.setFormatter(formatter)
        #logger.addHandler(ch)

        # Scan link type and fetch tracks accordingly
        """
        spotify_link_type = cls.get_spotify_link_type(playlist_url)
        if spotify_link_type == "unknown"or spotify_link_type == "episode" or spotify_link_type == "artist" or spotify_link_type == "album":
            logger.debug("Invalid Spotify track or playlist link.")
            return

        try:
            logger.debug("Extracting playlist ID...")
            item_id = cls.extract_item_id(playlist_url)
            logger.debug(f"Playlist ID extracted: {item_id}")
        except ValueError as e:
            logger.debug(f"Invalid playlist URL: {e}")
            return
        """


        # Fetch tracks from the playlist
        tracks = []
        for sp_idx, (client_id, client_secret, *_) in enumerate(CREDENTIALS['spotify']):
            try:
                logger.debug("Fetching tracks from Spotify playlist...")
                logger.debug(f"Spotify credential {sp_idx+1}/{len(CREDENTIALS['spotify'])}")
                logger.debug(f"CLIENT_ID = {client_id}")
                logger.debug(f"CLIENT_SECRET = {client_secret}")

                #sp = cls.authenticate_spotify(client_id, client_secret)
                #tracks = fetch_tracks(sp, str(spotify_link_type), item_id)
                tracks = [
                    {'name': "Beatufiul", 'artist': "Kmbra"},
                    {'name': "Beatufiul", 'artist': "Kmbra"},
                    {'name': "Beatufiul", 'artist': "Kmbra"},
                    {'name': "Beatufiul", 'artist': "Kmbra"},
                    {'name': "Beatufiul", 'artist': "Kmbra"},
                ]

                if not tracks:
                    logger.debug("No tracks found in the playlist. Please check the URL or playlist privacy settings.")
                    return

                # Fetch YouTube URLs for each track
                logger.debug("GET youtube urls of each track of the playlist")
                for track_idx, track in enumerate(tracks):
                    logger.debug(f"Fetch youtube url of track {track_idx+1} of {len(tracks)}")
                    for yt_idx, (youtube_api_key, *_) in enumerate(CREDENTIALS['youtube']):
                        track_search_query = f"{track['name']} {track['artist']}"
                        logger.debug(f"GET track '{track_search_query}' with API_KEY='{youtube_api_key}' {yt_idx+1}/{len(CREDENTIALS['youtube'])}")
                        track['track_url'] = cls.get_youtube_url(track_search_query, youtube_api_key)
                        if track['track_url'] is not None:
                            CREDENTIALS['youtube'][yt_idx][1] = "success"
                            break
                        CREDENTIALS['youtube'][yt_idx][1] = "failed"
                        logger.debug(f"{youtube_api_key} failed, trying another of the pool...")
                        print('---')
                CREDENTIALS['spotify'][sp_idx][2] = "success"
                break

            except Exception as e:
                logger.error(f"An error occurred in fetching tracks: {e}")
                CREDENTIALS['spotify'][sp_idx][2] = "failed"
                return

        if not tracks or all(track['track_url'] is None for track in tracks):
            logger.error('Failed fetching tracks from playlist')
            return

        # logger.debug youtube urls for each track
        for track in tracks:
            if track['track_url'] is None:
                continue

            logger.debug(f"Processing: {track['name']} - {track['artist']}")
            #print(f"{track['name']} - {track['artist']}: {track['track_url']}")
            Ytube.download_audio_yt_dlp(track['track_url'])

        # # Prepare download parameters
        # download_params = {
        #     "songs": {"urls": [{"save_path": output_path}]},  # Mock structure, adapt as needed
        #     "output_dir": output_path,
        #     "multi_core": 1,
        # }
        #
        # # Adding the real track data in the expected format
        # tracks_data = []
        # for track in tracks:
        #     track_data = {
        #         'name': track['name'],
        #         'artist': track['artist'],
        #         'album': track['album'],
        #         'year': track['year'],
        #         'track_url': track['track_url']  # Updated to use the YouTube URL
        #     }
        #     tracks_data.append(track_data)
        # download_params["songs"]["urls"] = tracks_data

        # Download tracks as MP3
        #download_songs(**download_params)

        logger.debug("Download completed!")

# Example usage:
# SpotifyDownloader.download_spotify_tracks('https://open.spotify.com/playlist/YOUR_PLAYLIST_ID')
