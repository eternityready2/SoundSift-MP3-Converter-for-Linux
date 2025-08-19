import logging
import re
import spotipy
import requests
from spotify_dl.spotify import fetch_tracks
from spotify_dl.youtube import download_songs
from spotipy.oauth2 import SpotifyClientCredentials
from soundsift.components.services.ConfigHandler import Config as CFG
from soundsift.components.drivers.YouTube import Ytube

class SpotifyDownloader:
    SPOTIPY_CLIENT_ID = '' # Update your Spotify client ID
    SPOTIPY_CLIENT_SECRET = '' # Update your Spotify client secret
    SPOTIPY_REDIRECT_URI = 'http://localhost:8888/callback'  # Update your Spotify redirect link
    YOUTUBE_API_KEY = ''  # Add your YouTube Data API key here. Create an account and access: https://console.cloud.google.com/apis/library/youtube.googleapis.com?inv=1&invt=Ab51TQ&project=my-project-5585-1751544391774

    @classmethod
    def authenticate_spotify(cls):
        client_credentials_manager = SpotifyClientCredentials(
            client_id=cls.SPOTIPY_CLIENT_ID,
            client_secret=cls.SPOTIPY_CLIENT_SECRET
        )
        sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)
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
    def get_youtube_url(cls, search_query):
        search_url = f"https://www.googleapis.com/youtube/v3/search?part=snippet&q={search_query}&key={cls.YOUTUBE_API_KEY}&type=video"
        response = requests.get(search_url)
        if response.status_code == 200:
            results = response.json().get('items', [])
            if results:
                return f"https://www.youtube.com/watch?v={results[0]['id']['videoId']}"
        return None

    @classmethod
    def download_spotify_tracks(cls, playlist_url, output_path="downloads"):
        # Set up logging
        sp = cls.authenticate_spotify()
        logger = logging.getLogger('spotify_dl')
        logger.setLevel(logging.DEBUG)
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        logger.addHandler(ch)

        # Scan link type and fetch tracks accordingly
        spotify_link_type = cls.get_spotify_link_type(playlist_url)
        if spotify_link_type == "unknown"or spotify_link_type == "episode" or spotify_link_type == "artist" or spotify_link_type == "album":
            print("Invalid Spotify track or playlist link.")
            return

        try:
            print("Extracting playlist ID...")
            item_id = cls.extract_item_id(playlist_url)
            print(f"Playlist ID extracted: {item_id}")
        except ValueError as e:
            print(f"Invalid playlist URL: {e}")
            return


        # Fetch tracks from the playlist
        try:
            print("Fetching tracks from Spotify playlist...")
            tracks = fetch_tracks(sp, str(spotify_link_type), item_id)

            if not tracks:
                print("No tracks found in the playlist. Please check the URL or playlist privacy settings.")
                return

            # Fetch YouTube URLs for each track
            for track in tracks:
                track_search_query = f"{track['name']} {track['artist']}"
                track['track_url'] = cls.get_youtube_url(track_search_query)

        except Exception as e:
            print(f"An error occurred in fetching tracks: {e}")
            return

        # print youtube urls for each track
        for track in tracks:
            print(f"Processing: {track['name']} - {track['artist']}")
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

        print("Download completed!")

# Example usage:
# SpotifyDownloader.download_spotify_tracks('https://open.spotify.com/playlist/YOUR_PLAYLIST_ID')
