import os
import logging
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3, APIC
from mutagen import File

logger = logging.getLogger(__name__)

class MetadataMP3:
    """A singleton-like class to handle MP3 metadata application and renaming."""

    @classmethod
    def apply_metadata(cls, mp3_path, metadata):
        """Apply metadata to the MP3 file, embed the thumbnail, delete it, and remove encoder settings."""
        try:
            # Load or create ID3 tags
            audio = EasyID3(mp3_path) if os.path.exists(mp3_path) and EasyID3.valid_keys else File(mp3_path, easy=True)
            if audio is None:
                audio = File(mp3_path, easy=True)
                audio.add_tags()

            # Apply available metadata
            if "title" in metadata and metadata["title"]:
                audio["title"] = metadata["title"]
            else:
                if logger.handlers and logger.isEnabledFor(logging.WARNING):
                    logger.warning(f"Title metadata missing for {mp3_path}")
                print(f"Warning: Title metadata missing for {mp3_path}")

            if "artist" in metadata and metadata["artist"]:
                audio["artist"] = metadata["artist"]
            else:
                if logger.handlers and logger.isEnabledFor(logging.WARNING):
                    logger.warning(f"Artist metadata missing for {mp3_path}")
                print(f"Warning: Artist metadata missing for {mp3_path}")

            if "album" in metadata and metadata["album"]:
                audio["album"] = metadata["album"]
            else:
                if logger.handlers and logger.isEnabledFor(logging.WARNING):
                    logger.warning(f"Album metadata missing for {mp3_path}")
                print(f"Warning: Album metadata missing for {mp3_path}")

            if "date" in metadata and metadata["date"]:
                audio["date"] = metadata["date"]
            else:
                if logger.handlers and logger.isEnabledFor(logging.WARNING):
                    logger.warning(f"Date metadata missing for {mp3_path}")
                print(f"Warning: Date metadata missing for {mp3_path}")

            # Save the basic metadata
            audio.save()

            # Load full ID3 tags for advanced operations
            audio = ID3(mp3_path)

            # Embed thumbnail if available and delete it afterward
            thumbnail_path = f"{os.path.splitext(mp3_path)[0]}.jpg"
            if "thumbnail" in metadata and metadata["thumbnail"] and os.path.exists(thumbnail_path):
                with open(thumbnail_path, "rb") as img_file:
                    audio["APIC"] = APIC(
                        encoding=3,  # UTF-8
                        mime="image/jpeg",
                        type=3,  # Cover (front)
                        desc="Cover",
                        data=img_file.read()
                    )
                audio.save()
                if logger.handlers and logger.isEnabledFor(logging.INFO):
                    logger.info(f"Embedded thumbnail in {mp3_path}")
                print(f"Embedded thumbnail in {mp3_path}")

                # Delete the thumbnail file after embedding
                try:
                    os.remove(thumbnail_path)
                    if logger.handlers and logger.isEnabledFor(logging.INFO):
                        logger.info(f"Deleted thumbnail file: {thumbnail_path}")
                    print(f"Deleted thumbnail file: {thumbnail_path}")
                except OSError as e:
                    if logger.handlers and logger.isEnabledFor(logging.WARNING):
                        logger.warning(f"Failed to delete thumbnail {thumbnail_path}: {e}")
                    print(f"Failed to delete thumbnail {thumbnail_path}: {e}")

            # Check for and remove Encoder Settings (TSSE)
            if "TSSE" in audio:
                del audio["TSSE"]
                audio.save()
                if logger.handlers and logger.isEnabledFor(logging.INFO):
                    logger.info(f"Removed Encoder Settings from {mp3_path}")
                print(f"Removed Encoder Settings from {mp3_path}")

        except Exception as e:
            if logger.handlers and logger.isEnabledFor(logging.ERROR):
                logger.error(f"Error applying metadata to {mp3_path}: {e}")
            print(f"Error applying metadata to {mp3_path}: {e}")

    @classmethod
    def rename_file(cls, mp3_path, metadata):
        """Rename the MP3 file to [artist] - [title].mp3 if metadata is available."""
        if "artist" in metadata and "title" in metadata and metadata["artist"] and metadata["title"]:
            new_name = f"{metadata['artist']} - {metadata['title']}.mp3"
            new_path = os.path.join(os.path.dirname(mp3_path), new_name)
            try:
                os.rename(mp3_path, new_path)
                if logger.handlers and logger.isEnabledFor(logging.INFO):
                    logger.info(f"Renamed {mp3_path} to {new_path}")
                print(f"Renamed {mp3_path} to {new_path}")
                return new_path
            except Exception as e:
                if logger.handlers and logger.isEnabledFor(logging.ERROR):
                    logger.error(f"Error renaming {mp3_path}: {e}")
                print(f"Error renaming {mp3_path}: {e}")
                return mp3_path
        else:
            if logger.handlers and logger.isEnabledFor(logging.WARNING):
                logger.warning(f"Cannot rename {mp3_path}: Missing artist or title metadata")
            print(f"Cannot rename {mp3_path}: Missing artist or title metadata")
            return mp3_path
