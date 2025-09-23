import os
import zipfile
import logging

class ArchiveManager:
    @classmethod
    def create_zip(cls, output_path, files, zip_name="archive.zip"):
        """
        Creates a ZIP file containing the specified files and deletes them afterward.

        :param output_path: Directory where the ZIP file will be created.
        :param files: List of file paths to include in the ZIP.
        :param zip_name: Name of the ZIP file to create.
        :return: Path to the created ZIP file or None if an error occurred.
        """
        logger = logging.getLogger("soundsift")
        zip_path = os.path.join(output_path, zip_name)

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in files:
                    if os.path.exists(file):
                        zipf.write(file, os.path.basename(file))
                        print(f"Added {file} to ZIP archive.")
                        if logger.handlers and logger.isEnabledFor(logging.INFO):
                            logger.info(f"Added {file} to ZIP archive.")
                    else:
                        print(f"File not found: {file}")
                        if logger.handlers and logger.isEnabledFor(logging.WARNING):
                            logger.warning(f"File not found: {file}")

            # Delete the individual files after adding to ZIP
            for file in files:
                if os.path.exists(file):
                    os.remove(file)
                    print(f"Deleted file: {file}")
                    if logger.handlers and logger.isEnabledFor(logging.INFO):
                        logger.info(f"Deleted file: {file}")

            print(f"ZIP archive created at {zip_path}")
            if logger.handlers and logger.isEnabledFor(logging.INFO):
                logger.info(f"ZIP archive created at {zip_path}")
            return zip_path

        except Exception as e:
            print(f"Error creating ZIP archive: {e}")
            if logger.handlers and logger.isEnabledFor(logging.ERROR):
                logger.error(f"Error creating ZIP archive: {e}")
            return None
