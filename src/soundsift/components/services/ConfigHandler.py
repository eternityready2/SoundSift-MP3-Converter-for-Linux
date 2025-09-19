import os
from configparser import ConfigParser

class Config:
    Class_Init_sts = False
    Root_Path = ''
    str_Download_path = ''



    @classmethod
    def get_lstr_removechars(cls):
        if cls.Class_Init_sts is False:
            cls.import_config()
        return cls.lstr_removechars

    @classmethod
    def get_Root_Path(cls):
        if cls.Class_Init_sts is False:
            cls.import_config()
        return cls.Root_Path

    @classmethod
    def get_Download_Path(cls):
        if cls.Class_Init_sts is False:
            cls.import_config()
        return cls.str_Download_path

    @classmethod
    def set_Download_Path(cls,path):
        if cls.Class_Init_sts is False:
            cls.import_config()
        cls.str_Download_path = path
    
    @classmethod
    def get_downloads_folder(cls):
        try:
            path = subprocess.check_output(
                ['xdg-user-dir', 'DOWNLOAD'],
                universal_newlines=True
            ).strip()
        
            # Verify the path exists
            if os.path.isdir(path):
                return path
            else:
                # Fallback to the default if the command fails or path doesn't exist
                return os.path.expanduser("~/Downloads")
        except Exception:
            return os.path.expanduser("~/Downloads")

    @classmethod
    def import_config(cls):
        ROOT_DIR = os.path.dirname(os.path.abspath('Redmine_Report'))
        cfg_file = os.path.join(ROOT_DIR, "cfg","config.ini")
        proj_config = ConfigParser()
        proj_config.read(cfg_file)
        cls.Root_Path = ROOT_DIR

        # ------Paths--------

        #cls.str_Download_path= str(proj_config['paths']['str_download_path'])
        #cls.str_Download_path= os.path.join(ROOT_DIR,'downloads')
        cls.str_Download_path = cls.get_downloads_folder()
        cls.Class_Init_sts = True


