import os
import re
from soundsift.components.services.ConfigHandler import Config as CFG
from soundsift.components.drivers.Excel import Excel as xls
from soundsift.components.drivers.YouTube import Ytube
from soundsift.components.drivers.Spotiffy import SpotifyDownloader

class appl:
    lso_instance = []
    b_initcls_flg = False
    str_imputs_path = ''
    str_imputs_sheet = ''

    def __init__(self,status,url,dlsource):
        self.status = str(status)
        self.url = str(url)
        self.dlsource = str(dlsource)
        appl.lso_instance.append(self)

    @classmethod
    def classinit(cls):
        cls.b_initcls_flg = True
        #cls.str_imputs_path = os.path.join(CFG.get_Root_Path(),'Inputs.xlsx')
        cls.str_imputs_sheet= 'Playlist'

    @classmethod
    def import_data(cls):
        if cls.b_initcls_flg == False:
            cls.classinit()
        xls.Import_Sheet(cls.str_imputs_path,cls.str_imputs_sheet)
        data = xls.Get_Excel_Data()
        for el in data:
            if bool(re.search('youtube', str(el.el2))):
                cls(el.el1,el.el2,'youtube')
            elif bool(re.search('spotify', str(el.el2))):
                cls(el.el1,el.el2,'spotify')
            else:
                cls(el.el1,el.el2,'na')
        xls.Remove_All_Data()

    @classmethod
    def import_data_cmd(cls,urls):
        if cls.b_initcls_flg == False:
            cls.classinit()
        for el in urls:
            if bool(re.search('youtube', str(el))):
                cls('Pending',el,'youtube')
            elif bool(re.search('spotify', str(el))):
                cls('Pending',el,'spotify')
            else:
                cls('Pending',el,'na')

    @classmethod
    def save_data(cls):
        if cls.b_initcls_flg == False:
            cls.classinit()
        xls.Remove_All_Data()
        #xls.Add_Data('Status','Link')
        for el in cls.lso_instance:
            xls.Add_Data(el.status,el.url)
        xls.Write_Sheet_Including_Headders(cls.str_imputs_path,cls.str_imputs_sheet)
        xls.Remove_All_Data()

    @classmethod
    def download_music(cls):
        if cls.b_initcls_flg == False:
            cls.classinit()
        for el in cls.lso_instance:
            if el.status != 'Downloded':
                if el.dlsource == 'youtube':
                    try:
                        Ytube.download_audio_yt_dlp(el.url,CFG.get_Download_Path())
                        el.status = 'Downloded'
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        el.status = 'Failed'


                elif el.dlsource == 'spotify':
                    try:
                        SpotifyDownloader.download_spotify_tracks(el.url)
                        el.status = 'Downloded'
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        el.status = 'Failed'
                else:
                    print ('Error: Unknown source of download - '+str(el.dlsource))



    @classmethod
    def download_music_direct(cls,url):
        if cls.b_initcls_flg == False:
            cls.classinit()

        if bool(re.search('youtube', str(url))):
            try:
                Ytube.download_audio_yt_dlp(url,CFG.get_Download_Path())
                return 'Success'
            except Exception as e:
                print(f"An error occurred: {e}")
                return 'Failed'

        elif bool(re.search('spotify', str(url))):
            try:
                SpotifyDownloader.download_spotify_tracks(url,CFG.get_Download_Path())
                return 'Success'
            except Exception as e:
                print(f"An error occurred: {e}")
                return 'Download Failed'
        else:
            print ('Error: Unknown source of download - '+str(url))
            return 'Incorrect Link'





