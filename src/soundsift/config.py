import pathlib

SRC_PATH = pathlib.Path(__file__).parent.resolve()
LOGGER_CONFIG_PATH = SRC_PATH / 'logger_conf.json'
LOGS_PATH = pathlib.Path.home() / '.soundsift' / 'soundsift.log'
CREDENTIALS_PATH = pathlib.Path.home() / '.soundsift'

try:
    CREDENTIALS_PATH.mkdir()

except FileExistsError:
    pass


CREDENTIALS = {
    'youtube': [],
    'spotify': [],
}
