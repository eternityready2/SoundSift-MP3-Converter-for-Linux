import pathlib

CONFIG_PATH = pathlib.Path(__file__).parent.resolve()
CREDENTIALS_PATH = pathlib.Path.home() / '.soundsift-credentials'
try:
    CREDENTIALS_PATH.mkdir()

except FileExistsError:
    pass


CREDENTIALS = {
    'youtube': [],
    'spotify': [],
}

try:
    with open(CREDENTIALS_PATH / 'spotify-credentials.csv', 'r') as file:
        for line in file.readlines():
            CREDENTIALS['spotify'].append(line.strip().split(','))

except FileNotFoundError:
    with open(CREDENTIALS_PATH / 'spotify-credentials.csv', 'w') as file:
        pass

try:
    with open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'r') as file:
        for line in file.readlines():
            CREDENTIALS['youtube'].append(line.strip().split(','))

except FileNotFoundError:
    with open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'w') as file:
        pass
