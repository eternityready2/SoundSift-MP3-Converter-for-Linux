import pathlib

CONFIG_PATH = pathlib.Path(__file__).parent.resolve()
CREDENTIALS = {
    'youtube': [],
    'spotify': [],
}

try:
    with open(CONFIG_PATH / 'spotify-credentials.csv', 'r') as file:
        for line in file.readlines():
            CREDENTIALS['spotify'].append(line.strip().split(','))

except FileNotFoundError:
    with open(CONFIG_PATH / 'spotify-credentials.csv', 'w') as file:
        pass

try:
    with open(CONFIG_PATH / 'youtube-credentials.csv', 'r') as file:
        for line in file.readlines():
            CREDENTIALS['youtube'].append(line.strip().split(','))

except FileNotFoundError:
    with open(CONFIG_PATH / 'youtube-credentials.csv', 'w') as file:
        pass
