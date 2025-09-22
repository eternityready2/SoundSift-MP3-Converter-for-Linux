import pathlib

CONFIG_PATH = pathlib.Path(__file__).parent.resolve()
CREDENTIALS_PATH = pathlib.Path.home() / '.soundsift'
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
        cmap = {
            'not-tested': [],
            'failed': [],
            'success': [],
        }

        for line in file.readlines():
            credential = line.strip().split(',')
            cmap[credential[2]].append(credential)
            
        CREDENTIALS['spotify'] = cmap['not-tested'] + cmap['success'] + cmap['failed']

except FileNotFoundError:
    with open(CREDENTIALS_PATH / 'spotify-credentials.csv', 'w') as file:
        pass

try:
    with open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'r') as file:
        cmap = {
            'not-tested': [],
            'failed': [],
            'success': [],
        }

        for line in file.readlines():
            credential = line.strip().split(',')
            cmap[credential[1]].append(credential)

        CREDENTIALS['youtube'] = cmap['not-tested'] + cmap['success'] + cmap['failed']

except FileNotFoundError:
    with open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'w') as file:
        pass
