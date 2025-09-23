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
    with (
        open(CREDENTIALS_PATH / 'spotify-credentials.csv', 'w') as file1,
        open(SRC_PATH / 'spotify-credentials.csv', 'r') as file2
    ):
        cmap = {
            'not-tested': [],
            'failed': [],
            'success': [],
        }

        for line in file2.readlines():
            file1.write(line)
            credential = line.strip().split(',')
            cmap[credential[2]].append(credential)
            
        CREDENTIALS['spotify'] = cmap['not-tested'] + cmap['success'] + cmap['failed']

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
    with (
        open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'w') as file1,
        open(SRC_PATH / 'youtube-credentials.csv', 'r') as file2
    ):
        cmap = {
            'not-tested': [],
            'failed': [],
            'success': [],
        }

        for line in file2.readlines():
            file1.write(line)
            credential = line.strip().split(',')
            cmap[credential[1]].append(credential)
            
        CREDENTIALS['youtube'] = cmap['not-tested'] + cmap['success'] + cmap['failed']

