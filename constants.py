__path = 'const.txt'


def __get_token__():
    with open(__path) as file:
        return file.readline().split('\n')[0]


def __get_proxy__():
    with open(__path) as file:
        file.readline()
        return file.readline().split('\n')[0]


def __get_credentials__():
    with open(__path) as file:
        file.readline()
        file.readline()
        return file.readline().split('\n')[0]


token = __get_token__()
proxy = __get_proxy__()
credentials = __get_credentials__()
