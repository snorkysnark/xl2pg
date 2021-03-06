import json
from typing import NamedTuple, Optional

class DbConfig(NamedTuple):
    dbname: str
    user: Optional[str]
    password: Optional[str]
    host: Optional[str]
    port: Optional[str]

def empty_is_none(value):
    if value != '':
        return value

def get_config_property(config, name):
    value = config.get(name, None)
    return empty_is_none(value)

def load(json_path):
    with open(json_path) as json_file:
        config = json.load(json_file)

        dbname = get_config_property(config, 'dbname')
        if not dbname:
            return None
        return DbConfig(
            dbname=dbname,
            user=get_config_property(config, 'user'),
            password=get_config_property(config, 'password'),
            host=get_config_property(config, 'host'),
            port=get_config_property(config, 'port')
        )

def prompt():
    dbname = ''
    while dbname == '':
        dbname = input('Database Name: ')

    user = empty_is_none(input('User: '))
    password = empty_is_none(input('Password: '))
    host = empty_is_none(input('Host: '))
    port = empty_is_none(input('Port: '))
    return DbConfig(dbname, user, password, host, port)

def load_or_prompt(json_path):
    config = None
    if json_path:
        config = load(json_path)

    if not config:
        config = prompt()

    return config
