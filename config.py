import os
class Config(object):
    DEBUG = False
    LEWIS_USERNAME = os.environ.get("LEWIS_USERNAME")
    LEWIS_PASSWORD = os.environ.get("LEWIS_PASSWORD")
    MAPFRE_USERNAME = os.environ.get("MAPFRE_USERNAME")
    MAPFRE_PASSWORD = os.environ.get("MAPFRE_PASSWORD")