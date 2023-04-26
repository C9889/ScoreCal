# coding=utf-8
import os
import time
import logging
from configparser import ConfigParser
from Src.Comm.constants import LOG_DIR
from Src.Comm.constants import CONFIG_FILE_NAME


def init_logger():
    log_file_name = LOG_DIR + "/log" + time.strftime("%Y_%m_%d_%H_%M_%S") + ".txt"
    logging.basicConfig(
        filename=log_file_name,
        filemode="w",
        format='%(asctime)s %(filename)s->%(name)s line:%(lineno)d %(levelname)s:%(message)s',
        level=logging.DEBUG
    )


def get_cf_value(section, key):
    if not os.path.exists(CONFIG_FILE_NAME):
        raise FileNotFoundError("配置文件" + CONFIG_FILE_NAME + "不存在")
    cf = ConfigParser()
    cf.read(CONFIG_FILE_NAME, encoding="utf-8")
    return cf.get(section, key).strip()