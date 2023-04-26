# coding=utf-8
import time
import logging
from Src.Comm.comm import init_logger
from Src.Core import run
logger = logging.getLogger(__name__)


if __name__ == '__main__':
    init_logger()
    logger.info("main start")
    logger.info("just test")

    run.run()

    logger.info("main end")
