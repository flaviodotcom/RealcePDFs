import logging


def get_logger(name=None):
    return RealceLogger.get_logger(name)


class RealceLogger:
    _logger = None
    logging.basicConfig(level=logging.DEBUG)

    @staticmethod
    def get_logger(name='RealcePDF'):
        if not RealceLogger._logger:
            RealceLogger._logger = logging.getLogger(name)

        return RealceLogger._logger
