import logging
from io import StringIO


def get_logger(name=None):
    return RealceLogger.get_logger(name)


class RealceLogger:
    _logger = None
    console_output = StringIO()
    logging.basicConfig(level=logging.DEBUG)

    @staticmethod
    def get_logger(name='RealcePDF'):
        if not RealceLogger._logger:
            RealceLogger._logger = logging.getLogger(name)

            handler_stdout = logging.StreamHandler(RealceLogger.console_output)
            handler_stdout.setLevel(logging.INFO)

            formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S')
            handler_stdout.setFormatter(formatter)
            RealceLogger._logger.addHandler(handler_stdout)

            handler_stderr = logging.StreamHandler(RealceLogger.console_output)
            handler_stderr.setLevel(logging.ERROR)
            formatter_stderr = logging.Formatter('ERROR: %(message)s')
            handler_stderr.setFormatter(formatter_stderr)
            RealceLogger._logger.addHandler(handler_stderr)

        return RealceLogger._logger
