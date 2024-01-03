from PySide6.QtCore import QThread, Signal

from realce.infra.logger import RealceLogger


class WorkerThread(QThread):
    progressUpdated = Signal(int)
    finished = Signal(object)

    def __init__(self, function_to_run, *args, priority=QThread.Priority.HighestPriority, parent=None, **kwargs):
        super().__init__(parent)
        self.function_to_run = function_to_run
        self.parent = parent
        self.args = args
        self.kwargs = kwargs
        self.priority = priority

    def run(self):
        try:
            self.currentThread().setPriority(self.priority)
            result = self.function_to_run(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.thread_stopped_bar()
            RealceLogger.get_logger().error(f'Tente novamente. Erro: {e}')
            self.finished.emit(e)

    def update_progress_bar(self, value=None):
        self.parent.progress_bar.setValue(value)

    @staticmethod
    def handle_thread_finished(button):
        button.setEnabled(True)

    def thread_stopped_bar(self):
        self.parent.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid;
                border-radius: 0px;
                border-top: none;
            }

            QProgressBar::chunk {
                background-color: #cc0000;
                border-radius: 0px;
            }
        """)
        self.parent.progress_bar.setValue(100)

    def stop_execution(self, button):
        button.setEnabled(True)
        self.thread_stopped_bar()
        RealceLogger.get_logger().info('Operação cancelada')
        self.terminate()
