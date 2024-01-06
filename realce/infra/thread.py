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
            self.parent.msg_tooltip(f'Erro: {e}')
            RealceLogger.get_logger().error(f'Tente novamente. Erro: {e}')
            self.finished.emit(e)

    def update_progress_bar(self, value=None):
        self.parent.progress_bar.setProperty('class', '')
        self.parent.progress_bar.style().polish(self.parent.progress_bar)
        self.parent.progress_bar.setValue(value)

    def handle_thread_finished(self, button):
        self.parent.msg_tooltip('Operação realizada com sucesso')
        button.setEnabled(True)

    def thread_stopped_bar(self):
        self.parent.progress_bar.setProperty('class', 'custom-color')
        self.parent.progress_bar.style().polish(self.parent.progress_bar)
        self.parent.progress_bar.setValue(100)

    def stop_execution(self, button):
        if self.isRunning():
            button.setEnabled(True)
            self.thread_stopped_bar()
            RealceLogger.get_logger().info('Operação cancelada')
            self.terminate()
