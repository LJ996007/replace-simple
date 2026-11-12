"""Serialized background jobs; Qt signals deliver completion on the GUI thread."""
from threading import Lock
from PySide6.QtCore import QObject, Signal, QRunnable, QThreadPool, QTimer


class JobSignals(QObject):
    done = Signal(object)
    failed = Signal(str)


class Job(QRunnable):
    def __init__(self, work):
        super().__init__()
        self.work, self.signals = work, JobSignals()
        self.lock, self.latest = Lock(), None

    def publish(self, *args):
        with self.lock:
            self.latest = args

    def run(self):
        # Legacy Office automation is apartment-threaded on Windows.
        com = None
        try:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com = pythoncom
            except ImportError:
                pass
            self.signals.done.emit(self.work(self.publish))
        except Exception as exc:
            try:
                from ui_common import write_error_log
                log = write_error_log(type(exc), exc, exc.__traceback__, context='Qt 后台任务')
                message = f'{type(exc).__name__}: {exc}\n错误日志：{log}'
            except Exception:
                message = f'{type(exc).__name__}: {exc}'
            self.signals.failed.emit(message)
        finally:
            if com:
                com.CoUninitialize()


class TaskController(QObject):
    busyChanged = Signal(bool)
    progress = Signal(object)
    error = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.job, self.callback = None, None
        self.pool = QThreadPool(self)
        self.pool.setMaxThreadCount(1)
        self.timer = QTimer(self)
        self.timer.setInterval(80)
        self.timer.timeout.connect(self.poll)

    @property
    def busy(self):
        return self.job is not None

    def submit(self, work, callback):
        if self.busy:
            return False
        self.job, self.callback = Job(work), callback
        self.job.signals.done.connect(self.finish)
        self.job.signals.failed.connect(self.fail)
        self.busyChanged.emit(True)
        self.timer.start()
        self.pool.start(self.job)
        return True

    def poll(self):
        if self.job:
            with self.job.lock:
                latest, self.job.latest = self.job.latest, None
            if latest:
                self.progress.emit(latest)

    def release(self):
        self.poll()
        self.timer.stop()
        callback = self.callback
        self.job, self.callback = None, None
        self.busyChanged.emit(False)
        return callback

    def finish(self, value):
        callback = self.release()
        try:
            callback(value)
        except Exception as exc:
            self.error.emit(str(exc))

    def fail(self, message):
        self.release()
        self.error.emit(message)
