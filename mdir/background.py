"""A bounded daemon queue for latest-only, read-only preview work."""
import threading


class LatestWorker:
    def __init__(self, name):
        self.condition = threading.Condition()
        self.pending = None
        self.closed = False
        self.thread = threading.Thread(target=self._run, name=name, daemon=True)
        self.thread.start()

    def submit(self, callback):
        with self.condition:
            if not self.closed:
                self.pending = callback
                self.condition.notify()

    def cancel(self):
        with self.condition:
            self.pending = None

    def close(self):
        with self.condition:
            self.closed = True
            self.pending = None
            self.condition.notify()

    def _run(self):
        while True:
            with self.condition:
                while self.pending is None and not self.closed:
                    self.condition.wait()
                if self.closed:
                    return
                callback, self.pending = self.pending, None
            try:
                callback()
            except Exception:
                import logging
                logging.getLogger(__name__).exception('Preview background task failed')
