"""Marshal background results through a queue; only Tk's thread calls Tk."""
import queue


def install_dispatch(widget):
    callbacks = queue.SimpleQueue()
    alive = [True]
    timer = [None]
    def post(callback):
        if alive[0]:
            callbacks.put(callback)
    def poll():
        if not alive[0]:
            return
        for _ in range(100):
            try:
                callback = callbacks.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                widget._root().report_callback_exception(*__import__('sys').exc_info())
        timer[0] = widget.after(30, poll)
    def destroy(event):
        if event.widget is widget:
            alive[0] = False
            if timer[0] is not None:
                widget.after_cancel(timer[0])
    widget.bind('<Destroy>', destroy, add='+')
    timer[0] = widget.after(30, poll)
    return post
