import time
from threading import Thread, Event
import win32gui, win32con


class MsgBoxListener(Thread):

    def __init__(self, interval: int):
        Thread.__init__(self)
        self._interval = interval
        self._stop_event = Event()

    def stop(self): self._stop_event.set()

    @property
    def is_running(self): return not self._stop_event.is_set()

    def run(self):
        while self.is_running:
            try:
                time.sleep(self._interval)
                self._close_msgbox()
            except Exception as e:
                print(e, flush=True)

    def _close_msgbox(self):
        '''
        Click any button ("OK", "Yes" or "Confirm") to close message box.
        '''
        # get handles of all top windows
        h_windows = []
        win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), h_windows) 
    
        # check each window    
        for h_window in h_windows:            
            # get child button with text OK, Yes or Confirm of given window
            h_btn = win32gui.FindWindowEx(h_window, None,'Button', None)
            if not h_btn: continue
    
            # check button text
            text = win32gui.GetWindowText(h_btn)
            if not text.strip().lower() in ('确定', 'ok', 'yes', 'confirm'): continue
    
            # click button
            win32gui.PostMessage(h_btn, win32con.WM_LBUTTONDOWN, None, None)
            time.sleep(0.2)
            win32gui.PostMessage(h_btn, win32con.WM_LBUTTONUP, None, None)
            time.sleep(0.2)


if __name__ == '__main__':
    t = MsgBoxListener(2)
    t.start()
    time.sleep(10)
    t.stop()
