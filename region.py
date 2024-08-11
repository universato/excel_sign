import sys
from pynput.mouse import Listener
# pip install pynput

# 機能:座標･サイズを調べたい画面で、左上でクリックし、そのまま右下に移動させて離す。
def on_click(x, y, button, pressed):
    global left, top
    if pressed:
        left = x
        top  = y
    else:
        if x - left < 0:
            x, left = left, x
        if y - top < 0:
            y, top = top, y
        width  = x - left
        height = y - top
        if 32 < width or 32 < height:
            print(f"left   = {left}\ntop    = {top}\nwidth  = {width}\nheight = {height}")
            print(f"region=({left},{top},{width},{height})")
            print(f"bbox=({left},{top},{x},{y})")
            listener.stop()  # Listenerを停止
            sys.exit()       # プログラムを終了

listener = Listener(on_click=on_click)
listener.start()
listener.join()
