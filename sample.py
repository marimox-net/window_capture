from window_capture import WindowCapture

# インスタンスの作成
wc = WindowCapture()

# ウィンドウの選択
selected_hwnd, selected_title = wc.select_window()
print(f"選択されたウィンドウ: [{selected_hwnd}] {selected_title}")

image = wc.capture_window_area(selected_hwnd)
image.save("output.png")