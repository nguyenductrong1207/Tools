import pyautogui
import time

def move_mouse_left_right(duration, pause_duration):
    while True:
        # Get the current mouse position
        x, y = pyautogui.position()

        # Move the mouse to the left by 500 pixels
        pyautogui.moveTo(x - 200, y, duration = duration)

        # Move the mouse back to the original position
        pyautogui.moveTo(x, y, duration = duration)

        # Pause before the next movement
        time.sleep(pause_duration)

if __name__ == "__main__":
    try:
        move_mouse_left_right(duration = 5, pause_duration = 30)
    except KeyboardInterrupt:
        print("Program stopped by user.")
