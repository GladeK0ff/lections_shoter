# click_handler.py

class ClickHandler:
    def __init__(self, mode="animation"):
        self.mode = mode
        self.positions = []

    def on_click(self, x, y, button, pressed):
        if pressed:
            self.positions.append((x, y))
            if self.mode == "animation" and len(self.positions) == 2:
                return False
            elif self.mode == "no_animation" and len(self.positions) == 3:
                return False

    def get_positions(self):
        return self.positions
