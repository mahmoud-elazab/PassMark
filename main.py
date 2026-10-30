"""
PassMarkTool
Main entry point - launches GUI
"""

from gui.app_window import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()