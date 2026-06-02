from src.utils.env import setup_environment

# Initialize environment and configure Tcl/Tk system paths
setup_environment()

from src.ui.app import App

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
