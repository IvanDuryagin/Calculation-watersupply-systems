# main.py

import tkinter as tk
from interface import MainApplication

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApplication(root)
    root.eval('tk::PlaceWindow . center')
    root.mainloop()