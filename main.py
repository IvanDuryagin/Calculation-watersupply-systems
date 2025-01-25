# main.py

import tkinter as tk
from interface import ConsumerCalculator

if __name__ == "__main__":
    root = tk.Tk()
    app = ConsumerCalculator(root)
    root.mainloop()