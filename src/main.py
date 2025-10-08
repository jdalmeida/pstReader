"""
@author João Gbriel de Almeida
"""

import tkinter as tk
from src.ui import AppUI


def main() -> None:
    root = tk.Tk()
    root.title("Leitor de PST")
    root.geometry("1100x700")
    AppUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
