# gui for pptx input
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
from tkinter import messagebox
from ttkthemes import ThemedStyle
from _pptx import ExtractPPTX
import unicodedata


class Gui(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    """
    Creates buttons in frame
    """

    def create_widgets(self):
        self.hi_there = ttk.Button(self, style="C.TButton")
        self.hi_there["text"] = "Select powerpoint file"
        self.hi_there["cursor"] = "hand2"
        self.hi_there["command"] = self.open_dialog
        self.hi_there.pack(side="top", pady=5, padx=5)

        self.quit = ttk.Button(
            self,
            text="CLOSE",
            cursor="hand2",
            style="C.TButton",
            command=self.master.destroy,
        )
        self.quit.pack(side="bottom", pady=5)

    """
    This is the button handler for pptx button
    """

    def open_dialog(self):
        self.select_file()

    """
    This is the action being called for the button
    """

    def select_file(self):

        self.name = filedialog.askopenfilename(
            initialdir="/",
            title="Select File",
            filetypes=(("Power Point Files", "*.pptx"), ("All files", "*")),
        )

        launch = ExtractPPTX(self.name)

        ppt_text = launch.function()

        length = len(ppt_text)

        normalized_list = []

        f = open("anki.txt", "w+", encoding="utf-8")

        for t in ppt_text:
            f.write(t)
        f.close()

        messagebox.showinfo("Finished", "Finished converting file")

        self.master.destroy()


root = tk.Tk()

"""
Window positioning for center
"""
window_height = root.winfo_reqheight()
window_width = root.winfo_reqwidth()
position_right = int(root.winfo_screenwidth() / 2 - window_width / 2)
position_down = int(root.winfo_screenheight() / 2 - window_height / 2)
root.geometry("+%d+%d" % (position_right, position_down))
root.resizable(0, 0)

app = Gui(master=root)

"""
Main theme settings
"""
app.master.title("PPTX Anki Card Maker")
style = ThemedStyle(root)
style.set_theme("equilux")
stye_name = style.theme_use()
# root.overrideredirect(True)
s = ttk.Style()
s.configure(root, background="#888")

"""
Styles for buttons
"""
mbtn_style = ttk.Style()
mbtn_style.configure("TButton", foreground="#fff", font=("Helevetica", "12", "bold"))
mbtn_style.map("C.TButton", foreground=[("active", "#bada55")])

app.mainloop()
