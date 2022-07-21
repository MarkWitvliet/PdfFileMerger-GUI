from tkinter import Tk, Text, Menu, Listbox, Entry, messagebox, Toplevel, filedialog, END, BOTH, W, N, E, S
from tkinter.ttk import Frame, Button, Label, Style
from configparser import ConfigParser
from PyPDF2 import PdfFileMerger
from subprocess import run
import sys, os
FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

class PDFMerger(Frame):

    def __init__(self):
        super().__init__()
        PDFMerger.filepaths = list()
        self.parser = ConfigParser()
        self.pdfmpath = os.path.expanduser("~\.pdfmerger")
        self.load_config()
        self.initUI()

    def load_config(self):
        """ Create and load config file at first startup """
        if not os.path.exists(self.pdfmpath):
            try:
                os.mkdir(self.pdfmpath)
            except OSError as e:
                print(f"Creation of directory {self.pdfmpath} failed")
                print(e)
                messagebox.showerror(title="Error!", message=e)
        try:
            with open(f'{self.pdfmpath}\config.ini') as f:
                self.config = self.parser.read_file(f)
        except IOError:
            self.parser.add_section("settings")
            if getattr(sys, 'frozen', False):
                self.parser.set("settings", "inputdir", os.path.dirname(sys.executable))
                self.parser.set("settings", "outputdir", os.path.dirname(sys.executable))
            else:
                self.parser.set("settings", "inputdir", os.path.abspath(os.path.dirname(__file__)))
                self.parser.set("settings", "outputdir", os.path.abspath(os.path.dirname(__file__)))
            self.parser.set("settings", "filename", "merged_pdfs")
            with open(f"{self.pdfmpath}\config.ini", "w") as configfile:
                self.parser.write(configfile)

    def settings(self):
        """ Settings window """
        win = Toplevel(self.master)
        win.title("Settings")
        x = int((self.master.winfo_x()) - (580/2)) + int(self.master.winfo_width()/2)
        y = int((self.master.winfo_y()) - (230/2)) + int(self.master.winfo_height()/2)
        win.geometry(f"580x230+{x}+{y}")
        win.resizable(False,False)
        # filler rows
        emptyLabel1 = Label(win, text=" ")
        emptyLabel1.grid(row=0, column=0)
        emptyLabel2 = Label(win, text=" ")
        emptyLabel2.grid(row=1, column=0)
        # input dir
        def seldir(func):
            new_dir = filedialog.askdirectory()
            if func == "input":
                inputDirEntry.delete(0, 'end')
                inputDirEntry.insert(END, new_dir)
            elif func == "output":
                outputDirEntry.delete(0, 'end')
                outputDirEntry.insert(END, new_dir)
            else:
                raise Exception("Unknown button pressed in settings")
        inputDirLabel = Label(win, text="Input directory")
        inputDirLabel.grid(row=2, column=0)
        inputDirEntry = Entry(win, width=60, borderwidth=2)
        inputDirEntry.insert(END, self.parser.get("settings", "inputdir"))
        inputDirEntry.grid(row=2, column=1, padx=10, pady=10)
        inputdir_btn = Button(win, width=2 ,text="...", command=lambda: seldir("input"))
        inputdir_btn.grid(row=2,column=2, sticky="W")
        # output dir
        outputDirLabel = Label(win, text="Output directory")
        outputDirLabel.grid(row=3, column=0)
        outputDirEntry = Entry(win, width=60, borderwidth=2)
        outputDirEntry.insert(END, self.parser.get("settings", "outputdir"))
        outputDirEntry.grid(row=3, column=1, padx=10, pady=10)
        outputdir_btn = Button(win, width=2 ,text="...", command=lambda: seldir("output"))
        outputdir_btn.grid(row=3,column=2, sticky="W")
        # default output name
        outputNameLabel = Label(win, text="Default outputname")
        outputNameLabel.grid(row=4, column=0)
        outputNameEntry = Entry(win, width=60, borderwidth=2)
        outputNameEntry.insert(END, self.parser.get("settings", "filename"))
        outputNameEntry.grid(row=4, column=1, padx=10, pady=10, sticky=W)
        #filler row
        emptyLabel3 = Label(win, text=" ")
        emptyLabel3.grid(row=5, column=0)
        # save button
        def save():
            self.parser.set("settings", "inputdir", inputDirEntry.get())
            self.parser.set("settings", "outputdir", outputDirEntry.get())
            self.parser.set("settings", "filename", outputNameEntry.get())
            self.filenameEntry.delete(0, 'end')
            self.filenameEntry.insert(END, outputNameEntry.get())
            with open(f"{self.pdfmpath}\config.ini", "w") as configfile:
                self.parser.write(configfile)
            win.destroy()
        save_btn = Button(win, text="Save", command=save)
        save_btn.grid(row=6,column=1, sticky="SE")
        # cancel button
        cancel_btn = Button(win, text="Cancel", command=win.destroy)
        cancel_btn.grid(row=6,column=2, sticky="SE")
        # place settings window on top
        win.transient(self.master)
        win.grab_set()
        win.master.wait_window(win)

    def about(self):
        """ About window """
        win_about = Toplevel(self.master)
        win_about.title("About")
        x = int((self.master.winfo_x()) - (200/2)) + int(self.master.winfo_width()/2)
        y = int((self.master.winfo_y()) - (100/2)) + int(self.master.winfo_height()/2)
        win_about.geometry(f"200x100+{x}+{y}")
        win_about.resizable(False,False)
        # labels
        spacerLabel= Label(win_about, text="")
        aboutLabel1 = Label(win_about, text="PDF Merger", justify="center", font="Helvetica 11 bold")
        aboutLabel2 = Label(win_about, text="Version: 0.2\nAuthor: Mark Witvliet", justify="center", font="Helvetica 10")
        spacerLabel.pack()
        aboutLabel1.pack()
        aboutLabel2.pack()
        # place about window on top
        win_about.transient(self.master)
        win_about.grab_set()

    def open(self):
        """ Open files, insert in listbox"""
        filepaths = filedialog.askopenfilenames(initialdir=self.parser.get("settings", "inputdir"), title="Select a pdf file", filetypes=(("pdf files","*.pdf"), ("all files","*.*")))
        PDFMerger.filepaths = [*PDFMerger.filepaths, *filepaths]
        filenames = []
        for filename in filepaths:
            new_filename = filename.split("/")[-1]
            filenames.append(new_filename)
            pos = len(PDFMerger.filepaths) - 1
            PDFMerger.pdf_listbox.insert(pos, new_filename)

    def delete(self):
        """ Delete item in listbox """
        try:
            idxs = PDFMerger.pdf_listbox.curselection()
            if not idxs:
                return
            for pos in idxs:
                text = PDFMerger.pdf_listbox.get(pos)
                PDFMerger.pdf_listbox.delete(pos)
                PDFMerger.filepaths.pop(pos)
        except Exception as e:
            messagebox.showerror(title="Error!", message=e)

    def clear(self):
        """ Clear all items in listbox """
        PDFMerger.pdf_listbox.delete(0,END)
        PDFMerger.filepaths = list()

    def merge(self):
        """ Merge multiple pdf's """
        pdfs = PDFMerger.filepaths
        if len(pdfs) <= 1:
            prompt = messagebox.showwarning(title = "Warning!",
                                message = "Need at least 2 PDFs to merge")
            return
        filename = self.filenameEntry.get()
        chars = set('<>:"/\|?*')
        if any((c in chars) for c in filename):
            prompt = messagebox.showwarning(title = "Warning!",
                                message = 'A file name can\'t contain any of the following characters:\n \ / : * ? " < > |')
            return
        try:
            merger = PdfFileMerger()

            for pdf in pdfs:
                merger.append(pdf)

            def uniquify(upath):
                filename, extension = os.path.splitext(upath)
                counter = 1
                while os.path.exists(upath):
                    upath = filename + " (" + str(counter) + ")" + extension
                    counter += 1
                return upath

            outputdir = self.parser.get("settings", "outputdir")
            path = f'{outputdir}\{filename}.pdf'
            if os.path.exists(outputdir):
                print(path)
                sdir = uniquify(path)
                print(sdir)
            else:
                prompt = messagebox.showerror(title = "Error!", message = "Path does not exist")
            merger.write(sdir)
            merger.close()
        except Exception as e:
            messagebox.showerror(title="Error!", message=e)
        self.filenameEntry.delete(0, 'end')
        self.filenameEntry.insert(END, self.parser.get("settings", "filename"))
        self.clear()
        self.explore(sdir)

    def move_up(self, *args):
        """ Moves the item at position pos up by one """
        try:
            idxs = PDFMerger.pdf_listbox.curselection()
            if not idxs:
                return
            for pos in idxs:
                if pos == 0:
                    return
                text = PDFMerger.pdf_listbox.get(pos)
                PDFMerger.pdf_listbox.delete(pos)
                PDFMerger.pdf_listbox.insert(pos-1, text)
                PDFMerger.pdf_listbox.selection_set(pos-1)
                PDFMerger.filepaths.insert(pos-1, PDFMerger.filepaths.pop(pos))
        except Exception as e:
            messagebox.showerror(title="Error!", message=e)

    def move_down(self, *args):
        """ Moves the item at position pos down by one """
        try:
            idxs = PDFMerger.pdf_listbox.curselection()
            if not idxs:
                return
            for pos in idxs:
                if pos == int(PDFMerger.pdf_listbox.size()) - 1:
                    return
                text = PDFMerger.pdf_listbox.get(pos)
                PDFMerger.pdf_listbox.delete(pos)
                PDFMerger.pdf_listbox.insert(pos+1, text)
                PDFMerger.pdf_listbox.selection_set(pos+1)
                PDFMerger.filepaths.insert(pos+1, PDFMerger.filepaths.pop(pos))
        except Exception as e:
            messagebox.showerror(title="Error!", message=e)

    def explore(self, path):
        """ explore path """
        # explorer will throw error on forward slashes
        path = os.path.normpath(path)

        if os.path.isdir(path):
            run([FILEBROWSER_PATH, path])
        elif os.path.isfile(path):
            run([FILEBROWSER_PATH, '/select,', path])

    def initUI(self):
        """ Main UI window """
        self.master.title("PDF Merger")
        self.pack(fill=BOTH, expand=True)

        self.columnconfigure(1, weight=1)
        self.columnconfigure(3, pad=7)
        self.rowconfigure(4, weight=1)
        self.rowconfigure(7, pad=1)

        lbl = Label(self, text="Selected PDFs")
        lbl.grid(sticky=W, pady=4, padx=5)

        PDFMerger.pdf_listbox = Listbox(self)
        PDFMerger.pdf_listbox.grid(row=1, column=0, columnspan=2, rowspan=4,
            padx=5, sticky=E+W+S+N)

        obtn = Button(self, text="Open PDF(s)", command=self.open)
        obtn.grid(row=1, column=3, pady=4)

        ubtn = Button(self, text="Up", command=self.move_up)
        ubtn.grid(row=2, column=3, pady=4)

        dbtn = Button(self, text="Down", command=self.move_down)
        dbtn.grid(row=3, column=3, pady=4)

        delbtn = Button(self, text="Delete", command=self.delete)
        delbtn.grid(row=4, column=3, pady=4, sticky=N)

        hbtn = Button(self, text="Clear", command=self.clear)
        hbtn.grid(row=5, column=0, pady=4, padx=5)

        mbtn = Button(self, text="Merge", command=self.merge)
        mbtn.grid(row=5, column=3, pady=4)

        self.filenameEntry = Entry(self, width=100, borderwidth=2)
        self.filenameEntry.insert(END, self.parser.get("settings", "filename"))
        self.filenameEntry.grid(row=5, column=1, padx=10, pady=10)

        # Menu
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.open)
        filemenu.add_command(label="Settings", command=self.settings)
        filemenu.add_command(label="About", command=self.about)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.master.quit)
        menubar.add_cascade(label="Menu", menu=filemenu)

def main():
    root = Tk()
    x = int((root.winfo_screenwidth()/2) - (600/2))
    y = int((root.winfo_screenheight()/2) - (400/2))
    root.geometry(f"600x400+{x}+{y}")
    app = PDFMerger()
    root.mainloop()

if __name__ == '__main__':
    main()
