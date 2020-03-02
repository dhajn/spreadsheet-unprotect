import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import os
from spreadsheet import SpreadsheetReader, SpreadsheetWriter, BadSpreadsheetError


def main():
    root = tk.Tk()
    app = Application(root)
    app.mainloop()


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master, padding="10 10 10 10")               
        self.master = master
        self.master.protocol("WM_DELETE_WINDOW", self.onCloseWindow)
        self.master.title("Spreadsheet Unprotector")
        self.createFrames()
        self.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.addPaddingToAll(self, 5, 5)
        self.fileSelectFrame.butOpen.focus()

    def createFrames(self):
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.fileSelectFrame = FileSelectFrame(master=self)
        self.fileSelectFrame.grid(column=1, row=1, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.columnconfigure(1, weight=1, minsize="250px")

        self.optionsFrame = OptionsFrame(master=self)
        self.optionsFrame.grid(column=1, row=2, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.rowconfigure(2, weight=1)

        self.buttonsFrame = ButtonsFrame(master=self)
        self.buttonsFrame.grid(column=1, row=3, sticky=(tk.E, tk.S))

    def addPaddingToAll(self, master, x, y):
        """
        Add padx and pady to all children of master, except right padding to listboxes and left padding to scrollbars
        """
        for child in master.winfo_children():
            if isinstance(child, tk.Listbox):
                child.grid_configure(padx=(x, 0), pady=(y, y))
            elif isinstance(child, ttk.Scrollbar):
                child.grid_configure(padx=(0, x), pady=(y, y))
            else:
                child.grid_configure(padx=(x, x), pady=(y, y))
            if len(child.winfo_children()) > 0:
                self.addPaddingToAll(child, x, y)

    def onCloseWindow(self):
        self.closeApplication()
    
    def closeApplication(self):
        self.fileSelectFrame.fclose()
        self.master.destroy()

    def unprotectWrite(self):
        reader = self.fileSelectFrame.reader
        if reader is None:
            return
        inName, inExt = os.path.splitext(reader.path)
        if inExt == ".xlsx":
            outType = ("Excel files", ".xlsx")
        elif inExt == ".xlsm":
            outType = ("Excel macro-enabled files", ".xlsm")
        else:
            outType = ("All files", "*.*")
        path = filedialog.asksaveasfilename(initialdir=os.path.dirname(inName), initialfile=(os.path.basename(inName)+"_unprotected"+inExt), filetypes=(outType,))
        if path is None or path == "":
            return
        if os.path.splitext(path)[1] == "":
            path += inExt
        wb, sheetIndexes, dumpVba = self.optionsFrame.getOptions()
        sheets = [reader.protectedSheets[i] for i in sheetIndexes]
        try:
            with SpreadsheetWriter(path) as w:
                w.loadUnprotect(reader, workbook=wb, sheets=sheets, dumpVba=dumpVba)
        except Exception as exc:
            messagebox.showerror(title="Failed to write file", message="Error: " + str(exc))
            return
        messagebox.showinfo(title="Successfully unprotected", message="Unprotected file saved to " + path)


class FileSelectFrame(ttk.Frame):
    def __init__(self, master, **kwargs):
        if kwargs is not None:
            super().__init__(master, **kwargs)
        else:
            super().__init__(master)
        self.master = master
        self.createWidgets()
        self.reader = None

    def createWidgets(self):
        self.textFilePath = tk.StringVar(value="No file opened")
        self.labFilePath = ttk.Label(self, textvariable=self.textFilePath)
        self.labFilePath.grid(column=2, row=1, sticky=tk.W)

        self.butOpen = ttk.Button(self, text="Select file...", command=self.fopen)
        self.butOpen.grid(column=1, row=1, sticky=(tk.N, tk.W))

        self.textInfo = tk.StringVar(value="")
        self.labInfo = ttk.Label(self, textvariable=self.textInfo)
        self.labInfo.grid(column=1, row=2, columnspan=2, sticky=(tk.N, tk.W))

    def fopen(self):
        self.fclose()
        self.reader = None
        self.textFilePath.set("Loading file, please wait...")
        filepath = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xlsm"),))
        if filepath:
            try:
                self.reader = SpreadsheetReader(filepath)
                self.reader.parseWbSheets()
                self.textFilePath.set("Loaded " + filepath)
            except FileNotFoundError:
                messagebox.showerror(title="File not found", message="The file does not exist.")
                self.textFilePath.set("No file opened")
                return
            except BadSpreadsheetError as exc:
                messagebox.showerror(title="Failed to read file", message="Error: " + str(exc) + f"\n" + "The file is either corrupt, encrypted, or not a xlsx/xlsm file.")
                self.textFilePath.set("No file opened")
                return
        else:
            self.textFilePath.set("No file opened")
            self.updateInfo()
        self.updateInfo()
        self.updateOptions()

    def updateOptions(self):
        if self.reader is not None:
            self.master.optionsFrame.updateOptions(wbProtection=self.reader.wbProt, protectedSheets=self.reader.protectedSheets, hasVba=self.reader.hasVba)
        else:
            self.master.optionsFrame.updateOptions()

    def updateInfo(self):
        if self.reader is None:
            self.textInfo.set("")
        else:
            self.textInfo.set("> Found %d protected sheets and %d sheets without protection. Workbook %s protected."
                                % (len(self.reader.protectedSheets), len(self.reader.unprotectedSheets), "is" if self.reader.wbProt else "not"))

    def fclose(self):
        if self.reader is not None:
            self.reader.close()


class OptionsFrame(ttk.Frame):
    def __init__(self, master, **kwargs):
        if kwargs is not None:
            super().__init__(master, **kwargs)
        else:
            super().__init__(master)
        self.master = master
        self.createWidgets()
        self.updateOptions()

    def createWidgets(self):
        self.chvarUnprotectWorkbook = tk.StringVar(value="0")
        self.chboxUnprotectWorkbook = ttk.Checkbutton(self, text="Unprotect workbook", variable=self.chvarUnprotectWorkbook)
        self.chboxUnprotectWorkbook.grid(column=1, row=1, sticky=tk.W)

        self.chvarDumpVba = tk.StringVar(value="0")
        self.chboxDumpVba = ttk.Checkbutton(self, text="Remove VBA macros", variable=self.chvarDumpVba)
        self.chboxDumpVba.grid(column=1, row=2, sticky=tk.W)

        self.chvarUnprotectSheets = tk.StringVar(value="0")
        self.chboxUnprotectSheets = ttk.Checkbutton(self, text="Unprotect sheets", command=self.onChangeChboxUnprotectSheets, variable=self.chvarUnprotectSheets)
        self.chboxUnprotectSheets.grid(column=1, row=3, sticky=tk.W)

        self.lvarProtectedSheets = tk.StringVar()
        self.lboxSheets = tk.Listbox(self, height=10, listvariable=self.lvarProtectedSheets, selectmode="extended")
        self.lboxSheets.grid(column=1, row=4, sticky=(tk.W, tk.N, tk.E, tk.S))

        self.sbarSheets = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.lboxSheets.yview)
        self.lboxSheets.configure(yscrollcommand=self.sbarSheets.set)
        self.sbarSheets.grid(column=2, row=4, sticky=(tk.W, tk.N, tk.S))

        self.selectButtonsFrame = ttk.Frame(self)
        self.selectButtonsFrame.grid(column=3, row=4, sticky=(tk.W, tk.N))

        self.butSelectAll = ttk.Button(self.selectButtonsFrame, text="Select all", command=self.onClickButSelectAll)
        self.butSelectAll.grid(column=1, row=1, sticky=(tk.W, tk.N))

        self.butUnselectAll = ttk.Button(self.selectButtonsFrame, text="Unselect all", command=self.onClickButUnselectAll)
        self.butUnselectAll.grid(column=1, row=2, sticky=(tk.W, tk.N))
        
        self.columnconfigure(1, weight=1)

    def onChangeChboxUnprotectSheets(self):
        if self.chvarUnprotectSheets.get() == "0":
            self.lboxSheets.configure(state=tk.DISABLED)
        else:
            self.lboxSheets.configure(state=tk.NORMAL)

    def onClickButSelectAll(self):
        self.lboxSheets.select_set(0, "end")

    def onClickButUnselectAll(self):
        self.lboxSheets.select_clear(0, "end")

    def updateOptions(self, wbProtection=False, protectedSheets=None, hasVba=False):
        self.chboxUnprotectWorkbook.configure(state=(tk.ACTIVE if wbProtection else tk.DISABLED))
        self.chvarUnprotectWorkbook.set("1" if wbProtection else "0")

        self.chboxDumpVba.configure(state=(tk.ACTIVE if hasVba else tk.DISABLED))

        if protectedSheets is not None and len(protectedSheets) > 0:
            self.lvarProtectedSheets.set([i.name for i in protectedSheets])
            self.lboxSheets.configure(state=tk.NORMAL)
            self.chvarUnprotectSheets.set("1")
            self.chboxUnprotectSheets.configure(state=tk.ACTIVE)
            self.lboxSheets.select_set(0, "end")
        else:
            self.lvarProtectedSheets.set("")
            self.lboxSheets.configure(state=tk.DISABLED)
            self.chvarUnprotectSheets.set("0")
            self.chboxUnprotectSheets.configure(state=tk.DISABLED)

    def getOptions(self):
        wb = True if self.chvarUnprotectWorkbook.get() == "1" else False
        dumpVba = True if self.chvarDumpVba.get() == "1" else False
        if self.chvarUnprotectSheets.get() == "1":
            sheetIndexes = [int(i) for i in self.lboxSheets.curselection()]
        else:
            sheetIndexes = []
        return wb, sheetIndexes, dumpVba



class ButtonsFrame(ttk.Frame):
    def __init__(self, master, **kwargs):
        if kwargs is not None:
            super().__init__(master, **kwargs)
        else:
            super().__init__(master)
        self.master = master
        self.createWidgets()

    def createWidgets(self):
        self.butUnprotect = ttk.Button(self, text="Unprotect...", command=self.onClickButUnprotect)
        self.butUnprotect.grid(column=2, row=1, sticky=(tk.S, tk.W))

        self.butQuit = ttk.Button(self, text="Quit", command=self.master.closeApplication)
        self.butQuit.grid(column=3, row=1, sticky=(tk.S, tk.E))

    def onClickButUnprotect(self):
        self.master.unprotectWrite()

if __name__ == "__main__":
    main()
