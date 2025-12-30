import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

# Third party
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES

# Local Modules
from merger_core import MergeWorker, get_file_metadata, is_valid_file
from ui_components import FileRowFrame, HeaderButton

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# --- ASSET PATH HELPER (CRITICAL FOR EXE) ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class FileMergerApp(TkinterDnD_CTk):
    def __init__(self):
        super().__init__()

        self.title("FileMerger")
        self.geometry("1100x750")

        # --- LOAD CUSTOM ICON ---
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path) # Sets window and taskbar icon
        
        # Data
        self.file_data = [] 
        self.row_widgets = [] 
        self.selected_indices = set()
        self.last_clicked_index = None
        
        self.sort_state = {"col": None, "reverse": False}
        self.drag_source_index = None

        self._init_ui()
        self._bind_events()

    def _init_ui(self):
        self.grid_rowconfigure(1, weight=1) 
        self.grid_columnconfigure(1, weight=1)

        # --- SIDEBAR CONTROLS ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=3, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="FileMerger", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(30,5))
        
        font_btn = ctk.CTkFont(size=15)
        
        ctk.CTkButton(self.sidebar, text="Add Files...", font=font_btn, command=self.dialog_add_files).pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="Add Folder...", font=font_btn, command=self.dialog_add_folder).pack(pady=5, padx=20, fill="x")
        
        ctk.CTkFrame(self.sidebar, height=2, fg_color="gray30").pack(pady=20, padx=20, fill="x")

        ctk.CTkLabel(self.sidebar, text="Reorder", text_color="gray", anchor="w", font=ctk.CTkFont(size=14)).pack(padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="⭡ Top", font=font_btn, fg_color="#333", hover_color="#444", command=self.move_top).pack(pady=2, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="▲ Up", font=font_btn, fg_color="transparent", border_width=1, command=self.move_up).pack(pady=2, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="▼ Down", font=font_btn, fg_color="transparent", border_width=1, command=self.move_down).pack(pady=2, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="⭣ Bottom", font=font_btn, fg_color="#333", hover_color="#444", command=self.move_bottom).pack(pady=2, padx=20, fill="x")
        
        ctk.CTkButton(self.sidebar, text="✕ Remove", font=font_btn, fg_color="#922B21", hover_color="#C0392B", command=self.delete_selected).pack(pady=20, padx=20, fill="x")

        # --- RIGHT MAIN AREA ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="#181818") 
        self.scroll_frame.grid(row=1, column=1, sticky="nsew", padx=20, pady=5)
        
        self.scroll_frame.drop_target_register(DND_FILES)
        self.scroll_frame.dnd_bind('<<Drop>>', self.handle_external_drop)

        # --- HEADER (Sortable) ---
        self.header = ctk.CTkFrame(self, height=40, fg_color="transparent")
        self.header.grid(row=0, column=1, sticky="ew", padx=20, pady=(10,0))
        self.header.grid_columnconfigure(1, weight=1)
        
        self.btn_head_type = HeaderButton(self.header, text="TYPE", width=60, command=lambda: self.sort_files("type"))
        self.btn_head_type.grid(row=0, column=0, padx=10)
        
        self.btn_head_name = HeaderButton(self.header, text="FILE NAME", anchor="w", command=lambda: self.sort_files("name"))
        self.btn_head_name.grid(row=0, column=1, sticky="ew", padx=5)
        
        self.btn_head_date = HeaderButton(self.header, text="DATE", width=120, command=lambda: self.sort_files("timestamp"))
        self.btn_head_date.grid(row=0, column=2, padx=10)

        # --- FOOTER ---
        self.footer = ctk.CTkFrame(self, height=60, fg_color="transparent")
        self.footer.grid(row=2, column=1, sticky="ew", padx=20, pady=20)
        self.footer.grid_columnconfigure(0, weight=1)
        
        self.footer.drop_target_register(DND_FILES)
        self.footer.dnd_bind('<<Drop>>', self.handle_external_drop)

        self.progressbar = ctk.CTkProgressBar(self.footer, height=15)
        self.progressbar.grid(row=0, column=0, sticky="ew", padx=(0, 20))
        self.progressbar.set(0)

        self.btn_convert = ctk.CTkButton(self.footer, text="CONVERT & MERGE PDF", height=50, 
                                         font=ctk.CTkFont(size=18, weight="bold"),
                                         fg_color="#27AE60", hover_color="#1E8449",
                                         command=self.run_merge)
        self.btn_convert.grid(row=0, column=1, padx=0)

        self.lbl_status = ctk.CTkLabel(self, text="Ready. Drag & Drop files or sort/reorder list.", 
                                       text_color="gray", font=ctk.CTkFont(size=14))
        self.lbl_status.grid(row=3, column=1, sticky="w", padx=25, pady=(0,10))

    def _bind_events(self):
        self.bind("<Delete>", lambda e: self.delete_selected())
        self.bind("<Control-a>", self.select_all) 
        self.bind("<Command-a>", self.select_all)

    # --- UI RENDERING ---
    
    def refresh_list_ui(self, full_redraw=False):
        if full_redraw:
            for widget in self.row_widgets:
                widget.destroy()
            self.row_widgets = []
            
            for i, item in enumerate(self.file_data):
                row = FileRowFrame(self.scroll_frame, i, item['path'], item['metadata'], 
                                   self.handle_click, self.handle_drag_start, self.handle_drag_drop)
                row.pack(fill="x", pady=2)
                self.row_widgets.append(row)
        
        for i, widget in enumerate(self.row_widgets):
            widget.index = i 
            widget.set_selected(i in self.selected_indices)

    # --- SELECTION & SHORTCUTS ---

    def select_all(self, event=None):
        if not self.file_data: return
        self.selected_indices = set(range(len(self.file_data)))
        self.refresh_list_ui(full_redraw=False)

    def handle_click(self, event, index):
        ctrl = (event.state & 0x0004) != 0
        shift = (event.state & 0x0001) != 0

        if shift and self.last_clicked_index is not None:
            start, end = min(self.last_clicked_index, index), max(self.last_clicked_index, index)
            self.selected_indices = set(range(start, end + 1))
        elif ctrl:
            if index in self.selected_indices: self.selected_indices.remove(index)
            else: self.selected_indices.add(index)
            self.last_clicked_index = index
        else:
            self.selected_indices = {index}
            self.last_clicked_index = index

        self.refresh_list_ui(full_redraw=False)

    # --- SORTING & DRAG REORDER ---

    def sort_files(self, key):
        if self.sort_state["col"] == key:
            self.sort_state["reverse"] = not self.sort_state["reverse"]
        else:
            self.sort_state["col"] = key
            self.sort_state["reverse"] = False 
            
        arrow = " ▼" if self.sort_state["reverse"] else " ▲"
        self.btn_head_type.configure(text="TYPE" + (arrow if key=="type" else ""))
        self.btn_head_name.configure(text="FILE NAME" + (arrow if key=="name" else ""))
        self.btn_head_date.configure(text="DATE" + (arrow if key=="timestamp" else ""))

        self.file_data.sort(key=lambda x: x['metadata'].get(key, ""), reverse=self.sort_state["reverse"])
        self.selected_indices.clear()
        self.refresh_list_ui(full_redraw=True)

    def handle_drag_start(self, index):
        self.drag_source_index = index

    def handle_drag_drop(self, event):
        if self.drag_source_index is None: return
        x, y = self.winfo_pointerxy()
        target_widget = self.winfo_containing(x, y)

        found_index = None
        current = target_widget
        while current:
            if isinstance(current, FileRowFrame):
                found_index = current.index
                break
            try: current = current.master
            except: break
        
        if found_index is not None and found_index != self.drag_source_index:
            item = self.file_data.pop(self.drag_source_index)
            self.file_data.insert(found_index, item)
            self.selected_indices = {found_index}
            self.refresh_list_ui(full_redraw=True)
        
        self.drag_source_index = None

    # --- MOVEMENT LOGIC ---
    def move_up(self):
        if not self.selected_indices: return
        sorted_indices = sorted(self.selected_indices)
        if sorted_indices[0] == 0: return
        new_selection = set()
        for i in sorted_indices:
            self.file_data[i], self.file_data[i-1] = self.file_data[i-1], self.file_data[i]
            new_selection.add(i-1)
        self.selected_indices = new_selection
        self.refresh_list_ui(full_redraw=True)

    def move_down(self):
        if not self.selected_indices: return
        sorted_indices = sorted(self.selected_indices, reverse=True)
        if sorted_indices[0] == len(self.file_data) - 1: return
        new_selection = set()
        for i in sorted_indices:
            self.file_data[i], self.file_data[i+1] = self.file_data[i+1], self.file_data[i]
            new_selection.add(i+1)
        self.selected_indices = new_selection
        self.refresh_list_ui(full_redraw=True)

    def move_top(self):
        if not self.selected_indices: return
        items = [self.file_data[i] for i in sorted(self.selected_indices)]
        for i in sorted(self.selected_indices, reverse=True): self.file_data.pop(i)
        for item in reversed(items): self.file_data.insert(0, item)
        self.selected_indices = set(range(len(items)))
        self.refresh_list_ui(full_redraw=True)

    def move_bottom(self):
        if not self.selected_indices: return
        items = [self.file_data[i] for i in sorted(self.selected_indices)]
        for i in sorted(self.selected_indices, reverse=True): self.file_data.pop(i)
        for item in items: self.file_data.append(item)
        start = len(self.file_data) - len(items)
        self.selected_indices = set(range(start, len(self.file_data)))
        self.refresh_list_ui(full_redraw=True)

    def delete_selected(self):
        if not self.selected_indices: return
        for i in sorted(self.selected_indices, reverse=True):
            if i < len(self.file_data): self.file_data.pop(i)
        self.selected_indices.clear()
        self.refresh_list_ui(full_redraw=True)

    # --- FILE ADDING ---
    def handle_external_drop(self, event):
        raw = event.data
        if raw.startswith('{') and raw.endswith('}'): paths = [raw[1:-1]]
        else: paths = self.tk.splitlist(raw)
        self.add_files(paths)

    def add_files(self, paths):
        added = 0
        for path in paths:
            path = os.path.normpath(path)
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        full = os.path.join(root, f)
                        if is_valid_file(full) and not any(d['path'] == full for d in self.file_data):
                            self.file_data.append({'path': full, 'metadata': get_file_metadata(full)})
                            added += 1
            else:
                if is_valid_file(path) and not any(d['path'] == path for d in self.file_data):
                    self.file_data.append({'path': path, 'metadata': get_file_metadata(path)})
                    added += 1
        if added > 0: self.refresh_list_ui(full_redraw=True)

    def dialog_add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Files", "*.pdf *.pptx *.docx")])
        if files: self.add_files(files)

    def dialog_add_folder(self):
        folder = filedialog.askdirectory()
        if folder: self.add_files([folder])

    def run_merge(self):
        if not self.file_data: return messagebox.showwarning("Empty", "No files.")
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not output: return

        self.btn_convert.configure(state="disabled")
        self.progressbar.set(0)
        worker = MergeWorker(self.file_data, output, self.update_progress, self.update_status, self.worker_done)
        worker.start()

    def update_progress(self, val): self.progressbar.set(val)
    def update_status(self, msg): self.lbl_status.configure(text=msg)
    def worker_done(self, success, message):
        self.btn_convert.configure(state="normal")
        self.progressbar.set(1)
        if success: messagebox.showinfo("Success", message)
        else: messagebox.showerror("Error", message)

if __name__ == "__main__":
    app = FileMergerApp()
    app.mainloop()