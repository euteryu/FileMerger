import customtkinter as ctk
import os

class HeaderButton(ctk.CTkButton):
    """
    A transparent button that acts as a sortable column header.
    """
    def __init__(self, master, text, command, width=100, anchor="center"):
        super().__init__(master, text=text, command=command, width=width, 
                         fg_color="transparent", hover_color="#333333", 
                         anchor=anchor, font=ctk.CTkFont(size=15, weight="bold"))

class FileRowFrame(ctk.CTkFrame):
    """
    A Selectable Row that supports Drag & Drop Reordering.
    """
    def __init__(self, master, index, filepath, metadata, click_callback, drag_start_callback, drag_drop_callback):
        super().__init__(master, fg_color="transparent", corner_radius=6)
        
        self.index = index
        self.filepath = filepath
        self.click_callback = click_callback
        self.drag_start_callback = drag_start_callback
        self.drag_drop_callback = drag_drop_callback
        self.selected = False

        self.grid_columnconfigure(1, weight=1)
        
        # 1. Type Badge
        color_map = {'WORD': '#2980B9', 'PPT': '#C0392B', 'PDF': '#27AE60'}
        badge_color = color_map.get(metadata['type'], 'gray')
        
        self.lbl_type = ctk.CTkLabel(self, text=metadata['type'], width=60, 
                                     fg_color=badge_color, text_color="white",
                                     corner_radius=4, font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_type.grid(row=0, column=0, padx=(10, 5), pady=8)

        # 2. Filename
        self.lbl_name = ctk.CTkLabel(self, text=metadata['name'], anchor="w", 
                                     font=ctk.CTkFont(size=18))
        self.lbl_name.grid(row=0, column=1, padx=5, sticky="ew")

        # 3. Date
        self.lbl_date = ctk.CTkLabel(self, text=metadata['date'], text_color="gray",
                                     font=ctk.CTkFont(size=14))
        self.lbl_date.grid(row=0, column=2, padx=(5, 10))

        # Separator
        self.sep = ctk.CTkFrame(self, height=1, fg_color="#444444")
        self.sep.grid(row=1, column=0, columnspan=3, sticky="ew")

        # Event Binding
        for widget in [self, self.lbl_type, self.lbl_name, self.lbl_date]:
            widget.bind("<Button-1>", self.on_press)
            widget.bind("<ButtonRelease-1>", self.on_release)
            widget.bind("<Enter>", lambda e: self.configure(border_color="#666", border_width=1))
            widget.bind("<Leave>", lambda e: self.configure(border_width=0))

    def on_press(self, event):
        self.drag_start_callback(self.index)
        self.click_callback(event, self.index)

    def on_release(self, event):
        self.drag_drop_callback(event)

    def set_selected(self, selected):
        self.selected = selected
        if selected:
            self.configure(fg_color="#1f538d") # Blue highlight
        else:
            self.configure(fg_color="transparent")