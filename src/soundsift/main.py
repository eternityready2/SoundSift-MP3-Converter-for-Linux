import customtkinter as ctk
from tkinter import ttk, messagebox
from soundsift.components.app import download_songs as dwl
from soundsift.config import CREDENTIALS


class DataStorage:
    def __init__(self):
        self.data = []  # Store rows as tuples (status, link)

    def add_data(self, status, link):
        self.data.append((status, link))

    def delete_data(self, index):
        if 0 <= index < len(self.data):
            self.data.pop(index)

    def update_data(self, index, status=None, link=None):
        if 0 <= index < len(self.data):
            current_status, current_link = self.data[index]
            self.data[index] = (status or current_status, link or current_link)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.title("Sound Sift - Easy MP3s")

        self.storage = DataStorage()

        # Configure main grid
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        # ================= MAIN TREEVIEW =================
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("STATUS", "LINK"),
            show="headings"
        )
        self.tree.heading("STATUS", text="STATUS")
        self.tree.heading("LINK", text="LINK")
        self.tree.column("STATUS", width=100, anchor="center")
        self.tree.column("LINK", width=200, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbars for main Tree
        tree_vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        tree_hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")

        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # ================= CONTROLS BELOW MAIN TREE =================
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.controls_frame.grid_columnconfigure(0, weight=1)
        self.controls_frame.grid_columnconfigure(1, weight=1)

        self.link_label = ctk.CTkLabel(self.controls_frame, text="Enter or Update Link:")
        self.link_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(5, 0))

        self.link_entry = ctk.CTkEntry(self.controls_frame, placeholder_text="Enter or Update Link")
        self.link_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        self.add_button = ctk.CTkButton(self.controls_frame, text="Add Row", command=self.add_row)
        self.add_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.process_button = ctk.CTkButton(self.controls_frame, text="Process Links", command=self.confirm_row)
        self.process_button.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        self.update_button = ctk.CTkButton(self.controls_frame, text="Update Link", command=self.update_row)
        self.update_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

        self.delete_button = ctk.CTkButton(self.controls_frame, text="Delete Row", command=self.delete_row)
        self.delete_button.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # ================= API FRAME (SPLIT INTO TWO SUBFRAMES) =================
        self.api_frame = ctk.CTkFrame(self)
        self.api_frame.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=10, pady=10)
        self.api_frame.grid_columnconfigure(0, weight=1)
        self.api_frame.grid_columnconfigure(1, weight=1)

        # --- Spotify Frame ---
        self.spotify_frame = ctk.CTkFrame(self.api_frame)
        self.spotify_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self.spotify_label = ctk.CTkLabel(self.spotify_frame, text="Spotify Credentials")
        self.spotify_label.pack(pady=(5, 5))

        spotify_tree_frame = ctk.CTkFrame(self.spotify_frame)
        spotify_tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.spotify_tree = ttk.Treeview(
            spotify_tree_frame,
            columns=("CLIENT_ID", "CLIENT_SECRET"),
            show="headings"
        )
        self.spotify_tree.heading("CLIENT_ID", text="CLIENT_ID")
        self.spotify_tree.heading("CLIENT_SECRET", text="CLIENT_SECRET")
        self.spotify_tree.column("CLIENT_ID", width=100, anchor="w")
        self.spotify_tree.column("CLIENT_SECRET", width=100, anchor="w")
        self.spotify_tree.grid(row=0, column=0, sticky="nsew")

        spotify_vsb = ttk.Scrollbar(spotify_tree_frame, orient="vertical", command=self.spotify_tree.yview)
        spotify_hsb = ttk.Scrollbar(spotify_tree_frame, orient="horizontal", command=self.spotify_tree.xview)
        self.spotify_tree.configure(yscrollcommand=spotify_vsb.set, xscrollcommand=spotify_hsb.set)
        spotify_vsb.grid(row=0, column=1, sticky="ns")
        spotify_hsb.grid(row=1, column=0, sticky="ew")

        spotify_tree_frame.grid_rowconfigure(0, weight=1)
        spotify_tree_frame.grid_columnconfigure(0, weight=1)

        self.spotify_id_entry = ctk.CTkEntry(self.spotify_frame, placeholder_text="CLIENT_ID")
        self.spotify_id_entry.pack(fill="x", padx=5, pady=(5, 2))

        self.spotify_secret_entry = ctk.CTkEntry(self.spotify_frame, placeholder_text="CLIENT_SECRET")
        self.spotify_secret_entry.pack(fill="x", padx=5, pady=(2, 5))

        self.spotify_add_button = ctk.CTkButton(self.spotify_frame, text="Add Row", command=self.add_spotify_row)
        self.spotify_add_button.pack(fill="x", padx=5, pady=2)

        self.spotify_update_button = ctk.CTkButton(self.spotify_frame, text="Update Row", command=self.update_spotify_row)
        self.spotify_update_button.pack(fill="x", padx=5, pady=2)

        self.spotify_delete_button = ctk.CTkButton(self.spotify_frame, text="Delete Row", command=self.delete_spotify_row)
        self.spotify_delete_button.pack(fill="x", padx=5, pady=(2, 10))

        # --- YouTube Frame ---
        self.youtube_frame = ctk.CTkFrame(self.api_frame)
        self.youtube_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        self.youtube_label = ctk.CTkLabel(self.youtube_frame, text="YouTube Credentials")
        self.youtube_label.pack(pady=(5, 5))

        youtube_tree_frame = ctk.CTkFrame(self.youtube_frame)
        youtube_tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.youtube_tree = ttk.Treeview(
            youtube_tree_frame,
            columns=("API_KEY",),
            show="headings"
        )
        self.youtube_tree.heading("API_KEY", text="API_KEY")
        self.youtube_tree.column("API_KEY", width=150, anchor="w")
        self.youtube_tree.grid(row=0, column=0, sticky="nsew")

        youtube_vsb = ttk.Scrollbar(youtube_tree_frame, orient="vertical", command=self.youtube_tree.yview)
        youtube_hsb = ttk.Scrollbar(youtube_tree_frame, orient="horizontal", command=self.youtube_tree.xview)
        self.youtube_tree.configure(yscrollcommand=youtube_vsb.set, xscrollcommand=youtube_hsb.set)
        youtube_vsb.grid(row=0, column=1, sticky="ns")
        youtube_hsb.grid(row=1, column=0, sticky="ew")

        youtube_tree_frame.grid_rowconfigure(0, weight=1)
        youtube_tree_frame.grid_columnconfigure(0, weight=1)

        self.youtube_key_entry = ctk.CTkEntry(self.youtube_frame, placeholder_text="API_KEY")
        self.youtube_key_entry.pack(fill="x", padx=5, pady=(5, 5))

        self.youtube_add_button = ctk.CTkButton(self.youtube_frame, text="Add Row", command=self.add_youtube_row)
        self.youtube_add_button.pack(fill="x", padx=5, pady=2)

        self.youtube_update_button = ctk.CTkButton(self.youtube_frame, text="Update Row", command=self.update_youtube_row)
        self.youtube_update_button.pack(fill="x", padx=5, pady=2)

        self.youtube_delete_button = ctk.CTkButton(self.youtube_frame, text="Delete Row", command=self.delete_youtube_row)
        self.youtube_delete_button.pack(fill="x", padx=5, pady=(2, 10))

        # Main Tree coloring
        self.tree.tag_configure("evenrow", background="#f2f2f2")
        self.tree.tag_configure("oddrow", background="#ffffff")
        
        for (cid, secret) in CREDENTIALS['spotify']:
            self.spotify_tree.insert("", "end", values=(cid, secret))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")
            
        for (api_key) in CREDENTIALS['youtube']:
            self.youtube_tree.insert("", "end", values=(api_key,))
            self.youtube_key_entry.delete(0, "end")


    # ================= Main Methods =================
    def add_row(self):
        link = self.link_entry.get()
        if link:
            tag = "evenrow" if len(self.tree.get_children()) % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=("NEW", link), tags=(tag,))
            self.storage.add_data("NEW", link)
            self.link_entry.delete(0, "end")

    def confirm_row(self):
        all_items = self.tree.get_children()
        for item in all_items:
            current_values = self.tree.item(item, "values")
            status, link = current_values
            self.tree.item(item, values=("PROCESSING", link))
            index = self.tree.index(item)
            self.storage.update_data(index, status="PROCESSING")
            status = dwl.appl.download_music_direct(link)
            self.tree.item(item, values=(status, link))
            index = self.tree.index(item)
            self.storage.update_data(index, status=status)

    def update_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            link = self.link_entry.get()
            if link:
                for item in selected_item:
                    self.tree.item(item, values=(self.tree.item(item, "values")[0], link))
                    index = self.tree.index(item)
                    self.storage.update_data(index, link=link)
                self.link_entry.delete(0, "end")

    def delete_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            for item in selected_item:
                index = self.tree.index(item)
                self.tree.delete(item)
                self.storage.delete_data(index)

    # --- Spotify Methods ---
    def add_spotify_row(self):
        cid = self.spotify_id_entry.get()
        secret = self.spotify_secret_entry.get()
        if cid and secret:
            self.spotify_tree.insert("", "end", values=(cid, secret))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")

    def update_spotify_row(self):
        selected = self.spotify_tree.selection()
        cid = self.spotify_id_entry.get()
        secret = self.spotify_secret_entry.get()
        if selected and (cid or secret):
            for item in selected:
                current = self.spotify_tree.item(item, "values")
                new_cid = cid if cid else current[0]
                new_secret = secret if secret else current[1]
                self.spotify_tree.item(item, values=(new_cid, new_secret))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")

    def delete_spotify_row(self):
        selected = self.spotify_tree.selection()
        for item in selected:
            self.spotify_tree.delete(item)

    # --- YouTube Methods ---
    def add_youtube_row(self):
        api_key = self.youtube_key_entry.get()
        if api_key:
            self.youtube_tree.insert("", "end", values=(api_key,))
            self.youtube_key_entry.delete(0, "end")
            CREDENTIALS['youtube'].append(api_key)

    def update_youtube_row(self):
        selected = self.youtube_tree.selection()
        key = self.youtube_key_entry.get()
        if selected and key:
            for item in selected:
                self.youtube_tree.item(item, values=(key,))
            self.youtube_key_entry.delete(0, "end")

    def delete_youtube_row(self):
        selected = self.youtube_tree.selection()
        for item in selected:
            self.youtube_tree.delete(item)

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.destroy()

def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("green")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
