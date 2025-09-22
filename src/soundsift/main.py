import json
import threading
import logging
import logging.config
import customtkinter as ctk
from tkinter import ttk, messagebox

from soundsift.components.app import download_songs as dwl
from soundsift.config import (
    CREDENTIALS,
    CREDENTIALS_PATH,
    LOGGER_CONFIG_PATH,
    LOGS_PATH
)

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

        # ================= API FRAME =================
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

        # Spotify input frame with grid layout for labels and entries
        self.spotify_input_frame = ctk.CTkFrame(self.spotify_frame)
        self.spotify_input_frame.pack(fill="x", padx=5, pady=(5, 10))

        self.spotify_id_label = ctk.CTkLabel(self.spotify_input_frame, text="CLIENT_ID:")
        self.spotify_id_label.grid(row=0, column=0, sticky="w", padx=(0, 5), pady=(0, 2))
        self.spotify_id_entry = ctk.CTkEntry(self.spotify_input_frame, placeholder_text="CLIENT_ID")
        self.spotify_id_entry.grid(row=0, column=1, sticky="ew", pady=(0, 2))

        self.spotify_secret_label = ctk.CTkLabel(self.spotify_input_frame, text="CLIENT_SECRET:")
        self.spotify_secret_label.grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(0, 2))
        self.spotify_secret_entry = ctk.CTkEntry(self.spotify_input_frame, placeholder_text="CLIENT_SECRET")
        self.spotify_secret_entry.grid(row=1, column=1, sticky="ew", pady=(0, 2))

        self.spotify_input_frame.grid_columnconfigure(1, weight=1)

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

        # YouTube input frame with grid layout for label and entry
        self.youtube_input_frame = ctk.CTkFrame(self.youtube_frame)
        self.youtube_input_frame.pack(fill="x", padx=5, pady=(5, 10))

        self.youtube_key_label = ctk.CTkLabel(self.youtube_input_frame, text="API_KEY:")
        self.youtube_key_label.grid(row=0, column=0, sticky="w", padx=(0, 5), pady=(0, 2))
        self.youtube_key_entry = ctk.CTkEntry(self.youtube_input_frame, placeholder_text="API_KEY")
        self.youtube_key_entry.grid(row=0, column=1, sticky="ew", pady=(0, 2))

        self.youtube_input_frame.grid_columnconfigure(1, weight=1)

        self.youtube_add_button = ctk.CTkButton(self.youtube_frame, text="Add Row", command=self.add_youtube_row)
        self.youtube_add_button.pack(fill="x", padx=5, pady=2)

        self.youtube_update_button = ctk.CTkButton(self.youtube_frame, text="Update Row", command=self.update_youtube_row)
        self.youtube_update_button.pack(fill="x", padx=5, pady=2)

        self.youtube_delete_button = ctk.CTkButton(self.youtube_frame, text="Delete Row", command=self.delete_youtube_row)
        self.youtube_delete_button.pack(fill="x", padx=5, pady=(2, 10))

        # ======= Tag Styles for Coloring =======
        self.spotify_tree.tag_configure("success", background="#d4fcd4")     # Green
        self.spotify_tree.tag_configure("failed", background="#f79999")      # Gray
        self.spotify_tree.tag_configure("not-tested", background="#eaeaea")  # Light gray
        self.youtube_tree.tag_configure("success", background="#d4fcd4")
        self.youtube_tree.tag_configure("failed", background="#f79999")
        self.youtube_tree.tag_configure("not-tested", background="#eaeaea")

        # Populate credential trees with color tags according to state
        for (cid, secret, state, *_) in CREDENTIALS['spotify']:
            self.spotify_tree.insert("", "end", values=(cid, secret), tags=(state,))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")

        for (api_key, state, *_) in CREDENTIALS['youtube']:
            self.youtube_tree.insert("", "end", values=(api_key,), tags=(state,))
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
        def process_all_links():
            all_items = self.tree.get_children()
            for item in all_items:
                current_values = self.tree.item(item, "values")
                status, link = current_values

                # Update TreeView for status = "PROCESSING" safely from main thread
                self.after(0, lambda i=item, l=link: self.tree.item(i, values=("PROCESSING", l)))
                index = self.tree.index(item)
                self.storage.update_data(index, status="PROCESSING")

                # Time-consuming processing
                status_result = dwl.appl.download_music_direct(link)

                # Update TreeView with final status safely from main thread
                self.after(0, lambda i=item, s=status_result, l=link: self.tree.item(i, values=(s, l)))
                self.storage.update_data(index, status=status_result)

        threading.Thread(target=process_all_links, daemon=True).start()

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
            self.spotify_tree.insert("", 0, values=(cid, secret), tags=("not-tested",))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")
            CREDENTIALS['spotify'].insert(0, [cid, secret, 'not-tested'])

    def update_spotify_row(self):
        selected = self.spotify_tree.selection()
        cid = self.spotify_id_entry.get()
        secret = self.spotify_secret_entry.get()
        if selected and (cid or secret):
            for item in selected:
                current = self.spotify_tree.item(item, "values")
                new_cid = cid if cid else current[0]
                new_secret = secret if secret else current[1]

                idx = self.spotify_tree.index(item)
                CREDENTIALS['spotify'][idx] = [new_cid, new_secret, "not-tested"]

                self.spotify_tree.item(item, values=(new_cid, new_secret), tags=("not-tested",))
            self.spotify_id_entry.delete(0, "end")
            self.spotify_secret_entry.delete(0, "end")

    def delete_spotify_row(self):
        selected = self.spotify_tree.selection()
        for item in selected:
            idx = self.spotify_tree.index(item)
            CREDENTIALS['spotify'].pop(min(
                idx,
                len(CREDENTIALS['spotify']) - 1
            ))
            self.spotify_tree.delete(item)

    # --- YouTube Methods ---
    def add_youtube_row(self):
        api_key = self.youtube_key_entry.get()
        if api_key:
            self.youtube_tree.insert("", 0, values=(api_key,), tags=("not-tested",))
            self.youtube_key_entry.delete(0, "end")
            CREDENTIALS['youtube'].insert(0, [api_key, 'not-tested'])

    def update_youtube_row(self):
        selected = self.youtube_tree.selection()
        key = self.youtube_key_entry.get()
        if selected and key:
            for item in selected:
                idx = self.youtube_tree.index(item)
                CREDENTIALS['youtube'][idx] = [key, "not-tested"]
                self.youtube_tree.item(item, values=(key,), tags=("not-tested",))
            self.youtube_key_entry.delete(0, "end")

    def delete_youtube_row(self):
        selected = self.youtube_tree.selection()
        for item in selected:
            idx = self.youtube_tree.index(item)
            CREDENTIALS['youtube'].pop(min(
                idx,
                len(CREDENTIALS['youtube']) - 1
            ))
            self.youtube_tree.delete(item)

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            with open(CREDENTIALS_PATH / 'spotify-credentials.csv', 'w') as file:
                for credential in CREDENTIALS['spotify']:
                    file.write(','.join(credential) + '\n')
            with open(CREDENTIALS_PATH / 'youtube-credentials.csv', 'w') as file:
                for credential in CREDENTIALS['youtube']:
                    file.write(','.join(credential) + '\n')
            self.destroy()

def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("green")

    with (open(LOGGER_CONFIG_PATH, 'r', encoding="utf-8")
          as logger_config_fp):
        logger_config = json.load(logger_config_fp)

    logger_config['handlers']['file']['filename'] = LOGS_PATH
    logging.config.dictConfig(logger_config)

    logging.getLogger("yt_dlp").setLevel(logging.ERROR)
    logging.getLogger("spotipy").setLevel(logging.CRITICAL)
    logging.getLogger("spotify_dl").setLevel(logging.CRITICAL)
    logging.getLogger("urllib3.connectionpool").setLevel(logging.CRITICAL)
    
    logger = logging.getLogger(__name__)

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
