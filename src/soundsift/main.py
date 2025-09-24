import customtkinter as ctk
from tkinter import ttk, PhotoImage, messagebox, filedialog
import threading
import logging
import logging.config
import os
from soundsift.components.app.download_songs import MusicDownloader
import concurrent.futures
import json

from soundsift.config import (
    CREDENTIALS,
    CREDENTIALS_PATH,
    LOGGER_CONFIG_PATH,
    LOGS_PATH
)

class DataStorage:
    def __init__(self):
        self.data = []

    def add_data(self, status, link, message):
        self.data.append((status, link, message))

    def delete_data(self, index):
        if 0 <= index < len(self.data):
            self.data.pop(index)

    def update_data(self, index, status=None, link=None, message=None):
        if 0 <= index < len(self.data):
            current_status, current_link, current_message = self.data[index]
            self.data[index] = (
                status or current_status,
                link or current_link,
                message or current_message,
            )

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.title("SoundSift")
        self.geometry("900x600")
        self.storage = DataStorage()

        self.logger = logging.getLogger(__name__)

        # Set window icon using PNG
        icon_path = "/usr/share/icons/hicolor/256x256/apps/soundsift.png"
        if os.path.exists(icon_path):
            icon = PhotoImage(file=icon_path)
            self.iconphoto(True, icon)
        else:
            self.logger.warning(f"Icon file not found at {icon_path}")

        # Style setup
        self.style = ttk.Style(self)
        self.style.theme_use("default")
        self.style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        self.style.configure("Treeview", rowheight=25, borderwidth=1, relief="solid")
        self.style.configure("Custom.Horizontal.TProgressbar", troughcolor="white", background="#1E90FF")

        # Treeview frame (Left side)
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.tree = ttk.Treeview(
            self.tree_frame, columns=("STATUS", "LINK"), show="headings", selectmode="browse"
        )

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.heading("STATUS", text="STATUS")
        self.tree.heading("LINK", text="LINK")
        self.tree.column("STATUS", width=100, anchor="center")
        self.tree.column("LINK", width=400, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Progress bar setup
        self.progress_var = ctk.IntVar()
        self.progress_bar = ttk.Progressbar(
            self, orient="horizontal", length=500, mode="determinate",
            variable=self.progress_var, style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # Input widgets
        self.link_label = ctk.CTkLabel(self, text="Enter or Update Link:")
        self.link_label.grid(row=2, column=0, sticky="w", padx=10, pady=(10, 0))

        self.link_entry = ctk.CTkEntry(self, placeholder_text="Enter or Update Link")
        self.link_entry.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

        # Buttons frame
        button_frame = ctk.CTkFrame(self)
        button_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)

        self.add_button = ctk.CTkButton(button_frame, text="Add Row", command=self.add_row)
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.process_button = ctk.CTkButton(button_frame, text="Process Links", command=self.process_links)
        self.process_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.update_button = ctk.CTkButton(button_frame, text="Update Link", command=self.update_row)
        self.update_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        self.delete_button = ctk.CTkButton(button_frame, text="Delete Row", command=self.delete_row)
        self.delete_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Configure grid weights on left side
        self.grid_columnconfigure(0, weight=3)
        self.grid_rowconfigure(0, weight=1)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        self.tree.tag_configure("evenrow", background="#f2f2f2")
        self.tree.tag_configure("oddrow", background="#ffffff")

        # Thread-safe progress tracking
        self.total_steps = 0
        self.completed_steps = 0
        self.lock = threading.Lock()

        # Dynamic download path
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")

        # ================= API FRAME (Right side) =================
        self.api_frame = ctk.CTkFrame(self)
        self.api_frame.grid(row=0, column=1, rowspan=5, sticky="nsew", padx=10, pady=10)
        self.api_frame.grid_columnconfigure(0, weight=1)
        self.api_frame.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=2)

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

        self.spotify_export_button = ctk.CTkButton(self.spotify_frame, text="Export Spotify Credential", command=self.export_spotify_credential)
        self.spotify_export_button.pack(fill="x", padx=5, pady=(0, 10))

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

        self.youtube_export_button = ctk.CTkButton(self.youtube_frame, text="Export Credential YouTube", command=self.export_youtube_credential)
        self.youtube_export_button.pack(fill="x", padx=5, pady=(0, 10))

        # ======= Tag Styles for Coloring =======
        self.spotify_tree.tag_configure("success", background="#d4fcd4")     # Green
        self.spotify_tree.tag_configure("failed", background="#f79999")      # Red
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

    def resize_buttons(self, event):
        window_width = self.winfo_width()
        button_width = (window_width - 40) // 2
        self.add_button.configure(width=button_width)
        self.process_button.configure(width=button_width)
        self.update_button.configure(width=button_width)
        self.delete_button.configure(width=button_width)

    def add_row(self):
        link = self.link_entry.get().strip()
        if link:
            tag = "evenrow" if len(self.tree.get_children()) % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=("NEW", link), tags=(tag,))
            self.storage.add_data("NEW", link, "")
            self.link_entry.delete(0, "end")
            self.logger.info(f"Added link: {link}")

    def calculate_total_steps(self, links):
        total_steps = 0
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future_to_link = {executor.submit(MusicDownloader.get_sub_item_count, link, MusicDownloader.identify_source(link)): link for link in links}
            for future in concurrent.futures.as_completed(future_to_link):
                try:
                    sub_items = future.result()
                except Exception as e:
                    link = future_to_link[future]
                    self.logger.error(f"Error fetching sub-item count for {link}: {e}")
                    sub_items = 1
                source = MusicDownloader.identify_source(future_to_link[future])
                if source in ['youtube_playlist', 'spotify_album', 'spotify_playlist']:
                    total_steps += 2 + (sub_items * 3)
                else:
                    total_steps += 5
        total_steps += 1
        return total_steps

    def update_progress(self):
        with self.lock:
            self.completed_steps += 1
            if self.total_steps > 0:
                progress = (self.completed_steps / self.total_steps) * 100
                self.after(0, lambda: self.progress_var.set(progress))

    def process_links(self):
        threading.Thread(target=self._process_links_thread, daemon=True).start()

    def _process_links_thread(self):
        all_items = self.tree.get_children()
        links_to_process = []
        item_map = {}
        for item in all_items:
            status, link = self.tree.item(item, "values")
            if status != "Success":
                links_to_process.append(link)
                item_map[link] = item
            else:
                self.logger.info(f"Skipping already successful link: {link}")

        if not links_to_process:
            self.logger.info("No links to process.")
            return

        self.total_steps = self.calculate_total_steps(links_to_process)
        self.completed_steps = 0
        self.after(0, lambda: self.progress_var.set(0))

        for link in links_to_process:
            item = item_map[link]
            index = self.tree.index(item)
            self.after(0, lambda i=item: self.tree.item(i, values=("Processing", self.tree.item(i, "values")[1])))
            self.storage.update_data(index, status="Processing", message="")
            threading.Thread(target=self.download_link, args=(item, link, index), daemon=True).start()

        self.update_progress()

    def download_link(self, item, link, index):
        def callback(stage):
            if stage in ["start_download", "start_conversion", "end_conversion", "start_link", "end_link"]:
                self.update_progress()

        callback("start_link")
        status, message = MusicDownloader.download_music(link, output_path=self.download_path, callback=callback)
        callback("end_link")
        self.after(0, lambda: self.tree.item(item, values=(status, link)))
        self.storage.update_data(index, status=status, message=message)
        if status == "Success":
            self.logger.info(f"Downloaded {link}: {message}")
        else:
            self.logger.error(f"Failed to download {link}: {message}")

    def update_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            link = self.link_entry.get().strip()
            if link:
                for item in selected_item:
                    self.tree.item(item, values=(self.tree.item(item, "values")[0], link))
                    index = self.tree.index(item)
                    self.storage.update_data(index, link=link, message="")
                self.link_entry.delete(0, "end")
                self.logger.info(f"Updated link to: {link}")

    def delete_row(self):
        selected_items = self.tree.selection()
        if selected_items:
            for item in reversed(selected_items):
                index = self.tree.index(item)
                link = self.tree.item(item, "values")[1]
                self.tree.delete(item)
                self.storage.delete_data(index)
                self.logger.info(f"Deleted link: {link}")

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

    def on_double_click(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return

        index = self.tree.index(selected_item)
        status, link, message = self.storage.data[index]
        messagebox.showinfo(
            title=link,
            message=message,
        )

    def export_spotify_credential(self):
        folder_path = filedialog.askdirectory(
            initialdir=os.path.expanduser("~"),
            title="Select Folder to Export Spotify Credentials"
        )
        if folder_path:
            self.logger.info("Exporting spotify credentials...")
            with open(os.path.join(folder_path, 'spotify-credentials.csv'), 'w') as file:
                for credential in CREDENTIALS['spotify']:
                    file.write(','.join(credential) + '\n')

    def export_youtube_credential(self):
        folder_path = filedialog.askdirectory(
            initialdir=os.path.expanduser("~"),
            title="Select Folder to Export YouTube Credentials",
        )
        if folder_path:
            self.logger.info("Exporting youtube credentials...")
            with open(os.path.join(folder_path, 'youtube-credentials.csv'), 'w') as file:
                for credential in CREDENTIALS['youtube']:
                    file.write(','.join(credential) + '\n')
def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    with (open(LOGGER_CONFIG_PATH, 'r', encoding="utf-8")
          as logger_config_fp):
        logger_config = json.load(logger_config_fp)

    logger_config['handlers']['file']['filename'] = LOGS_PATH
    logging.config.dictConfig(logger_config)

    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
