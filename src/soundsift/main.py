import customtkinter as ctk
from tkinter import ttk, PhotoImage
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
        self.title("SoundSift")
        self.geometry("600x400")
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

        # Treeview frame
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        self.tree = ttk.Treeview(
            self.tree_frame, columns=("STATUS", "LINK"), show="headings", selectmode="browse"
        )
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
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        # Input widgets
        self.link_label = ctk.CTkLabel(self, text="Enter or Update Link:")
        self.link_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 0))

        self.link_entry = ctk.CTkEntry(self, placeholder_text="Enter or Update Link")
        self.link_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        # Buttons frame
        button_frame = ctk.CTkFrame(self)
        button_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        self.add_button = ctk.CTkButton(button_frame, text="Add Row", command=self.add_row)
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.process_button = ctk.CTkButton(button_frame, text="Process Links", command=self.process_links)
        self.process_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.update_button = ctk.CTkButton(button_frame, text="Update Link", command=self.update_row)
        self.update_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        self.delete_button = ctk.CTkButton(button_frame, text="Delete Row", command=self.delete_row)
        self.delete_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        self.bind("<Configure>", self.resize_buttons)

        self.tree.tag_configure("evenrow", background="#f2f2f2")
        self.tree.tag_configure("oddrow", background="#ffffff")

        # Thread-safe progress tracking
        self.total_steps = 0
        self.completed_steps = 0
        self.lock = threading.Lock()

        # Dynamic download path for the logged-in user
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")

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
            self.storage.add_data("NEW", link)
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
            self.storage.update_data(index, status="Processing")
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
        self.storage.update_data(index, status=status)
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
                    self.storage.update_data(index, link=link)
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
