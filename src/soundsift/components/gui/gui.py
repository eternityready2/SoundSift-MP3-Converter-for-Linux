
import os
import logging
from customtkinter import CTk, CTkEntry, CTkButton, CTkFrame
from tkinter import ttk, Scrollbar
from soundsift.components.app.download_songs import MusicDownloader
import threading

class DownloaderGUI(CTk):
	"""A GUI for SoundSift."""
	def __init__(self):
		super().__init__()

		# Window setup
		self.title("SoundSift")
		self.minsize(width=1000, height=600)
		self.geometry("1000x600")
		self.iconbitmap(os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico"))  # Adjust path as needed

		# Logging setup
		self.logger = logging.getLogger("soundsift")

		# Style setup
		self.style = ttk.Style()
		self.style.theme_use("default")
		self.style.configure("Treeview",
							 background="#D3D3D3",
							 foreground="black",
							 rowheight=25,
							 fieldbackground="#D3D3D3")
		self.style.map("Treeview", background=[("selected", "#347083")])

		# URL entry and download button
		self.input_frame = CTkFrame(self)
		self.input_frame.pack(fill="x", padx=20, pady=10)

		self.url_entry = CTkEntry(self.input_frame, placeholder_text="Enter Spotify or YouTube URL", width=800)
		self.url_entry.pack(side="left", padx=5)

		# Button with dynamic width
		self.download_btn = CTkButton(self.input_frame, text="Download", command=self.download_url)
		self.download_btn.pack(side="left", padx=5)

		# Treeview frame for download status
		tree_frame = CTkFrame(self)
		tree_frame.pack(fill="both", expand=True, padx=20, pady=10)

		self.tree = ttk.Treeview(tree_frame, columns=("URL", "Source", "Status"), show="headings", selectmode="extended")
		self.tree.heading("URL", text="URL")
		self.tree.heading("Source", text="Source")
		self.tree.heading("Status", text="Status")
		self.tree.column("URL", width=600)
		self.tree.column("Source", width=150)
		self.tree.column("Status", width=150)
		self.tree.pack(side="left", fill="both", expand=True)

		scrollbar = Scrollbar(tree_frame, command=self.tree.yview)
		scrollbar.pack(side="right", fill="y")
		self.tree.configure(yscrollcommand=scrollbar.set)

		# Bind resize event to adjust button size
		self.bind("<Configure>", self.resize_buttons)

	def resize_buttons(self, event):
		"""Adjust button size dynamically based on window width."""
		window_width = self.winfo_width()
		button_width = (window_width - 850) // 2  # Subtract entry width (800) and padding
		if button_width > 0:  # Ensure positive width
			self.download_btn.configure(width=button_width)

	def download_url(self):
		"""Handle the download button click."""
		url = self.url_entry.get().strip()
		if not url:
			return

		# Add to Treeview
		source = MusicDownloader.identify_source(url)
		item_id = self.tree.insert("", "end", values=(url, source, "Pending"))

		# Start download in a thread
		threading.Thread(target=self.process_download, args=(item_id, url), daemon=True).start()

	def process_download(self, item_id, url):
		status, message = MusicDownloader.download_music(url)
		self.tree.item(item_id, values=(url, MusicDownloader.identify_source(url), status))
		if ENABLE_LOGGING:
			logger = logging.getLogger("soundsift")
			if status == "Success":
				logger.info(f"Downloaded {url}: {message}")
			else:
				logger.error(f"Failed to download {url}: {message}")

if __name__ == "__main__":
	from main import ENABLE_LOGGING, LOG_FILE_PATH, setup_logging
	setup_logging()
	app = DownloaderGUI()
	app.mainloop()
