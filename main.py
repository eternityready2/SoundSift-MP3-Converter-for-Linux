import customtkinter as ctk
from tkinter import ttk
from soundsift.components.app import download_songs as dwl



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
        self.title("Sound Sift - Easy MP3s")
        self.geometry("600x400")
        self.storage = DataStorage()

        self.style = ttk.Style(self)
        self.style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        self.style.configure("Treeview", rowheight=25, borderwidth=1, relief="solid")

        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("STATUS", "LINK"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("STATUS", text="STATUS")
        self.tree.heading("LINK", text="LINK")
        self.tree.column("STATUS", width=100, anchor="center")
        self.tree.column("LINK", width=400, anchor="w")

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.link_label = ctk.CTkLabel(self, text="Enter or Update Link:")
        self.link_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 0))

        self.link_entry = ctk.CTkEntry(self, placeholder_text="Enter or Update Link")
        self.link_entry.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        self.add_button = ctk.CTkButton(self, text="Add Row", command=self.add_row)
        self.add_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

        self.process_button = ctk.CTkButton(self, text="Process Links", command=self.confirm_row)
        self.process_button.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        self.update_button = ctk.CTkButton(self, text="Update Link", command=self.update_row)
        self.update_button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

        self.delete_button = ctk.CTkButton(self, text="Delete Row", command=self.delete_row)
        self.delete_button.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tree.tag_configure("evenrow", background="#f2f2f2")
        self.tree.tag_configure("oddrow", background="#ffffff")

    def add_row(self):
        link = self.link_entry.get()
        if link:
            tag = "evenrow" if len(self.tree.get_children()) % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=("NEW", link), tags=(tag,))
            self.storage.add_data("NEW", link)
            self.link_entry.delete(0, "end")

    def confirm_row(self):
        # Get all items (all rows) in the Treeview
        all_items = self.tree.get_children()

        for item in all_items:
            current_values = self.tree.item(item, "values")  # (status, link)
            status, link = current_values

            # Update the status to "CONFIRMED"
            self.tree.item(item, values=("PROCESSING", link))
            index = self.tree.index(item)
            self.storage.update_data(index, status="PROCESSING")

            # Print the second column (the link)
            status = dwl.appl.download_music_direct(link)
            #print(link)

            # Update the status to "DOWN"
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

def main ():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()
    


if __name__ == "__main__":
    main()

