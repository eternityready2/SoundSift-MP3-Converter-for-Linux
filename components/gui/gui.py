from customtkinter import CTk


class Root(CTk):
	"""A class to model customtkinter needed windows.
	"""
	def __init__(self):
		super().__init__()
		
		# # Add widgets to root.
		# Set a window title.
		self.title('PW storage')
		# Set the min size.
		self.minsize(width=1000, height=600)
		# Set the geometry.
		self.geometry('1000x600')
		# Setup a style.
		self.style = ttk.Style()
		self.style.theme_use("default")
	# Configure treeview color.
	style.configure("Treeview",
			background='#D3D3D3',
			foreground="black",
			rowheight=25,
			fieldbackground='#D3D3D3'
	)
	style.map("Treeview",
			background=[('selected', '#347083')])
		
		
		


# Set an icon.
root.iconbitmap(os.path.join(cwd, 'images\\pw.ico'))
# Set the min size.
root.minsize(width=1000, height=600)
# Set the geometry.
root.geometry('1000x600')
# Setup a style.
style = ttk.Style()
# Seleect a theme.
style.theme_use("default")

# Configure treeview color.
style.configure("Treeview",
                background= '#D3D3D3',
                foreground= "black",
                rowheight=25,
                fieldbackground='#D3D3D3'
                )
style.map("Treeview",
          background=[('selected', '#347083')])

# Create a treeview frame.
tree_frame = Frame(root)
tree_frame.pack(fill="x", padx=20, pady=20)

# Create a treeview Scrollbar.
tree_scroll = Scrollbar(tree_frame)
tree_scroll.pack(side=RIGHT, fill=Y)

# Create the Treeview and pack it on the screen.
my_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode='extended')
my_tree.pack(fill="x")

# Configure the Scrollbar.
tree_scroll.config(command=my_tree.yview)