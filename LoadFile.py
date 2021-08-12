import tkinter
from tkinter import filedialog
import tkinter as tk
import GridApp


class LoadFile:
	def __init__(self, main: tk.Tk):

		self.main_frame: tk.Tk = main
		self.main_frame.title("Load Files")

		min_width: int = 500
		min_height: int = 150

		max_width: int = 600
		max_height: int = 250

		plus_x: int = 350
		plus_y: int = 200

		self.main_frame.minsize(width=min_width, height=min_height)
		self.main_frame.maxsize(width=max_width, height=max_height)
		self.main_frame.geometry("%sx%s+%s+%s" % (min_width, min_height, plus_x, plus_y))
		self.main_frame.rowconfigure(0, weight=1)
		self.main_frame.rowconfigure(1, weight=1)
		self.main_frame.rowconfigure(2, weight=1)
		self.main_frame.columnconfigure(0, weight=3)
		self.main_frame.columnconfigure(1, weight=1)

		self.entry_frame: tk.Frame = tk.Frame(master=self.main_frame)
		for row in range(2):
			self.entry_frame.rowconfigure(row, weight=1)
		self.entry_frame.columnconfigure(0, weight=1)
		self.entry_frame.grid(row=0, column=0, sticky='nsew')

		self.button_frame: tk.Frame = tk.Frame(master=self.main_frame)
		for row in range(2):
			self.button_frame.rowconfigure(row, weight=1)
		self.button_frame.columnconfigure(0, weight=1)
		self.button_frame.grid(row=0, column=1, sticky='nsew')

		self.data_entry_variable: tk.StringVar = tk.StringVar()
		self.data_entry: tk.Entry = tk.Entry(master=self.entry_frame, textvariable=self.data_entry_variable)
		self.data_entry.grid(row=0, column=0, pady=20, padx=(20, 0), sticky='ew')

		self.drug_entry_variable: tk.StringVar = tk.StringVar()
		self.drug_entry: tk.Entry = tk.Entry(master=self.entry_frame, textvariable=self.drug_entry_variable)
		self.drug_entry.grid(row=1, column=0, pady=20, padx=(20, 0), sticky='ew')

		self.data_button: tk.Button = tk.Button(master=self.button_frame, text='Open Data File')
		self.data_button.config(command=self.data_button_command)
		self.data_button.grid(row=0, column=0, pady=(2, 12))

		self.drug_button: tk.Button = tk.Button(master=self.button_frame, text='Open Drug File')
		self.drug_button.config(command=self.drug_button_command)
		self.drug_button.grid(row=1, column=0, pady=(0, 32))

		self.continue_button_widget: tk.Button = tk.Button(master=self.entry_frame, text='Continue')
		self.continue_button_widget.config(command=lambda: self.continue_button_method(main_frame=self.main_frame))
		self.continue_button_widget.grid(row=2, column=0, padx=(160, 0))

	def data_button_command(self):
		file_name = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
		self.data_entry_variable.set(file_name)

	def drug_button_command(self):
		file_name: str = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
		self.drug_entry_variable.set(file_name)

	def continue_button_method(self, main_frame):
		self.main_frame.destroy()

		grid_main: tk.Tk = tk.Tk()
		GridApp.Grid(grid_main, data_file=self.data_entry_variable.get(), drug_file=self.drug_entry_variable.get())
		grid_main.mainloop()


if __name__ == '__main__':
	root: tk.Tk = tk.Tk()
	LoadFile(root)
	root.mainloop()
