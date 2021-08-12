import sys
import tkinter
import tkinter as tk

import tkmacosx
from tkmacosx import Button as Button
import AnalysisApp

# MacOS does not allow coloring button backgrounds from default tkinter package
# Must use tkmacosx instead
if sys.platform == "darwin":
    from tkmacosx import Button as Button
    buttonType = tkmacosx.Button
else:
    from tkinter import Button as Button
    buttonType = tkinter.Button


class Grid:
    def __init__(self, main, data_file, drug_file):
        try:
            self.data_file: str = data_file
            self.drug_file: str = drug_file
        except TypeError:
            pass

        self.root: tk.Tk = main
        self.root.title("Data Analysis")
        # set the window location and min/max width and height
        min_width: int = 775
        min_height: int = 300

        max_width: int = 800
        max_height: int = 325

        plus_x: int = 200
        plus_y: int = 150

        self.root.minsize(width=min_width, height=min_height)
        self.root.maxsize(width=max_width, height=max_height)
        self.root.geometry("%sx%s+%s+%s" % (min_width, min_height, plus_x, plus_y))

        self.main_frame: tk.Tk = main
        tk.Grid.rowconfigure(self.main_frame, 0, weight=1)
        tk.Grid.columnconfigure(self.main_frame, 0, weight=1)

        self.checkbutton_frame: tk.Frame = tk.Frame(master=self.main_frame)
        for row in range(7):
            tk.Grid.rowconfigure(self.checkbutton_frame, row, weight=1)
        self.checkbutton_frame.columnconfigure(0, weight=1)
        self.checkbutton_frame.grid(row=0, column=1, sticky='ns')

        self.button_frame: tk.Frame = tk.Frame(master=self.main_frame)
        for row in range(7):
            tk.Grid.rowconfigure(self.button_frame, row, weight=1)
        tk.Grid.columnconfigure(self.button_frame, 0, weight=1)
        self.button_frame.grid(row=0, column=0, sticky='nsew')

        self.checkbutton_names: list[str] = ['Sensitive',
                                             'Resistant',
                                             'Sensitive Neg. Control',
                                             'Sensitive Pos. Control',
                                             'Resistant Neg. Control',
                                             'Resistant Pos. Control',
                                             'Clear']

        self.checkbutton_vars: list[tk.IntVar] = [tk.IntVar(),
                                                  tk.IntVar(),
                                                  tk.IntVar(),
                                                  tk.IntVar(),
                                                  tk.IntVar(),
                                                  tk.IntVar(),
                                                  tk.IntVar()]

        self.colors: list[str] = ['#6666ff',  # Sensitive                  -- Plum
                                  '#ff0000',  # Resistant                  -- Red
                                  '#00ff00',  # Sensitive Negative Control -- Green
                                  '#ff00ff',  # Sensitive Positive Control -- Magenta
                                  '#00ffff',  # Resistant Negative Control -- Cyan
                                  '#008a45',  # Resistant Positive Control -- Dark Cyan
                                  'SystemButtonFace']  # No Color          -- Default font

        self.checkbutton_instances: list = []
        self.button_list: dict = {}
        self.first_button = None
        self.second_button = None
        self.error_label: tk.Label = tk.Label(master=self.button_frame)
        self.note_label: tk.Label = tk.Label(self.button_frame)
        self.colored_buttons_row: int = 0
        self.colored_buttons_column: int = 0

        self.all_colored_buttons: dict = {}

        self.createbuttongrid()
        self.createcheckbuttons()

    def createbuttongrid(self):
        label: int = 1
        for row in range(8):
            for column in range(12):
                button: buttonType = Button(self.button_frame, text=f"Well {label}", overrelief="sunken")
                button.bind("<Button-1>", self.clickbutton)
                button.bind("<ButtonRelease-1>", self.releasebutton)
                button.grid(row=row, column=column, sticky='nsew')
                self.button_list[button] = (row, column)
                label += 1

        for row in range(8):
            tk.Grid.rowconfigure(self.button_frame, row, weight=1)
        for column in range(12):
            tk.Grid.columnconfigure(self.button_frame, column, weight=1)

        self.note_label.grid(row=9, columnspan=7, sticky='w')

        self.error_label.grid(row=9, columnspan=3, column=8, sticky='w')

    def createcheckbuttons(self):
        # creates checkbuttons on right side

        for var in range(len(self.checkbutton_names)):
            # place checkbuttons on the right side of the screen
            self.checkbutton_instances.append(tk.Checkbutton(master=self.checkbutton_frame,
                                                             text=self.checkbutton_names[var],
                                                             variable=self.checkbutton_vars[var],
                                                             anchor='w'))
            self.checkbutton_instances[var].config(bg=self.colors[var], height=1, pady=4)
            self.checkbutton_instances[var].grid(sticky='nsew', columnspan=2)

        continue_button = Button(self.checkbutton_frame,
                                 text="Continue",
                                 bg="#66ff33",
                                 activebackground="#00cc44",
                                 command=self.continuebutton)

        continue_button.grid(row=6, column=1, sticky='nsew')

    def clickbutton(self, event):
        # get the first button clicked
        self.first_button: buttonType = event.widget

    def releasebutton(self, event):
        # collect row/column of second button
        self.second_button: buttonType = event.widget.winfo_containing(event.x_root, event.y_root)
        self.colorbuttons()

    def colorbuttons(self):
        color: str = ""  # used to get rid of 'this variable may not be defined' in the for loops below

        if self.checkbutton_vars[0].get():
            if self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[0]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[1].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[1]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[2].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[2]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[3].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[3]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[4].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[4]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[5].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[6].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[5]
                self.error_label.config(bg='SystemButtonFace', text='')

        elif self.checkbutton_vars[6].get():
            if self.checkbutton_vars[0].get() == 1 or \
                    self.checkbutton_vars[1].get() == 1 or \
                    self.checkbutton_vars[2].get() == 1 or \
                    self.checkbutton_vars[3].get() == 1 or \
                    self.checkbutton_vars[4].get() == 1 or \
                    self.checkbutton_vars[5].get() == 1:
                self.error_label.config(bg='red', fg='white', text="Sorry, that doesn't work")
            else:
                color: str = self.colors[6]
                self.error_label.config(bg='SystemButtonFace', text='')

        else:
            color: str = ""
            self.error_label.config(bg='red', fg='white', text='Select a color')

        try:
            row_button_one: int = self.button_list[self.first_button][0]  # row of first button clicked
            column_button_one: int = self.button_list[self.first_button][1]  # column of first button clicked

            row_button_two: int = self.button_list[self.second_button][0]  # row of second button clicked
            column_button_two: int = self.button_list[self.second_button][1]  # column of second button clicked

            # these lines allow dragging from the bottom right to the top left
            if row_button_one > row_button_two:
                row_button_one: int = self.button_list[self.second_button][0]
                row_button_two: int = self.button_list[self.first_button][0]
            if column_button_one > column_button_two:
                column_button_one: int = self.button_list[self.second_button][1]
                column_button_two: int = self.button_list[self.first_button][1]

            for button in self.button_list:
                if self.button_list[button][0] in range(row_button_one, row_button_two + 1):
                    # if the row value is in the range of click & drag buttons, add it to buttons_to_color list

                    if self.button_list[button][1] in range(column_button_one, column_button_two + 1):
                        # if the column value is in the range of click & drag button, color the button
                        button.config(bg=color)
                        self.all_colored_buttons[button] = color

                        # add one to the row/column after so that (0, 0) can be accounted for
                        self.colored_buttons_row += 1
                        self.colored_buttons_column += 1

        except (KeyError, tk.TclError):
            self.error_label.config(text='Please try that again.', bg='red', fg='white')

    def continuebutton(self):
        continue_window: tk.Tk = tk.Tk()
        continue_window.title("Continue")
        min_width: int = 300
        min_height: int = 85
        plus_x: int = 400
        plus_y: int = 200

        continue_window.geometry("%sx%s+%s+%s" % (min_width, min_height, plus_x, plus_y))

        continue_window.rowconfigure(1, weight=1)
        for column in range(2):
            continue_window.columnconfigure(column, weight=1)

        continue_label: tk.Label = tk.Label(master=continue_window, text="Are you sure your layout is correct?\n"
                                                               "Selecting 'Yes' will overwrite the selected data file.\n"
                                                               "Please close this Excel file before continuing.")
        yes_button: buttonType = Button(master=continue_window, text='Yes')
        no_button: buttonType = Button(master=continue_window, text='No')

        yes_button.config(command=lambda: self.continue_button_yes(continue_window=continue_window))
        no_button.config(command=lambda: continue_window.destroy())

        continue_label.grid(row=0, column=0, columnspan=2, sticky='n')
        yes_button.grid(row=1, column=0, sticky='e', ipadx=12, pady=(0, 5))
        no_button.grid(row=1, column=1, sticky='w', ipadx=12, pady=(0, 5))

        continue_window.mainloop()

    def continue_button_yes(self, continue_window):
        continue_window.destroy()
        self.root.destroy()

        analysis_root: tk.Tk = tk.Tk()
        AnalysisApp.Analysis(main=analysis_root,
                             data_file=self.data_file,
                             drug_file=self.drug_file,
                             colored_buttons=self.all_colored_buttons,
                             button_list=self.button_list,
                             font_colors=self.colors)
        analysis_root.mainloop()

    def continue_button_no(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    Grid(main=root, data_file=None, drug_file=None)
    root.mainloop()
