
from tkinter import messagebox
import tkinter as tk
import time

class GUI(tk.Frame):
    def __init__(self, master=None):

        # Create API Access

        tk.Frame.__init__(self, master)

        self.main_left_frame = tk.Frame(master=master, width=200, height=400, background="grey")
        self.main_right_frame = tk.Frame(master=master, width=200, height=400, background="grey")
        self.information_frame = tk.Frame(master=self.main_left_frame, width=150, height=200,
                                          background="white", highlightbackground="black", highlightthickness=1)
        self.output_frame = tk.Frame(master=self.main_right_frame, width=150, height=200, background="white",
                                     highlightbackground="green", highlightthickness=1)

        # Create Input Frame

        self.information_title = tk.Label(master=self.information_frame,
                                          foreground="black", background="white", width=50, height=5,
                                          text="Information Section")
        self.output_title = tk.Label(master=self.output_frame,
                                     foreground="black", background="white", width=50, height=5,
                                     text="Output Section")
        self.output_log = tk.StringVar()
        self.output_section = tk.Label(master=self.output_frame,
                                     foreground="black", background="white", width=100, height=5,
                                     textvariable=self.output_log)

        self.output_log.set("Hello There")
        # self.button_search = tk.Button(master=self.frame_input, text="Search",
        #                                width=10, height=1, bg="white", fg="black")
        #
        # self.entry_search = tk.Entry(master=self.frame_input, bg="white", fg="black", width=20)
        #
        # self.label.grid(column=0, row=0)
        self.information_title.grid(column=0, row=0)
        self.output_title.grid(column=0, row=0)
        self.output_section.grid(column=0, row=1)

        # master panel
        self.main_left_frame.grid(column=0, row=0)
        self.main_right_frame.grid(column=1, row=0, sticky=tk.N)

        # left panel
        self.information_frame.grid(column=0, row=0)

        # right panel
        self.output_frame.grid(column=0, row=0, sticky=tk.N)


    def update_output(self):

        self.output_log.set("Updated")



