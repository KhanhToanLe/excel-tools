from tkinter import *

class InputData:
  def __init__(self, frame):
    self.frame = frame

class ChangeData:
  def __init__(self, cell, value):
    self.cell = cell
    self.value = value


change_data = []
command_data = []

# static variable
INFINITE_SHEET_CHECKER = "∞"
CHANGE_DATA_TAB_NAME = ".!notebook.change_data"
CONTROL_HOME_TAB_NAME = ".!notebook.control_home"