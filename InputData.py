from tkinter import *

class InputData:
  def __init__(self, frame):
    self.frame = frame

class ChangeData:
  def __init__(self, cell, value):
    self.cell = cell
    self.value = value
INFINITE_SHEET_CHECKER = "∞"

change_data = []
command_data = []