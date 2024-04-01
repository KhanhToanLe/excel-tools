from tkinter import *

class InputData:
  def __init__(self, frame):
    self.frame = frame

class ExcelChangeData:
  def __init__(self,cell,value):
    self.cell = cell
    self.value = value

change_data = []
excel_change_data_list = []