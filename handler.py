from tkinter import *
from workbook import Workbook as Wb
import os
from math import ceil
import tkinter.font as tkFont
import tkinter.ttk as ttk
import threading
import tkinter.messagebox as msgbox

class Handler:
  def __init__(self):
    self.dataFolder = "./listOfNames"
    self.root = Tk()
    self.root.title("과학실 자리 정하기")
    self.root.geometry("450x650")
    self.root.resizable(False, False)
    self.active_frames = list()
    self.screenIdx = 1
    self.wb = None
    self.txtFont = tkFont.Font(family="굴림", size=14)
    self.lblFont = tkFont.Font(family="굴림", size=18)
    self.classRoom = ''
    self.leftArrow = PhotoImage(file="./images/left_arrow_resized.png")
    self.rightArrow = PhotoImage(file="./images/right_arrow_resized.png")
    self.drawScreen()
    self.root.mainloop()

  def drawScreen(self):
    self.remove_all_frames()
    screen = Frame(self.root)
    screen.pack(fill="both")
    self.insertToActiveFrames(screen)

    if self.screenIdx == 1:
      listbox = Listbox(screen, selectmode="single", height=0)
      listbox.configure(font=self.txtFont)
      listbox.pack()
      filenames = os.listdir(self.dataFolder)
      for filename in filenames:
        name, ext = os.path.splitext(filename)
        if ext == '.xlsx':
          listbox.insert(END, name)
      listbox.bind('<<ListboxSelect>>', self.event_for_listbox)
    
    elif self.screenIdx == 2:
      classlbl = Label(screen, text=self.classRoom)
      classlbl.configure(font=self.lblFont)
      classlbl.grid(row=0, column=0, columnspan=4)
      btn_left = Button(screen, image=self.leftArrow, relief="flat")
      btn_right = Button(screen, image=self.rightArrow, relief="flat")
      btn_seats = Button(screen, text="자리 바꾸기")
      btn_delete = Button(screen, text="삭제하기")
      btn_seats.configure(font=self.txtFont)
      btn_delete.configure(font=self.txtFont)
      btn_left.grid(row=1, column=0, rowspan=3, padx=(15, 0))
      btn_right.grid(row=1, column=3, rowspan=3)
      btn_seats.grid(row=4, column=2, pady=5)
      btn_delete.grid(row=4, column=1, pady=5)
      for k, group in enumerate(self.wb.currentArrangement, start=1):
        txt = Text(screen, width=11, height=8)
        txt.configure(font=self.txtFont)
        txt.grid(row=ceil(k/2), column=1 if k%2 == 1 else 2, padx=20, pady=15)
        for idx, member in enumerate(group, start=1):
          txt.insert(f'{2 * idx - 1}.0', member + '\n\n')
      btn_seats.bind('<ButtonRelease-1>', self.generate_new_seats)
      btn_delete.bind('<ButtonRelease-1>', self.delete_latest_seats)
      btn_left.bind('<ButtonRelease-1>', self.to_previous_index)
      btn_right.bind('<ButtonRelease-1>', self.to_next_index)
      if self.wb.index == 1:
        btn_right.config(state='disabled')
      else:
        btn_seats.grid_forget()
        btn_delete.grid_forget()
      if self.wb.index == self.wb.lenOfRemarksColumn:
        btn_left.config(state='disabled')

    elif self.screenIdx == 3:
      classlbl = Label(screen, text=self.classRoom)
      classlbl.configure(font=self.lblFont)
      classlbl.grid(row=0, column=0, columnspan=4)
      btn_left = Button(screen, image=self.leftArrow, relief="flat")
      btn_right = Button(screen, image=self.rightArrow, relief="flat")
      btn_discard = Button(screen, text="취소하기")
      btn_save = Button(screen, text="저장하기")
      btn_discard.configure(font=self.txtFont)
      btn_save.configure(font=self.txtFont)
      btn_left.grid(row=1, column=0, rowspan=3, padx=(15, 0))
      btn_right.grid(row=1, column=3, rowspan=3)
      btn_discard.grid(row=4, column=1, pady=5)
      btn_save.grid(row=4, column=2, pady=5)
      for k, group in enumerate(self.wb.processedResult, start=1):
        txt = Text(screen, width=11, height=8)
        txt.configure(font=self.txtFont)
        txt.grid(row=ceil(k/2), column=1 if k%2 == 1 else 2, padx=20, pady=15)
        for idx, member in enumerate(group, start=1):
          txt.insert(f'{2 * idx - 1}.0', member + '\n\n')
      btn_discard.bind('<ButtonRelease-1>', self.to_previous_screen)
      btn_save.bind('<ButtonRelease-1>', self.save_generated_seats)

  def readFile(self):
    file_name = self.classRoom + '.xlsx'
    self.wb = Wb(file_name)
    self.screenIdx = 2
    self.drawScreen()

  def event_for_listbox(self, event):
    listbox = event.widget
    selection = listbox.curselection()
    if selection:
      index = selection[0]
      self.classRoom = listbox.get(index)
      self.readFile()

  def generate_new_seats(self, event):
    parent_widget = event.widget.master
    event.widget.destroy()
    for child in parent_widget.winfo_children():
      if isinstance(child, Button) and child['text'] == '삭제하기':
        child.destroy()
    parent_widget.update_idletasks()
    progressbar = ttk.Progressbar(parent_widget, maximum=100, length=150, mode="indeterminate")
    progressbar.grid(row=4, column=1, columnspan=2, pady=5)
    # progress_label = Label(parent_widget, text="처리중...", anchor="center")
    # progress_label.grid(row=4, column=1, columnspan=2, pady=5)
    progressbar.start(3)
    parent_widget.update_idletasks()
    def long_task():
      self.wb.rearrangementOfSeats()
      print(self.wb.processedResult)
      self.screenIdx = 3
      self.drawScreen()
    threading.Thread(target=long_task).start()

  def delete_latest_seats(self, event):
    response = msgbox.askyesno("예 / 아니오", f"{self.classRoom}의 이 자리 배치를 삭제하시겠습니까?")
    if response == 1:
      try:
        self.wb.deleteLatestValues()
        msgbox.showinfo("삭제완료", "자리를 성공적으로 삭제하였습니다.")
        self.readFile()
      except:
        msgbox.showerror("삭제실패", "삭제중에 오류가 발생했습니다.")
    elif response == 0:
      msgbox.showwarning("미삭제", "삭제를 하지 않았습니다.")

  def to_previous_index(self, event):
    self.wb.toPreviousIndex()
    self.screenIdx = 2
    self.drawScreen()

  def to_next_index(self, event):
    self.wb.toNextIndex()
    self.screenIdx = 2
    self.drawScreen()

  def to_previous_screen(self, event):
    self.readFile()

  def save_generated_seats(self, event):
    response = msgbox.askyesno("예 / 아니오", f"화면상의 자리 배치를 {self.classRoom} 엑셀 파일에 추가하시겠습니까?")
    if response == 1:
      try:
        self.wb.writeResultToFile()
        msgbox.showinfo("저장완료", "자리를 성공적으로 저장하였습니다.")
        self.readFile()
      except:
        msgbox.showerror("저장실패", "저장중에 오류가 발생했습니다.")
    elif response == 0:
      msgbox.showwarning("미저장", "저장을 하지 않았습니다.")

  def insertToActiveFrames(self, frame):
    self.active_frames.append(frame)
    
  def remove_all_frames(self):
    if self.active_frames:
      for _ in self.active_frames:
        frame = self.active_frames.pop()
        frame.destroy()

temp = Handler()