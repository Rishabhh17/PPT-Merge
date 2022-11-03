from cProfile import label
from email.mime import image
import os
import re
from tkinter import *
from tkinter import filedialog
from TkinterDnD2 import *
from threading import Thread
from PIL import Image, ImageTk

from main import merge_presentations
from presentation_file import FileForm, FileFormList


def merge(output_entry, message_label, input_list):
  output_path = output_entry.get()
  input_errors = input_list.input_errors()
  output_error = output_entry.file_errors()
  if len(input_errors) == 0 and output_error is None:
    presentations = input_list.get_filepaths()
    def wrapper_worker(pres, out):
      merge_presentations(pres, out)
      message_label.config(text="Merged presentations successfully!", fg="green")
    
    thread = Thread(target=wrapper_worker, args=(presentations, output_path))
    thread.start()
    message_label.config(text="Merging presentations...", fg="blue")
  elif len(input_errors) > 0:
    message_label.config(text="Error(s) occurred in input fields:" + "\n* ".join(input_errors), fg="red")
  else:
    message_label.config(text="Error(s) occurred in output fields:\n* " + output_error, fg="red")

def setup_DnD(frame, input_list):
  # Make the window a drop target for files.
  frame.drop_target_register(DND_FILES)
  def drop(event):
    # Detect if one or several files are beeing droped
    # print(event.data)
    # Match any number of characters between brackets.
    # Any character: [\s\S]; One or more characters: *; Match as few characters as possible: ?; Group: (); In brackets: {}
    regExp = r"{([\s\S]*?)}"
    with_whitespace = re.findall(regExp, event.data)
    no_whitespace = [path for path in re.sub(regExp, "", event.data).split(" ") if path != ""]
    filepaths = with_whitespace + no_whitespace
    # Get the last empty input, from bottom to top. Those empty imputs will be filled with the dragged files.
    last_empty_var = len(input_list.inputs)
    for _ in range(last_empty_var):
      if last_empty_var >= 1:
        if input_list.get_value(last_empty_var-1) == "":
          last_empty_var -= 1
          continue
      break

    for i, fpath in enumerate(filepaths):
      if i >= last_empty_var and i < len(input_list.inputs):
        input_list.set_value(i, fpath)
      else:
        input_list.add_input()
        input_list.set_value(-1, fpath)
  
  frame.dnd_bind('<<Drop>>', drop)

def render_gui():
  # root = Tk()
  root = TkinterDnD.Tk()
  root.title("PPT Merger App")
  root.geometry("700x700")
  root.iconbitmap(os.path.abspath("icon.ico"))

  ## Top heading for the hole view.
  heading = Label(text="PPT Merger App", bg="#004696", fg="white", height="3", font=("TkHeadingFont", 30))
  heading.pack(side="top", fill="both")


  ## Image for heading.
  img= (Image.open("BASF.png"))
  resize_image=img.resize((65,35))
  my_img = ImageTk.PhotoImage(resize_image)
  l2 = Label(heading,  image=my_img, bg="#004696")
  l2.pack(side="right") 
  

  # ## Configure the output path input frame, label and entry ##
  output_frame = Frame(root, pady="30")
  output_frame.pack(side=TOP)
  output_entry = FileForm(output_frame, is_input=False, text="Save as", deletable=False)

  ## Label for input files ##
  input_label = Label(root, text="Input PowerPoint Presentations", bg="#004696", fg="white", font="Helvetica 12 bold")
  input_label.pack(side=TOP, fill=X)

  ## Create frame to store all presentation inputs
  canvas_frame = Frame(root, pady="10")
  canvas_frame.pack(side=TOP)

  # Make the input frame scrollable
  canvas = Canvas(canvas_frame, highlightthickness=0)
  input_frame = Frame(canvas)
  scrollbar = Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
  canvas.configure(yscrollcommand=scrollbar.set)

  def update_canvas(event):
    canvas.configure(scrollregion=canvas.bbox("all"), width=400, height=300)

  scrollbar.pack(side=RIGHT, fill=Y)
  canvas.pack(side=LEFT)
  canvas.create_window((200, 50), window=input_frame)
  input_frame.bind("<Configure>", update_canvas)

  # Put the two input fields for merging the presentations (at least two)
  input_list = FileFormList(input_frame)
  input_list.add_input(deletable=False)
  input_list.add_input(deletable=False)
  

  add_btn_frame = Frame(root)
  add_btn_frame.pack(side=TOP)

  setup_DnD(root, input_list)

  # Button to add a new presentation to merge
  click_add_btn= PhotoImage(file='button_add-presentation.png')
  add_img_label= Label(image=click_add_btn)
  
  add_btn = Button(add_btn_frame, image=click_add_btn, command=lambda: input_list.add_input(input_frame),bd=0)
  add_btn.pack(side=LEFT, pady="10", padx="5")

  def browse_many_dialog():
    filenames = filedialog.askopenfilenames(title="Select presentations", filetype=[("PowerPoint files", "*.pptx")])
    if len(input_list.inputs) == 2 and input_list.get_value(0) == "" and input_list.get_value(1) == "" and len(filenames) >= 2:
      for i in range(len(filenames)):
        if i < 2:
          input_list.set_value(i, filenames[i])
        else:
          input_list.add_input(input_frame)
          input_list.set_value(-1, filenames[i])
    else:
      for fname in filenames:
        input_list.add_input(input_frame)
        input_list.set_value(-1, fname)
  click_select_files= PhotoImage(file='button_select-files.png')
  select_img_label= Label(image=click_select_files)
  browse_btn = Button(add_btn_frame, image=click_select_files, command=browse_many_dialog,bd=0)
  browse_btn.pack(side=RIGHT, pady="10", padx="5")

  # Label where all errors are displayed
  error_frame = Frame(root)
  error_frame.pack(side=TOP)
  error_label = Label(error_frame, text="", fg="red", font="Consolas 8 bold")
  error_label.pack(side=TOP)

  # Button to merge the presentations
  click_btn= PhotoImage(file='button_merge.png')
  img_label= Label(image=click_btn)
  button = Button(root, image=click_btn, command=lambda: merge(output_entry, error_label, input_list),bd=0)
  button.pack(side=BOTTOM, pady="30")

  root.mainloop()