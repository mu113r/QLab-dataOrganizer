import tkinter as tk
import tkinter.filedialog as fl
import read_file as rd
import p5 as p5
import neuro_lu2 as neuro_lu2

root = tk.Tk()
root.title("Excel Manipulator")
files_list = []
def get_files_list():
    global files_list
    # returns a tuple with paths to files
    files = fl.askopenfilenames(parent=root,title='Choose files')
    e.delete(0, len(files_list)-1)
    files_list += list(files)
    for fil in files_list:
        e.insert(tk.END, fil)

        
def remove_file_list():
    global files_list
    removed = e.get(tk.ACTIVE)
    files_list.remove(removed)
    e.delete(tk.ACTIVE)


def execute():
    global files_list
    rd.build_output(files_list)
    root.destroy()
    
def executeP5():
    global files_list
    p5.build_output(files_list)
    root.destroy()

def execute2():
    global files_list
    neuro_lu2.build_output(files_list)
    root.destroy()

scroll = tk.Scrollbar(root)
scroll.grid(row=0, column=2, sticky=tk.N+tk.S)
e = tk.Listbox(root, bg="white", height=15, width=50, bd=3, yscrollcommand=scroll.set)
e.grid(row=0, column=1)
scroll.config(command=e.yview)
select_files_button = tk.Button(root, text="select files", width=15, command=get_files_list)
select_files_button.grid(row=0, column=0, sticky=tk.N, pady=2)
remove_files_button = tk.Button(root, text="remove file", width=15, command=remove_file_list)
remove_files_button.grid(row=0, column=0, sticky=tk.N, pady=34)
execute_button = tk.Button(root, text="Neurolucida", width=15, command=execute)
execute_button2 = tk.Button(root, text="Stereo Investigator", width=15, command=executeP5)
execute_button3 = tk.Button(root, text="Neurolucida II", width=15, command=execute2)
execute_button.grid(row=0, column=0, sticky=tk.S, pady=2)
execute_button2.grid(row=0, column=0, sticky=tk.S, pady=66)
execute_button3.grid(row=0, column=0, sticky=tk.S, pady=34)
root.mainloop()

