import tkinter as tk #Import tkinter module to create a Graphical User Interface
from tkinter.filedialog import askopenfile, asksaveasfile
from tkinter.messagebox import showinfo
from csv import reader

def create_sp(rows, columns, width, lst):
    global root
    try:
        ask_dims.destroy()
    except:
        pass
    try:
        ask_width.destroy()
    except:
        pass
    root = tk.Tk()
    root.title('Python Spreadsheet Software')
    columns = str(int(columns)-1)
    for c in range(int(columns)):
        for r in range(int(rows)):
            cmd = 'r'+str(r)+'c'+str(c)
            if r == 0:
                (tk.Label(root, text=str(c+1))).place(x = (int(r)*12*int(width)), y = str((int(c)*26)+40))
                exec('global '+cmd+'; '+cmd+ '= tk.Entry(root, bg = "light blue", width='+str(width)+')')
            else:
                if c == 0:
                    (tk.Label(root, text=str(r))).place(x = str((int(r)*12*int(width))), y = 0)
                    exec('global '+cmd+'; '+cmd+ '= tk.Entry(root, bg = "light blue", width='+str(width)+')')
                else:
                    exec('global '+cmd+'; '+cmd+ '= tk.Entry(root, width='+str(width)+')')
            if lst == None:
                pass
            else:
                try:
                    exec('global '+cmd+'; '+cmd+ '.insert(tk.END, lst[r][c])')
                except:
                    pass
            exec(cmd+'.place(x='+str((int(r)*12*int(width))+40)+', y ='+str((int(c)*26)+40)+')')
    (tk.Label(root, text=str(r+1))).place(x = str((int(r+1)*12*int(width))), y = 0)

    def savefile():
        global csv_text
        file_loc = asksaveasfile(filetypes = (("Comma Separated Values","*.csv"),))
        csv_text = []
        if hasattr(file_loc, 'name'):
            file_loc = open(file_loc.name, 'w')
            for c in range(int(columns)):
                csv_text.append('\n')
                for r in range(int(rows)):
                    cmd = 'r'+str(r)+'c'+str(c)
                    if not r == int(columns)-1:
                        exec('global  csv_text; csv_text.append((' + cmd + '.get()).replace(",", "&"))')
            csv_text.pop(0)
            csv_text = str(csv_text)
            csv_text = csv_text.replace('[', '')
            csv_text = csv_text.replace("'", '')
            csv_text = csv_text.replace(']', '')
            csv_text = csv_text.replace('\\n, ', '\n')
            file_loc.write(csv_text)
            file_loc.close()
        
    menubar = tk.Menu(root)
    file_drop= tk.Menu(menubar, tearoff=0)
    file_drop.add_command(label="New CSV", command=get_dims)
    file_drop.add_command(label="Load a CSV", command = load_csv)
    file_drop.add_command(label="Save as a CSV", command=savefile)
    menubar.add_cascade(label="File", menu=file_drop)
    root.config(menu=menubar)
    root.mainloop()

def load_csv():
    file_to_open = askopenfile(filetypes = (("Comma Separated Values","*.csv"),))
    if hasattr(file_to_open, 'name'):
        file_to_open = open(file_to_open.name)
        csv_reader = reader(file_to_open)
        list_ = list(csv_reader)
        root.destroy()
        global ask_width
        ask_width = tk.Tk()
        ask_width.title('Set Cell Width')
        set_cell_width_label = tk.Label(ask_width, text='Cell Width')
        set_cell_width_spinbox = tk.Spinbox(ask_width, from_= 4, to = 10)
        next_but = tk.Button(ask_width, text='Plot Spreadsheet', command = lambda: create_sp(columns = len(list_[0]), rows=len(list_), width=int(set_cell_width_spinbox.get()), lst = list_))
        set_cell_width_label.pack()
        set_cell_width_spinbox.pack()
    next_but.pack()
def get_dims():
    try:
        root.destroy()
    except:
        pass
    global ask_dims
    ask_dims = tk.Tk()
    ask_dims.title('Set dimensions')
    rows_label = tk.Label(ask_dims, text = 'Rows')
    rows_spinbox = tk.Spinbox(ask_dims, from_ = 1, to = 100)
    columns_label = tk.Label(ask_dims, text = 'Columns')
    columns_spinbox = tk.Spinbox(ask_dims, from_ = 1, to = 100)
    cell_width_label = tk.Label(ask_dims, text='Cell Width')
    cell_width_spinbox = tk.Spinbox(ask_dims, from_= 4, to = 10)
    rows_label.pack()
    rows_spinbox.pack()
    columns_label.pack()
    columns_spinbox.pack()
    cell_width_label.pack()
    cell_width_spinbox.pack()
    data = None
    create_but = tk.Button(ask_dims, text='Create Spreadsheet',command=lambda: create_sp(rows = columns_spinbox.get(), columns=int(rows_spinbox.get())+1, width=cell_width_spinbox.get(), lst = data))
    create_but.pack()
    ask_dims.mainloop()

get_dims()