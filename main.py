import pandas as pd
from pandas import DataFrame
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter.ttk import Treeview

app = Tk()

topbar = ttk.Frame(app)
topbar.grid(row=0, column=0, sticky="ew")

col_to_filter = StringVar(app)
col_to_filter.set("Select column to filter")
col_to_filter_options = []

app.title("ExcelViewer V.1")
app.state("zoomed")

def updateTree(tree: Treeview, df: DataFrame):
  df = df.sort_values(by="occurred_at", ascending=True)
  tree.delete(*tree.get_children())
  tree.heading("line", text="line")
  tree.column("line", width=100, stretch=False)
  for col in df.columns:
    tree.heading(col, text=col)
    tree.column(col, stretch=False)

  line = 1
  for _, row in df.iterrows():
    tree.insert("", END, values=[line] + list(row))
    line += 1

file_cache = None
def open_excel():
  global file_cache
  global col_to_filter_options
  global col_to_filter_dropdown
  global PLACEHOLDER
  global COL_FILTER_DROPDOWN_DEFAULT
  global search_keyword
  global search_bar
  global find_button

  file_path = filedialog.askopenfilename(
    title="Select excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
  )

  if not file_path:
    return
  
  try:
    df = pd.read_excel(file_path)

    df["fullname"] = (
      df["firstname"].fillna("") + " " + df["lastname"].fillna("")
    )

    cols = list(df.columns)
    cols.remove("fullname")

    idx = cols.index("lastname") + 1
    cols.insert(idx, "fullname")

    df = df[cols]

    file_cache = df.copy()

    tree.delete(*tree.get_children())

    tree["columns"] = ["line"] + list(df.columns)
    tree["show"] = "headings"

    col_to_filter_options = [
      col for col in df.columns 
      if df[col].dtype == "object"
    ]
    print(col_to_filter_options)

    if col_to_filter_dropdown:
      col_to_filter_dropdown.destroy()

    col_to_filter_dropdown = ttk.OptionMenu(topbar, col_to_filter, COL_FILTER_DROPDOWN_DEFAULT, *col_to_filter_options)
    col_to_filter_dropdown.grid(row=0, column=1)

    search_bar = ttk.Entry(topbar, textvariable=search_keyword)
    search_bar.grid(row=0, column=3)
    search_keyword.set(PLACEHOLDER)
    search_bar.bind("<FocusIn>", search_focus)
    search_bar.bind("<FocusOut>", search_unfocus)

    updateTree(tree=tree, df=df)

  except Exception as e:
    print(e)
    messagebox.showerror("Error", str(e))

open_ecxel_button = ttk.Button(topbar, text="Open Excel", command=open_excel)
open_ecxel_button.grid(row=0, column=0)

COL_FILTER_DROPDOWN_DEFAULT = "Select column to filter"
col_to_filter = StringVar()
col_to_filter_dropdown = None

PLACEHOLDER = "Search..."
search_keyword = StringVar()
search_bar = None

def search_focus(*args):
  if (search_keyword.get() != PLACEHOLDER):
    return
  if (search_bar):
    search_keyword.set("")

def search_unfocus(*args):
  if (search_keyword.get()):
    return
  if (search_bar):
    search_keyword.set(PLACEHOLDER)

def search(*args):
  if not search_keyword.get():
    print("reset")
    updateTree(tree=tree, df=file_cache)
    return

  if search_keyword.get() != PLACEHOLDER:
    keyword = search_keyword.get()
    filter = col_to_filter.get()
    try:
      if (filter != COL_FILTER_DROPDOWN_DEFAULT):
        df = file_cache
        filtered_df = df[
          df[filter].astype(str).str.contains(keyword)
        ]
        updateTree(tree=tree, df=filtered_df)
    except Exception as e:
      print(e)
      messagebox.showerror("Error", str(e))

search_keyword.trace_add("write", search)

table = ttk.Frame(app)
table.grid(row=1, column=0, sticky="nsew")

tree = ttk.Treeview(table)
tree.grid(row=0, column=0, sticky="nsew")

scroll_y = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
scroll_x = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)

tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

scroll_y.grid(row=0, column=1, sticky="ns")
scroll_x.grid(row=1, column=0, sticky="ew")

app.rowconfigure(1, weight=1)
app.columnconfigure(0, weight=1)

table.rowconfigure(0, weight=1)
table.columnconfigure(0, weight=1)

app.mainloop()