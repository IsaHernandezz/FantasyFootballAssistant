import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import os

# file path
base_path = r"C:\Users\ihern\OneDrive\Desktop"


excel_file = os.path.join(base_path, "FantasyFootball- Rickys Version.xlsx")
icon_path = os.path.join(base_path, "football_icon.png")
xls = pd.ExcelFile(excel_file)
sheets = xls.sheet_names  # Get all sheet names
current_sheet = sheets[0]  # Start with the first sheet
data = xls.parse(current_sheet)  # Load the first sheet as default

# Window
root = tk.Tk()
root.title("Fantasy Football Draft Assistant")
root.geometry("1000x600")
root.configure(bg="#001f3d")  # Navy blue background for the main window

# App Icon
try:
    icon_img = Image.open(icon_path).resize((30, 30))
    icon = ImageTk.PhotoImage(icon_img)
except Exception:
    icon = None

#Title Frame
title_frame = tk.Frame(root, bg="#001f3d")
title_frame.pack(pady=20)

if icon:
    icon_label = tk.Label(title_frame, image=icon, bg="#001f3d")
    icon_label.pack(side="left", padx=(10, 5))

title_label = tk.Label(title_frame, text="Fantasy Football Draft Assistant", font=("Arial", 18, "bold"), bg="#001f3d", fg="white")
title_label.pack(side="left", padx=10)

# Button Frame for Sheets
button_frame = tk.Frame(root, bg="#001f3d")  # Dark navy blue background for sheet buttons
button_frame.pack(pady=10)

# Sheet change
def on_sheet_button_click(sheet_name):
    global current_sheet, data  # Make sure to update current_sheet and data when a new sheet is selected
    try:
        current_sheet = sheet_name
        data = xls.parse(sheet_name)
        update_table(data)
        status_label.config(text=f"Sheet loaded: {sheet_name}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheet: {e}")

# Create buttons
for sheet_name in sheets:
    button = tk.Button(button_frame, text=sheet_name, command=lambda sheet_name=sheet_name: on_sheet_button_click(sheet_name), bg="#004080", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
    button.pack(side="left", padx=5)

#Status
status_label = tk.Label(root, text="Sheet loaded: Overall", font=("Arial", 12), bg="#001f3d", fg="white")
status_label.pack(pady=(5, 0))

# Search
search_frame = tk.Frame(root, bg="#001f3d")
search_frame.pack(pady=10)

search_label = tk.Label(search_frame, text="Search:", bg="#001f3d", font=("Arial", 12), fg="white")
search_label.pack(side="left", padx=(10, 5))

search_var = tk.StringVar()
search_entry = tk.Entry(search_frame, textvariable=search_var, width=30, font=("Arial", 12), bd=2, bg="#1d1d1d", fg="white")
search_entry.pack(side="left", padx=(0, 10))

# Search and Reset Functions Button
def search_players():
    query = search_var.get().strip()
    if not query:
        messagebox.showinfo("Search", "Please enter a search term.")
        return

    try:
        filtered_df = data[data.apply(lambda row: query.lower() in str(row).lower(), axis=1)]  # Filter rows based on the search term
        if filtered_df.empty:
            messagebox.showinfo("Search", "No players matched your search.")
        else:
            update_table(filtered_df)  # Update table with the filtered data
            status_label.config(text=f"Search results for: '{query}'")
    except Exception as e:
        messagebox.showerror("Search Error", f"Failed to search players: {e}")

def reset_table():
    update_table(data)  # Reset the table to the full data of the current sheet
    search_var.set("")
    status_label.config(text=f"Reset to full sheet: {current_sheet}")

# Search and reset buttons
search_button = tk.Button(search_frame, text="Search", command=search_players, bg="#004080", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
search_button.pack(side="left", padx=10)

reset_button = tk.Button(search_frame, text="Reset", command=reset_table, bg="#2c3e50", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
reset_button.pack(side="left", padx=5)

#Table Frame
tree_frame = tk.Frame(root)
tree_frame.pack(expand=True, fill='both', padx=10, pady=20)

tree_scroll = tk.Scrollbar(tree_frame)
tree_scroll.pack(side='right', fill='y')

tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode='browse', style="Treeview")
tree.pack(expand=True, fill='both')
tree_scroll.config(command=tree.yview)

#Treeview Styling
style = ttk.Style()
style.configure("Treeview", font=("Arial", 11), rowheight=30, background="#2c3e50", fieldbackground="#2c3e50", foreground="white")
style.map('Treeview', background=[('selected', '#004080')])

style.configure("Treeview.Heading", font=("Arial", 12, "bold"), foreground="white", background="#001f3d")

# Update Table Function
def update_table(df):
    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor='center')

    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

    status_label.config(text=f"Displaying {len(df)} players from '{current_sheet}'")

#Sorting Functions
def sort_by_rank():
    global data
    try:
        sorted_df = data.sort_values("RK")  # Sort data by the "RK" column
        update_table(sorted_df)
        status_label.config(text=f"Sorted by Rank: {current_sheet}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not sort by Rank: {e}")

#Next Best and Undo Functions
def highlight_next_best():
    global current_index, taken_players
    try:
        sorted_df = data.sort_values("RK").reset_index(drop=True)
    except Exception as e:
        messagebox.showerror("Error", f"Could not load or sort sheet: {e}")
        return

    while current_index < len(sorted_df) and sorted_df.loc[current_index, "PLAYER NAME"] in taken_players:
        current_index += 1

    if current_index >= len(sorted_df):
        messagebox.showinfo("End", "No more available players.")
        return

    player_name = sorted_df.loc[current_index, "PLAYER NAME"]
    taken_players.add(player_name)
    current_index += 1
    update_table(sorted_df)

def undo_last():
    global current_index, taken_players
    if not taken_players:
        messagebox.showinfo("Undo", "No previous selection to undo.")
        return

    current_index -= 1
    taken_players.discard(sorted_df.loc[current_index, "PLAYER NAME"])
    update_table(sorted_df)

#Buttons Frame
nav_frame = tk.Frame(root, bg="#001f3d")
nav_frame.pack(pady=20)

sort_button = tk.Button(nav_frame, text="Sort by Rank", command=sort_by_rank, bg="#004080", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
sort_button.pack(side="left", padx=15)

next_button = tk.Button(nav_frame, text="Next Best", command=highlight_next_best, bg="#2c3e50", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
next_button.pack(side="left", padx=15)

undo_button = tk.Button(nav_frame, text="Undo", command=undo_last, bg="#e74c3c", fg="white", font=("Arial", 12, "bold"), width=15, height=2)
undo_button.pack(side="left", padx=15)

# Load the first sheet initially
update_table(data)

root.mainloop()



