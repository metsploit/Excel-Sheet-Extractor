import os
import re
import copy
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from tkinter import Tk, filedialog, BooleanVar, Entry, Button, Label, DISABLED, NORMAL, Canvas, Frame, Scrollbar
from tkinter import ttk

def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def select_excel_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )
    root.destroy()
    return file_path

def select_sheets(wb):
    dialog = Tk()
    dialog.title("Select Sheets to Export")
    dialog.geometry('550x450')

    ctrl_frame = Frame(dialog)
    ctrl_frame.pack(fill='x', pady=5)
    Button(ctrl_frame, text="Select All", command=lambda: toggle_all(True)).pack(side='left', padx=5)
    Button(ctrl_frame, text="Deselect All", command=lambda: toggle_all(False)).pack(side='left', padx=5)

    container = Frame(dialog)
    container.pack(fill='both', expand=True)
    canvas = Canvas(container)
    scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    include_vars = {}
    rename_vars = {}
    entry_map = {}

    scroll_frame.columnconfigure(2, weight=1)

    Label(scroll_frame, text="Include", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
    Label(scroll_frame, text="Sheet Name", font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5)
    Label(scroll_frame, text="Rename", font=('Arial', 10, 'bold')).grid(row=0, column=2, padx=5)

    for idx, name in enumerate(wb.sheetnames, start=1):
        inc_var = BooleanVar(value=True)
        include_vars[name] = inc_var
        chk = ttk.Checkbutton(scroll_frame, variable=inc_var)
        chk.grid(row=idx, column=0, padx=(5, 20), sticky='w')

        lbl = Label(scroll_frame, text=name)
        lbl.grid(row=idx, column=1, padx=(0, 20), sticky='w')

        ren_var = BooleanVar(value=False)
        rename_vars[name] = ren_var
        ren_chk = ttk.Checkbutton(
            scroll_frame,
            variable=ren_var,
            command=lambda n=name: entry_map[n].config(
                state=NORMAL if rename_vars[n].get() else DISABLED
            )
        )
        ren_chk.grid(row=idx, column=2, padx=(0, 20))

        entry = ttk.Entry(scroll_frame)
        entry.insert(0, name)
        entry.config(state=DISABLED)
        entry.grid(row=idx, column=3, sticky='we', padx=5)
        entry_map[name] = entry

    def toggle_all(state):
        for var in include_vars.values():
            var.set(state)

    def on_ok():
        dialog.selected = []
        for s in wb.sheetnames:
            if include_vars[s].get():
                out_name = entry_map[s].get() if rename_vars[s].get() else s
                dialog.selected.append((s, out_name))
        dialog.destroy()

    btn_frame = Frame(dialog)
    btn_frame.pack(fill='x', pady=5)
    Button(btn_frame, text="OK", command=on_ok).pack(pady=5)

    dialog.mainloop()
    return getattr(dialog, 'selected', [])

def copy_sheet_content(source_ws, target_ws, debug=False):
    # Copy all cells with exact styles and values without unhiding hidden rows/columns
    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target_ws.cell(row=cell.row, column=cell.column)
            new_cell.value = cell.value
            new_cell.number_format = cell.number_format
            try:
                new_cell.font = copy.copy(cell.font)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.border = copy.copy(cell.border)
                new_cell.alignment = copy.copy(cell.alignment)
                new_cell.protection = copy.copy(cell.protection)
            except:
                pass

    # Copy merged cells
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    # Copy row dimensions including hidden state
    for idx, rd in source_ws.row_dimensions.items():
        tgt_rd = target_ws.row_dimensions[idx]
        if rd.height is not None:
            tgt_rd.height = rd.height
        tgt_rd.hidden = rd.hidden

    # Copy column dimensions including hidden state
    default_width = source_ws.sheet_format.defaultColWidth or 8.43
    max_col = source_ws.max_column
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        src_cd = source_ws.column_dimensions.get(col_letter, openpyxl.worksheet.dimensions.ColumnDimension(source_ws, index=col_letter))
        tgt_cd = target_ws.column_dimensions[col_letter]
        # Preserve width and hidden
        tgt_cd.width = src_cd.width if src_cd.width is not None else default_width
        tgt_cd.hidden = src_cd.hidden

    # Copy sheet-level settings without unhiding
    target_ws.page_margins = copy.copy(source_ws.page_margins)
    target_ws.page_setup = copy.copy(source_ws.page_setup)
    target_ws.print_options = copy.copy(source_ws.print_options)
    target_ws.freeze_panes = source_ws.freeze_panes
    try:
        target_ws.sheet_properties = copy.copy(source_ws.sheet_properties)
    except:
        pass


def split_sheets_to_workbooks(file_path):
    if not file_path:
        print("No file selected.")
        return
    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        print("Error opening workbook:", e)
        return

    selections = select_sheets(wb)
    if not selections:
        print("No sheets selected. Exiting.")
        return

    folder = os.path.dirname(file_path)

    for sheet_name, out_name in selections:
        clean_name = clean_filename(out_name)
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = sheet_name

        # Copy default dimensions
        try:
            new_ws.sheet_format.defaultColWidth = wb[sheet_name].sheet_format.defaultColWidth
            new_ws.sheet_format.defaultRowHeight = wb[sheet_name].sheet_format.defaultRowHeight
        except:
            pass

        # Perform exact copy without unhiding
        copy_sheet_content(wb[sheet_name], new_ws, debug=False)

        path = os.path.join(folder, f"{clean_name}.xlsx")
        try:
            new_wb.save(path)
            print(f"✅ Created: {os.path.basename(path)}")
        except Exception as e:
            print(f"❌ Error saving {clean_name}.xlsx: {e}")

    print("🎉 Done!")

if __name__ == '__main__':
    try:
        root_path = select_excel_file()
        split_sheets_to_workbooks(root_path)
    except Exception as e:
        import traceback
        print("❌ An unexpected error occurred:")
        traceback.print_exc()
    input("\nPress Enter to exit...")
