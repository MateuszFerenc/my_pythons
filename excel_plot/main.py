import matplotlib.pyplot as plotter
from openpyxl import load_workbook
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.colorchooser import askcolor

def print_help():
    print("Plot generator from excel options (2023) Mateusz Ferenc:")
    print("\tq, quit, exit, leave - to exit")
    print("\tload - open file dialog window to select file to load data from")
    print("\tunload - delete loaded file from temporary memory")
    print("\tselect <row-start> <row-end> <column-start> <column-end> - select data within given range (column accepts excel column representation)")
    print("\tdata or data <data-set> - print loaded datasets (only informations) / print contents of <data-set>")
    print("\tclear or clear <data-set> - clear whole database / clear <data-set> entry")
    print("\tconfig plot or config <data-set> - configure plotter properties / configure <data-set> plot properties")
    print("\tgenerate - generate plot from given data")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    workbook = None
    current_sheet = None
    data = {}
    plot_properties = {
        "dpi": 500,
        "xlabel": None,
        "ylabel": None,
        "label_weight": "light",
        "label_style": "italic",
        "title": None,
        "title_weight": "bold",
        "title_style": "bold",
        "grid": False
    }
    ds_properties = {}
    plot_properties_ok = False
    # color and label

    while(True):
        command = input(">")
        command, command_data = command.split()[0].lower(), command.split()[1:]
        if command in ("h", "help"):
            print_help()
        elif command in ("q", "quit", "exit", "leave"):
            if workbook is not None:
                workbook.close
            exit()
        elif command == "load":
            ans = ""
            if workbook is not None:
                ans = input(f"{current_sheet.title} is currently loaded, unload? (y/n)")
            if ans.lower() in ("y", "yes", "yep") or workbook is None:
                path = askopenfilename(title="Choose excel file", filetypes=[("Excel files", ".xlsx")], parent=root)
                try:
                    workbook = load_workbook(path)
                    current_sheet = workbook.active
                    print(f"{current_sheet.title} loaded")
                except FileNotFoundError:
                    print(f"{path} file not found...")
                except Exception as e:
                    print(f"{e}")                   
        elif command == "unload":
            if workbook is not None:
                workbook.remove()
                print(f"{current_sheet.title} unloaded")
                workbook = None
            current_sheet = None
        elif command == "select":
            if len(command_data) == 4:
                row_from, row_to, col_from, col_to = int(command_data[0]), int(command_data[1]), command_data[2], command_data[3]
                col_from_normal = 1
                for c in range(len(col_from)):
                    col_from_normal += (ord(col_from[-(1 + c)]) - 65)*pow(26, c)
                col_to_normal = 1
                for c in range(len(col_to)):
                    col_to_normal += (ord(col_to[-(1 + c)]) - 65)*pow(26, c)
                temp_data = []
                data_is_valid = True
                if row_from <= row_to and col_from_normal <= col_to_normal:
                    try:
                        for row in range(row_from, row_to + 1):
                            for col in range(col_from_normal, col_to_normal + 1):
                                value = current_sheet.cell(row=row, column=col).value
                                if type(value) is int or type(value) is float:
                                    temp_data.append(round(value, 4))
                                else:
                                    print(f"value: {value} in row: {row}, col: {col} is not int of float type\nType: {type(value)}\nRemoving selected data...")
                                    data_is_valid = False
                                    raise Exception("data is not valid")
                    except Exception as e:
                        print(e)  
                    if data_is_valid:
                        data_name = input("Data is valid, enter name of data set\n?")
                        data[data_name] = temp_data  
                        label = input("Enter data label\n?")
                        ds_properties[data_name]["label"] = label
                        ds_properties[data_name]["color"] = askcolor(title="Choose plot color", parent=root)[1]
                else:
                    pass
            else:
                pass
        elif command == "data":
            if len(command_data) == 0:
                if len(data):
                    for data_set in data.keys():
                        print(f"Dataset: \"{data_set}\" of size: {len(data[data_set])} elements")
                else:
                    print(f"No data loaded.\nUse load and select commands to load new data")
        elif command == "clear":
            if len(command_data) == 0:
                data = {}
                plot_properties_ok = False
                ds_properties = {}
                print("All data cleared")
            elif len(command_data) == 1:
                if command_data[0] in data.keys():
                    del ds_properties[command_data[0]]
                    del data[command_data[0]]
                    print(f"\"{command_data[0]}\" entries was removed from database")
        elif command == "config":
            if len(command_data) == 1:
                if command_data[0] == "plot":
                    can_exit = False
                    while(not can_exit):
                        plot_properties["dpi"] = int(input("Enter plot DPI\n?"))
                        plot_properties["xlabel"] = int(input("Enter x (horizontal) label name\n?"))
                        plot_properties["ylabel"] = int(input("Enter y (vertical) label name\n?"))
                        font_weights = ('ultralight', 'light', 'normal', 'regular', 'book', 'medium', 'roman', 'semibold', 'demibold', 'demi', 'bold', 'heavy', 'extra bold', 'black')
                        w = int(input(f"Enter labels font weight (valid: {font_weights})\n?"))
                        if w in font_weights:
                            plot_properties["label_weight"] = w
                        else:
                            print(f"fontweight: {w} is not valid..\n\"bold\" used as default")
                            plot_properties["label_weight"] = "bold"
                        font_styles = ('normal', 'italic', 'oblique')
                        s = int(input(f"Enter labels font style (valid: {font_styles})\n?"))
                        if s in font_styles:
                            plot_properties["label_style"] = s
                        else:
                            print(f"fontstyle: {s} is not valid..\n\"light\" used as default")
                            plot_properties["label_style"] = "light"
                        plot_properties["title"] = int(input("Enter plot title\n?"))
                        w = int(input(f"Enter title font weight (valid: {font_weights})\n?"))
                        if w in font_weights:
                            plot_properties["title_weight"] = w
                        else:
                            print(f"fontweight: {w} is not valid..\n\"bold\" used as default")
                            plot_properties["title_weight"] = "bold"
                        s = int(input(f"Enter labels font style (valid: {font_styles})\n?"))
                        if s in font_styles:
                            plot_properties["title_style"] = s
                        else:
                            print(f"fontstyle: {s} is not valid..\n\"blod\" used as default")
                            plot_properties["title_style"] = "bold"
                        g = int(input("Enable grid? (valid: True, False)\n?"))
                        if g in ("True", "False"):
                            plot_properties["grid"] = g
                        else:
                            print(f"option: {g} is not valid..\n\"False\" used as default")
                        print()
                        for key in plot_properties.keys():
                            print(f"{key}: {plot_properties[key]}")
                        ans = input("Are you ok with these values? (y/n)?")
                        if ans.lower() in ("y", "yes", "yep"):
                            can_exit = True
                            plot_properties_ok = True
                            print("OK!")
                elif command_data[0] in data.keys():
                    can_exit = False
                    while(not can_exit):
                        for ds_name in data.keys():
                            print(f"{ds_name}:")
                            ds_properties[ds_name]["label"] = input(f"Enter data label for \"{ds_name}\"\n?")
                            ds_properties[ds_name]["color"] = askcolor(title=f"Choose plot color for \"{ds_name}\"", parent=root)[1]
                        for key in ds_properties.keys():
                            print(f"{key}:")
                            for k in ds_properties[key].keys():
                                print(f"\t{k}: {ds_properties[key][k]}")
                        ans = input("Are you ok with these values? (y/n)?")
                        if ans.lower() in ("y", "yes", "yep"):
                            can_exit = True
                            print("OK!")
        elif command == "generate":
            if plot_properties_ok:
                for data_set in data.keys():
                    plotter.plot(data[data_set], color=ds_properties[data_set]["color"], label=ds_properties[data_set]["label"])
                plotter.xlabel(xlabel=plot_properties["xlabel"], weight=plot_properties["label_weight"], style=plot_properties["label_style"])
                plotter.ylabel(ylabel=plot_properties["ylabel"], weight=plot_properties["label_weight"], style=plot_properties["label_style"])
                plotter.title(label=plot_properties["title"], weight=plot_properties["title_weight"], style=plot_properties["title_style"])
                plotter.grid(visible=plot_properties["grid"])
                plotter.legend()

                plotter.show()

                ans = input("Are you ok with this plot? (y/n)?")
                if ans.lower() in ("y", "yes", "yep"):
                    plotter.savefig("output.png", dpi=plot_properties["dpi"])
                plotter.clf()
                plotter.close(None)
            else:
                print("Use \"config plot\" command to configure plotter properites..")
        elif command == "exec": # debug-only
            print(data[0])
            try:
                exec(data[0])
            except Exception as e:
                print(e)
