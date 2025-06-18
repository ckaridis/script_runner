import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import subprocess
import os
import sys
import configparser
import send2trash
import shutil
import ctypes
import win32com.client
from datetime import datetime
from functools import partial
from PIL import Image, ImageTk

class BatFileRunner:
    def __init__(self):
        self.app_version = "1.1.2"
        self.build_date = self.get_build_date()

        self.bat_files_folder = ""
        self.search_keyword = ""
        self.sections = {}
        self.message_frame = None
        self.max_sections_per_row = 4  # Default value for max sections per row

        self.skip_validation = True  # Track the state of the "Skip Validation" checkbox

        self.config = configparser.ConfigParser()
        self.config.read("config.cfg")
        self.bat_files_folder = self.config.get("Settings", "bat_files_folder")

        if self.config.has_option("Settings", "max_sections_per_row"):
            self.max_sections_per_row = self.config.getint("Settings", "max_sections_per_row")

        self.root = tk.Tk()
        self.root.title("Script Runner " + self.app_version)
        self.setup_ui()

    def get_build_date(self):
        # If the executable is frozen (e.g., using PyInstaller)
        if getattr(sys, 'frozen', False):
            # Use the creation time of the executable file
            return datetime.fromtimestamp(os.path.getctime(sys.executable)).strftime("%Y%m%d%H%M%S")
        else:
            # If running as a script, return the current time
            return datetime.now().strftime("%Y%m%d%H%M%S")

    def run_bat(self, file_path):
        if not self.skip_validation:
            subprocess.Popen(["start", "cmd", "/k", file_path], shell=True)
        else:
            if messagebox.askokcancel("Run Script", f"Are you sure you want to run {os.path.basename(file_path)}?"):
                subprocess.Popen(["start", "cmd", "/k", file_path], shell=True)

    def run_bat_as_admin(self, file_path):
        if messagebox.askokcancel("Run Script as Admin", f"Are you sure you want to run {os.path.basename(file_path)} with elevated rights?"):
            if sys.platform == 'win32':
                # Separate each argument as an element in the array
                cmd_args = ["/k", file_path]
                # Properly format the arguments for PowerShell
                formatted_args = ','.join(['"{}"'.format(arg) for arg in cmd_args])
                subprocess.Popen(["powershell", "Start-Process", "cmd", "-ArgumentList", formatted_args, "-Verb", "runAs"])
            else:
                messagebox.showinfo("Not Supported", "This feature is only supported on Windows.")

    def toggle_skip_validation(self):
        self.skip_validation = not self.skip_validation
    
    def edit_script(self, file):
        file_path = os.path.join(self.bat_files_folder, file)
        if file.endswith(".lnk"):
            target_path = self.get_target_from_shortcut(file_path)
            if target_path and os.path.exists(target_path) and target_path.endswith(".bat"):
                file_path = target_path  # Use the target path only if it's a .bat file
        try:
            if sys.platform == "win32":
                os.startfile(file_path, "edit")
            else:
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit the script: {str(e)}")

    def set_default_folder(self):
        selected_folder = filedialog.askdirectory(title="Select Default Folder")
        if selected_folder:
            self.bat_files_folder = selected_folder
            self.config.set("Settings", "bat_files_folder", self.bat_files_folder)
            with open("config.cfg", "w") as config_file:
                self.config.write(config_file)
            self.refresh_ui()

    def set_max_sections_per_row(self):
        new_value = simpledialog.askinteger("Set Columns", "Enter max number of sections per row:",
                                            minvalue=1, maxvalue=10, initialvalue=self.max_sections_per_row)
        if new_value is not None and new_value != self.max_sections_per_row:
            self.max_sections_per_row = new_value
            self.config.set("Settings", "max_sections_per_row", str(new_value))
            with open("config.cfg", "w") as config_file:
                self.config.write(config_file)
            self.refresh_ui()

    def search_files(self, event):
        self.search_keyword = self.search_entry.get().lower()
        self.refresh_ui()

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.search_keyword = ""
        self.refresh_ui()

    def create_button(self, frame, text, command):
        button = tk.Button(frame, text=text, width=self.max_button_width - 10, command=command)
        button.pack(side=tk.LEFT)

        # Create a right-click menu for the button
        button_menu = tk.Menu(button, tearoff=0)

        # Add commands with icons to the menu
        button_menu.add_command(label="Rename", image=self.rename_icon, compound=tk.LEFT, command=lambda: self.rename_script(os.path.join(self.bat_files_folder, text)))
        # button_menu.add_command(label="Run as Admin", image=self.admin_icon, compound=tk.LEFT, command=lambda: self.run_bat_as_admin(os.path.join(self.bat_files_folder, text)))
        button_menu.add_command(label="Edit", image=self.edit_icon, compound=tk.LEFT, command=lambda: self.edit_script(os.path.join(self.bat_files_folder, text)))
        button_menu.add_command(label="Delete", image=self.trash_icon, compound=tk.LEFT, command=lambda: self.delete_script(os.path.join(self.bat_files_folder, text)))
        button_menu.add_command(label="Duplicate", image=self.duplicate_icon, compound=tk.LEFT, command=lambda: self.duplicate_script(text))

        # Bind the right-click event to the button
        button.bind("<Button-3>", lambda event: button_menu.post(event.x_root, event.y_root))

        # Bind tooltip functionality
        button.bind("<Enter>", lambda event, file=text: self.show_tooltip(event, file))
        button.bind("<Leave>", self.hide_tooltip)


    def show_tooltip(self, event, file):
        tooltip_text = self.get_script_preview(file)
        x, y, _, _ = event.widget.bbox("insert")
        x += event.widget.winfo_rootx() + 25
        y += event.widget.winfo_rooty() + 25
        self.tooltip = tk.Toplevel(event.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=tooltip_text, background="white", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event):
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()

    def get_script_preview(self, file):
        file_path = os.path.join(self.bat_files_folder, file)
        try:
            if file.endswith(".bat"):
                with open(file_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                return "".join(lines[:20])  # Display first 20 lines of the bat file
            elif file.endswith(".lnk"):
                target_path = self.get_target_from_shortcut(file_path)
                if target_path and os.path.exists(target_path):
                    with open(target_path, "r", encoding="utf-8") as f:
                        lines = f.readlines()
                    return "".join(lines[:20])  # Display first 20 lines of the target file
                else:
                    return "Target file not found"
            else:
                return "Unsupported file format"
        except Exception as e:
            return f"Error reading file: {str(e)}"

    def get_target_from_shortcut(self, lnk_file):
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(lnk_file)
            return shortcut.TargetPath
        except Exception as e:
            return f"Error reading shortcut: {str(e)}"


    def duplicate_script(self, file_name):
        old_path = os.path.join(self.bat_files_folder, file_name)
        new_name = file_name.split('.')[0] + '_copy.bat'
        new_path = os.path.join(self.bat_files_folder, new_name)
        try:
            shutil.copy(old_path, new_path)
            self.refresh_ui()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to duplicate the script: {str(e)}")

    def refresh_ui(self):

        if hasattr(sys, '_MEIPASS'):
            # Running as a bundled executable
            img_path = os.path.join(sys._MEIPASS, "res/img")
        else:
            # Running as a regular Python script
            img_path = "res/img"

        # Update footer with statistics
        total_scripts = len(self.bat_files)
        self.footer_label.config(text=f"Total Scripts and Links: {total_scripts} | App Version: {self.app_version} | Build: {self.build_date}")

        # Load and resize the trash bin icon to 16x16
        self.trash_icon = ImageTk.PhotoImage(Image.open(os.path.join(img_path, "recycle-bin.png")).resize((20, 20), Image.ADAPTIVE))
        self.edit_icon = ImageTk.PhotoImage(Image.open(os.path.join(img_path, "edit.png")).resize((20, 20), Image.ADAPTIVE))
        self.rename_icon = ImageTk.PhotoImage(Image.open(os.path.join(img_path, "edit-text.png")).resize((20, 20), Image.ADAPTIVE))
        self.admin_icon = ImageTk.PhotoImage(Image.open(os.path.join(img_path, "admin.png")).resize((20, 20), Image.ADAPTIVE))
        self.duplicate_icon = ImageTk.PhotoImage(Image.open(os.path.join(img_path, "duplicate.png")).resize((20, 20), Image.ADAPTIVE))

        img = Image.open(os.path.join(img_path, "icon.png"))
        icon = ImageTk.PhotoImage(img)
        self.root.iconphoto(True, icon)

        for section in self.sections.values():
            section.destroy()

        if self.message_frame: 
            self.message_frame.destroy()
            self.message_frame = None

        self.bat_files = sorted([file for file in os.listdir(self.bat_files_folder) if file.endswith((".bat", ".lnk"))])

        if not self.bat_files:
            self.display_empty_list_message()
            return

        self.sections = {}
        section_row = 0
        section_column = 0

        for file_name in self.bat_files:
            if '-' in file_name:
                section_name = file_name.split('-')[0]
            else:
                section_name = "#no_section"

            if section_name not in self.sections:
                self.sections[section_name] = self.create_section_frame(section_name, section_row, section_column)
                section_column += 1
                if section_column >= self.max_sections_per_row:
                    section_column = 0
                    section_row += 1

            if not self.search_keyword or self.search_keyword in file_name.lower():
                file_path = os.path.join(self.bat_files_folder, file_name)
                button_frame = tk.Frame(self.sections[section_name])
                button_frame.pack(padx=10, pady=2, fill=tk.X)

                self.create_button(button_frame, file_name, partial(self.run_bat, file_path))

                # rename_button = tk.Button(button_frame, image=self.rename_icon, command=partial(self.rename_script, file_path))
                # rename_button.pack(side=tk.LEFT)
                # Tooltip(rename_button, "Rename script")

                run_as_admin_button = tk.Button(button_frame, image=self.admin_icon, command=partial(self.run_bat_as_admin, file_path))
                run_as_admin_button.pack(side=tk.LEFT)
                Tooltip(run_as_admin_button, "Run script with elevated rights")

                # edit_button = tk.Button(button_frame, image=self.edit_icon, command=partial(self.edit_script, file_path))
                # edit_button.pack(side=tk.RIGHT)
                #  # Create a tooltip for the Edit button
                # Tooltip(edit_button, "Edit script")

                # delete_button = tk.Button(button_frame, image=self.trash_icon, command=partial(self.delete_script, file_path))
                # delete_button.image = self.trash_icon  # Keep a reference
                # delete_button.pack(side='left')
                # Tooltip(delete_button, "Move script to trash")


    def rename_script(self, old_path):
        old_name = os.path.basename(old_path)
        _, old_extension = os.path.splitext(old_name)  # Extract the extension from the old name
        
        new_name = simpledialog.askstring("Rename Script", "Enter the new name for the script:",
                                        initialvalue=old_name, parent=self.root)
        
        if new_name:
            new_name_extension = os.path.splitext(new_name)[1]
            
            # If the user did not include an extension in the new name, append the old one
            if not new_name_extension:
                new_name += old_extension
            
            if new_name != old_name:
                new_path = os.path.join(self.bat_files_folder, new_name)
                
                # Validate old_path exists before attempting to rename
                if not os.path.exists(old_path):
                    messagebox.showerror("Error", f"The file {old_name} does not exist.")
                    return
                
                try:
                    os.rename(old_path, new_path)
                    self.refresh_ui()
                except PermissionError as e:
                    messagebox.showerror("Error", "The file is currently open or in use. Please close it and try again.")
                except FileNotFoundError as e:
                    messagebox.showerror("Error", f"Failed to rename the script because the file was not found: {str(e)}")
                except Exception as e:
                    messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

                self.refresh_ui()

    def delete_script(self, file_path):
        # Normalize the file path
        normalized_path = os.path.normpath(file_path)

        # Check if the file exists
        if not os.path.exists(normalized_path):
            messagebox.showerror("Error", f"The file {os.path.basename(file_path)} does not exist.")
            return

        if messagebox.askokcancel("Delete Script", f"Are you sure you want to delete {os.path.basename(file_path)}?", icon='warning'):
            send2trash.send2trash(normalized_path)
            self.refresh_ui()

    def display_empty_list_message(self):
        if self.message_frame: 
            self.message_frame.destroy()

        self.message_frame = tk.Frame(self.main_frame)
        self.message_frame.pack(expand=True)
    
        message_label = tk.Label(self.message_frame, text="The list is empty, please select a script folder using the top bar menu", font=("Arial", 12))
        message_label.pack(side=tk.TOP, expand=True)

    def create_section_frame(self, section_name, row, column):
        section_frame = tk.LabelFrame(self.main_frame, text=section_name, font=("Arial", 12, "bold"))
        section_frame.grid(row=row, column=column, padx=10, pady=5, sticky="nsew")
        return section_frame

    def setup_ui(self):
        self.root.resizable(False, False)

        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, expand=True)

        # Center frame for the search bar
        center_frame = tk.Frame(self.top_frame)
        center_frame.pack(side=tk.LEFT, expand=True)

        search_label = tk.Label(center_frame, text="Search:")
        search_label.pack(side=tk.LEFT)

        self.search_entry = tk.Entry(center_frame, width=40)
        self.search_entry.pack(side=tk.LEFT)
        self.search_entry.bind("<KeyRelease>", self.search_files)

        clear_button = tk.Button(center_frame, text="Clear", command=self.clear_search)
        clear_button.pack(side=tk.LEFT)

        tk.Label(self.top_frame).pack()  # Add a space between the search components and the checkbox

        # Menu bar
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Add "Open Scripts Folder" option to open the scripts folder defined in settings
        file_menu.add_command(label="New Script", command=self.create_new_script)  # Add New Script option
        file_menu.add_command(label="Open Scripts Folder", command=self.open_scripts_folder)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Set Default Folder", command=self.set_default_folder)
        settings_menu.add_command(label="Set Max Sections Per Row", command=self.set_max_sections_per_row)
        settings_menu.add_checkbutton(label="Skip Validation", variable=self.skip_validation, command=self.toggle_skip_validation)

        menu_bar.add_command(label="Refresh", command=self.refresh_ui)

        # About menu
        about_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="About", menu=about_menu)
        about_menu.add_command(label="About Script Runner", command=self.show_about_window)

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        if not os.listdir(self.bat_files_folder):
            self.set_default_folder()

        self.bat_files = sorted([file for file in os.listdir(self.bat_files_folder) if file.endswith((".bat", ".lnk"))])

        if self.bat_files:
            self.max_button_width = max(len(file_name) for file_name in self.bat_files)
        else:
            self.max_button_width = 30  # Default width in case there are no files

        # Footer for statistics including version and build date
        self.footer_label = tk.Label(self.root, text=f"Version: {self.app_version} | Build: {self.build_date}")
        self.footer_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.refresh_ui()

    def open_scripts_folder(self):
        # Get the absolute path of the scripts folder defined in settings
        scripts_folder = os.path.abspath(self.bat_files_folder)

        # Open the scripts folder using the default file explorer
        os.system(f'start "" "{scripts_folder}"')

    def create_new_script(self):
        new_script_name = simpledialog.askstring("New Script", "Enter the name of the new script:", parent=self.root)
        if new_script_name:
            if not new_script_name.endswith(".bat"):
                new_script_name += ".bat"
            new_script_path = os.path.join(self.bat_files_folder, new_script_name)
            # Create the new script file
            with open(new_script_path, 'w') as new_file:
                new_file.write("@echo off\n")
            # Refresh the UI to display the new script
            self.refresh_ui()

            # Ask if the user wants to edit the new script
            if messagebox.askyesno("Edit Script", "Do you want to edit the new script?"):
                self.edit_script(new_script_path)

    # Function to display the "About" window
    def show_about_window(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("About Script Runner")
        about_window.geometry("600x100")

        version_label = tk.Label(about_window, text=f"Script Runner Version {self.app_version}\nPython Version {sys.version}")
        version_label.pack(padx=20, pady=20)

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=self.text, background="white", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()

if __name__ == "__main__":
    app = BatFileRunner()
    app.root.mainloop()
