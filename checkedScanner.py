import win32com.client
import tkinter as tk
from tkinter import messagebox, filedialog
import pythoncom
from PIL import Image, ImageTk
from datetime import datetime
import os

# Declare the scanner variable globally
scanner = None
scanned_files = []

def get_connected_scanners():
    wia = win32com.client.Dispatch("WIA.CommonDialog")

    try:
        # Use the WIA CommonDialog to show the device selection dialog
        dev = wia.ShowSelectDevice()

        # Get the name of the selected device (scanner)
        scanner_name = dev.Properties("Name").Value

        return dev, scanner_name
    except Exception as e:
        print(f"Error: {e}")
        return None, None

def scan_document():
    global scanner
    try:
        # Check if scanner is available
        if scanner:
            # Prompt user to select a directory for saving the scanned file
            save_directory = filedialog.askdirectory(title="Select Save Directory")
            if not save_directory:
                return  # User canceled directory selection

            # Invoke the WIA interfaces directly to perform scanning
            image_item = scanner.Items[1]  # Assuming the first item is the scanned image
            image_item.Properties("6146").Value = 2  # Set color mode to 2 (Color)
            image_item.Properties("6147").Value = 300  # Set resolution to 300 DPI

            image = image_item.Transfer()  # Transfer the scanned image

            # Generate a unique file name based on the current datetime
            current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")
            file_path = f"{save_directory}/scanned_image_{current_datetime}.jpg"
            image.SaveFile(file_path)

            # Add the scanned file path to the list
            scanned_files.append(file_path)

            # Show the scanned image
            show_scanned_image(file_path)

            messagebox.showinfo("Scanning Complete", f"Document scanned and saved to:\n{file_path}")
        else:
            messagebox.showerror("No Scanner", "Scanner not available.")
    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Scanning Error", "Error occurred during scanning.")

def show_scanned_image(image_path):
    image = Image.open(image_path)
    image.show()

def open_selected_file(listbox):
    selected_indices = listbox.curselection()
    if selected_indices:
        selected_index = selected_indices[0]
        if 0 <= selected_index < len(scanned_files):
            selected_file = scanned_files[selected_index]
            if os.path.exists(selected_file):
                os.startfile(selected_file)
            else:
                messagebox.showerror("File Not Found", f"The selected file '{selected_file}' does not exist.")
        else:
            messagebox.showerror("Invalid Selection", "Invalid file selection.")
    else:
        messagebox.showinfo("No Selection", "No file selected.")

def show_scanned_files():
    # Display the list of scanned files
    if scanned_files:
        files_window = tk.Toplevel(root)
        files_window.title("Scanned Files")

        label = tk.Label(files_window, text="List of Scanned Files:")
        label.pack(pady=10)

        listbox = tk.Listbox(files_window)
        for file_path in scanned_files:
            listbox.insert(tk.END, file_path)
        listbox.pack(pady=20)

        open_button = tk.Button(files_window, text="Open Selected File", command=lambda: open_selected_file(listbox))
        open_button.pack(pady=10)

        close_button = tk.Button(files_window, text="Close", command=files_window.destroy)
        close_button.pack(pady=10)
    else:
        messagebox.showinfo("No Scanned Files", "No files have been scanned yet.")

def show_scanner_operations(scanner_name):
    global scanner
    operations_window = tk.Toplevel(root)
    operations_window.title("Scanner Operations")

    label = tk.Label(operations_window, text=f"Scanner: {scanner_name}")
    label.pack(pady=20)

    scan_button = tk.Button(operations_window, text="Scan Document", command=scan_document)
    scan_button.pack(pady=10)

    show_files_button = tk.Button(operations_window, text="Show Scanned Files", command=show_scanned_files)
    show_files_button.pack(pady=10)

    close_button = tk.Button(operations_window, text="Close", command=operations_window.destroy)
    close_button.pack(pady=10)

def show_scanner_gui():
    global scanner
    scanner, scanner_name = get_connected_scanners()

    if scanner:
        messagebox.showinfo("Connected Scanner", f"Connected Scanner: {scanner_name}")
        show_scanner_operations(scanner_name)
    else:
        messagebox.showinfo("No Scanner", "No scanner found or selection canceled.")

if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    root.title("Scanner Detector")

    # Create a button to trigger the scanner detection
    detect_button = tk.Button(root, text="Detect Scanner", command=show_scanner_gui)
    detect_button.pack(pady=20)

    # Run the Tkinter event loop
    root.mainloop()
