import win32com.client
import tkinter as tk
from tkinter import messagebox, filedialog
import pythoncom
from PIL import Image, ImageTk

# Declare the scanner variable globally
scanner = None

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

            # Save the scanned image to the selected directory
            file_path = f"{save_directory}/scanned_image.jpg"
            image.SaveFile(file_path)

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

def show_scanner_operations(scanner_name):
    global scanner
    operations_window = tk.Toplevel(root)
    operations_window.title("Scanner Operations")

    label = tk.Label(operations_window, text=f"Scanner: {scanner_name}")
    label.pack(pady=20)

    scan_button = tk.Button(operations_window, text="Scan Document", command=scan_document)
    scan_button.pack(pady=10)

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
