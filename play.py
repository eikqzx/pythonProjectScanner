import twain
import tkinter

root = tkinter.Tk()
# Initialize TWAIN Source Manager
source_manager = twain.SourceManager(root)

# Get a list of available TWAIN sources
source_list = source_manager.GetSourceList()

# Now you can print or use the list of available sources
print("Available TWAIN sources:", source_list)
