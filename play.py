import twain
import tkinter
import os

root = tkinter.Tk()
# Initialize TWAIN Source Manager
path = os.path.join(os.environ["WINDIR"], 'twain_32.dll')
print(path)
source_manager = twain.SourceManager(parent_window=1)


# Get a list of available TWAIN sources
source_list = source_manager.GetSourceList()

open = source_manager.open_source("Plustek A3 ADF Series Scanner-TWA")

# Now you can print or use the list of available sources
print("Available TWAIN sources:", source_list)
print("Open",open)
