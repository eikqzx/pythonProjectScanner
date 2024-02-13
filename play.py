import twain
import tkinter
import os

root = tkinter.Tk()
# Initialize TWAIN Source Manager
path = os.path.join(os.environ["WINDIR"], 'twain_32.dll')
print(path)
source_manager = twain.SourceManager()


# Get a list of available TWAIN sources
source_list = source_manager.GetSourceList()

# Now you can print or use the list of available sources
print("Available TWAIN sources:", source_list)
