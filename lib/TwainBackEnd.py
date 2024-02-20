from io import BytesIO
import twain
import PIL.ImageTk
import PIL.Image
import base64
import tkinter as tk
import secrets
import string

class ClassTwainBackEnd():
    manager = None
    source = None
    resolution = 300
    duplex = False

    dummy_window = tk.Tk()

    def matchMode(modeStr):
        if modeStr == 'bw':
            return twain.TWPT_BW
        if modeStr == 'gray':
            return twain.TWPT_GRAY
        if modeStr == 'color':
            return twain.TWPT_RGB

    def scan(self,scannerName,postDpi,mode,isDuplex,removeBlank):
        # sm = twain.SourceManager()
        print("looping scan...")
        imageList =[]
        index = 1
        boolDuplex = None
        boolRemove = None

        if isDuplex == 1:
            boolDuplex = True
        else:
            boolDuplex = False

        if removeBlank == 1:
            boolRemove = True
        else:
            boolRemove = False

            # this will show UI to allow user to select source
        self.open(self,name=scannerName)
        print("Open source...",scannerName)
        if self.source:
            try:
                self.source.set_capability(twain.ICAP_XRESOLUTION,twain.TWTY_FIX32, float(postDpi))
                self.source.set_capability(twain.ICAP_YRESOLUTION,twain.TWTY_FIX32, float(postDpi))
            except:
                print("Could not set resolution to '%s'" % postDpi)
                pass

            try:
                self.source.set_capability(twain.CAP_DUPLEXENABLED,twain.TWTY_BOOL, boolDuplex)
            except:
                print("Could not set duplex to '%s'" % isDuplex)
                pass

            try:
                self.source.set_capability(twain.ICAP_PIXELTYPE,twain.TWTY_UINT16,  self.matchMode(mode))
            except:
                print("Could not set mode to '%s'" % mode)
                pass

            # try:
            #     self.source.set_capability(twain.ICAP_AUTODISCARDBLANKPAGES,twain.TWTY_BOOL, boolRemove)
            # except:
            #     print("Could not set auto discard blank pages to '%s'" % removeBlank)
            #     pass

            try:
                self.source.request_acquire(0,0)
            except:
                print("AcquisitionError")
                return "Acquisition Error"
            while self.next(self):
                image = self.capture(self,index)
                if image is None :
                    break
                file64 = self.read_file_and_encode_base64(image)
                imageList.append(({"imageIndex":index,"base64Image":file64,"scanedPath":image }))
                index += 1
            self.close(self)
            return imageList
        else:
            print("User clicked cancel")

    def scannerList():
        listObjectArray = []
        index = 1
        dummy_window = tk.Tk()
        sourceList = twain.SourceManager(parent_window=dummy_window).GetSourceList()
        for scannerName in sourceList:
            listObjectArray.append({"id":index,"scanner":scannerName})
            index += 1
        return listObjectArray

    def open(self, name):
        dummy_window = tk.Tk()
        self.manager = twain.SourceManager(parent_window=dummy_window, ProductName=name)
        if not self.manager:
            return
        if self.source:
            self.source.destroy()
            self.source = None
        self.source = self.manager.OpenSource(name)
        if self.source:
            print("%s: %s" % (name, self.source.GetSourceName()))

    def open_from_main_thread(self, name):
        self.dummy_window.after(0, lambda: self.open(name))

    def next(self):
        try:
            self.source.GetImageInfo()
            return True
        except:
            return False
    def capture(self,index):
        random_string = ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(10))
        fileName = "C:/Applications/pythonProjectScanner/Image/test_"+random_string+".jpg"
        try:
            (handle, more_to_come) = self.source.XferImageNatively()
        except twain.excDSTransferCancelled:
            print("Scanner ran out of paper.")
            return None
        except:
            print("Error capturing image.")
            return None

        bmp_bytes = twain.dib_to_bm_file(handle)
        img = PIL.Image.open(BytesIO(bmp_bytes), formats=["bmp"])
        img_buffer = BytesIO()
        img.save(fileName, format="JPEG")
        twain.global_handle_free(handle)
        return fileName
    def close(self):
        del self.manager

    def read_file_and_encode_base64(file_path):
        try:
            # Open the file in binary read mode ('rb')
            with open(file_path, 'rb') as file:
                # Read the binary content of the file
                file_content_binary = file.read()

            # Encode the binary content as Base64
            file_content_base64 = base64.b64encode(file_content_binary).decode('utf-8')

            return file_content_base64

        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
            return None