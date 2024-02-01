from io import BytesIO
import twain
import PIL.ImageTk
import PIL.Image
import base64

class ClassTwainBackEnd():
    manager = None
    source = None
    resolution = 300
    duplex = False

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

            try:
                self.source.set_capability(twain.ICAP_AUTODISCARDBLANKPAGES,twain.TWTY_BOOL, boolRemove)
            except:
                print("Could not set auto discard blank pages to '%s'" % removeBlank)
                pass

            try:
                self.source.request_acquire(0,0)
            except:
                print("AcquisitionError")
                return ""
            while self.next(self):
                image = self.capture(self)
                imageList.append(({"imageIndex":index,"base64Image":base64.b64encode(image.getvalue()).decode("utf-8") }))
                index += 1
            self.close()
            return imageList
        else:
            print("User clicked cancel")

    def scannerList():
        listObjectArray = []
        sourceList = twain.SourceManager().GetSourceList()
        for scannerName in sourceList:
            scanner = twain.SourceManager(0,ProductName=scannerName).GetIdentity()
            print("scanner",scanner)
            listObjectArray.append(scanner)
        return listObjectArray

    def open(self, name):
        self.manager = twain.SourceManager(0,ProductName=name )
        if not self.manager:
            return
        if self.source:
            self.source.destroy()
            self.source=None
        self.source = self.manager.OpenSource(name)
        if self.source:
            print("%s: %s" % ( name, self.source.GetSourceName() ))

    def next(self):
        try:
            self.source.GetImageInfo()
            return True
        except:
            return False
    def capture(self):
            fileName = "tmp.tmp"
            try:
                (handle, more_to_come) = self.source.XferImageNatively()
            except:
                return None

            bmp_bytes = twain.dib_to_bm_file(handle)
            img = PIL.Image.open(BytesIO(bmp_bytes), formats=["bmp"])
            img_buffer = BytesIO()
            img.save(img_buffer, format="BMP")
            twain.global_handle_free(handle)
            return img_buffer

    def close(self):
        del self.manager

