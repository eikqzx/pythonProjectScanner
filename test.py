from io import BytesIO
import twain
import tkinter
from tkinter import ttk
import logging
import PIL.ImageTk
import PIL.Image
import base64
from waitress import serve
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import imghdr


app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/scan',methods=['POST'])
def indexx():
    scannerName = request.json.get('scannerName')
    postDpi = request.json.get('postDpi')
    numPage = request.json.get('numPage')
    modeStr = request.json.get('mode')
    list = scan(scannerName,postDpi,modeStr,numPage)
    # print("list:",list)
    return jsonify({"ImageList":list})

def matchMode(modeStr):
    if modeStr == 'bw':
        return twain.TWPT_BW
    if modeStr == 'gray':
        return twain.TWPT_GRAY
    if modeStr == 'color':
        return twain.TWPT_RGB

def scan(scannerName,postDpi,mode,numPage):
    imageList =[]
    index = 1
    with twain.SourceManager() as sm:
        # this will show UI to allow user to select source
        src = sm.open_source(scannerName)
        if src:
            src.set_capability(twain.ICAP_XRESOLUTION,twain.TWTY_FIX32, float(postDpi))
            src.set_capability(twain.ICAP_YRESOLUTION,twain.TWTY_FIX32, float(postDpi))
            src.set_capability(twain.CAP_FEEDERENABLED, twain.TWTY_FIX32, float(postDpi))
            #  can be ‘bw’, ‘gray’, ‘color’
            src.set_capability(twain.ICAP_PIXELTYPE,twain.TWTY_UINT16,  matchMode(mode))

            src.request_acquire(show_ui=False, modal_ui=False)
            (handle, remaining_count) = src.xfer_image_natively()

            # while index <= numPage:
            bmp_bytes = twain.dib_to_bm_file(handle)
            img = PIL.Image.open(BytesIO(bmp_bytes), formats=["bmp"])
            img_buffer = BytesIO()
            img.save(img_buffer, format="BMP")
            imageList.append(({"imageIndex":remaining_count,"base64Image":base64.b64encode(img_buffer.getvalue()).decode("utf-8") }))
            # index += 1

            return imageList
        else:
            print("User clicked cancel")

@app.route("/getList",methods=['GET'])
def getList():
    list = scannerList()
    # print("list:",list)
    return jsonify({"scannerList":list})

def scannerList():
    sourceList = twain.SourceManager().GetSourceList()
    return sourceList

mode = "dev"
host = "0.0.0.0"

if __name__ == '__main__':
    if mode == "dev":
        app.run(host=host,port=5100, debug=True)
    else:
        print("Flask API is running at "+host)
        serve(app,host=host,port=5100)

# logging.basicConfig(level=logging.DEBUG)
# root = tkinter.Tk()
# frm = ttk.Frame(root, padding=10)
# frm.grid()
# ttk.Button(frm, text="Scan", command=scan).grid(column=0, row=0)
# root.mainloop()