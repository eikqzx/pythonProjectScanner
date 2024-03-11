from waitress import serve
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from lib.TwainBackEnd import ClassTwainBackEnd

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/scan',methods=['POST'])
def indexx():
    scannerName = request.json.get('scannerName')
    postDpi = request.json.get('postDpi')
    modeStr = request.json.get('mode') # ‘bw’, ‘gray’, ‘color’
    isDuplex = request.json.get('isDuplex')
    removeBlankInt = request.json.get('removeBlank')
    ADF = request.json.get('isADF')
    list = ClassTwainBackEnd.scan(self=ClassTwainBackEnd,scannerName=scannerName,postDpi=postDpi,mode=modeStr,isDuplex=isDuplex,removeBlank=removeBlankInt,isADF=ADF)
    # print("list:",list)
    if list == None:
        list = "Scanner is Offline"
    return jsonify({"ImageList":list})

@app.route("/getList",methods=['GET'])
def getList():
    list = ClassTwainBackEnd.scannerList()
    # print("list:",list)
    return jsonify({"scannerList":list})

@app.route('/createFile',methods=['POST'])
def createFileByPath():
    path = request.json.get('path')
    fileBase64 = request.json.get('fileBase64')
    if os.path.exists(path):
        return {"result":False,"msg":f"File already exists at: {path}"}
    else:
        # Write the decoded image data to the file
        decodeStr = base64.b64decode(fileBase64)
        with open(path, 'wb') as file:
            file.write(decodeStr)
        return {"result":True,"msg":f"File created at: {path}"}

mode = "dev"
host = "0.0.0.0"

if __name__ == '__main__':
    if mode == "dev":
        app.run(host=host,port=5100, debug=True)
    else:
        print("Flask API is running at "+host)
        serve(app,host=host,port=5100)