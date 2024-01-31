from waitress import serve
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from TwainBackEnd import ClassTwainBackEnd

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/scan',methods=['POST'])
def indexx():
    scannerName = request.json.get('scannerName')
    postDpi = request.json.get('postDpi')
    modeStr = request.json.get('mode') # ‘bw’, ‘gray’, ‘color’
    isDuplex = request.json.get('isDuplex')
    removeBlank = request.json.get('removeBlank')
    list = ClassTwainBackEnd.scan(self=ClassTwainBackEnd,scannerName=scannerName,postDpi=postDpi,mode=modeStr,isDuplex=isDuplex,removeBlank=removeBlank)
    # print("list:",list)
    if list == None:
        list = "Scanner is Offline"
    return jsonify({"ImageList":list})

@app.route("/getList",methods=['GET'])
def getList():
    list = ClassTwainBackEnd.scannerList()
    # print("list:",list)
    return jsonify({"scannerList":list})

mode = "dev"
host = "0.0.0.0"

if __name__ == '__main__':
    if mode == "dev":
        app.run(host=host,port=5100, debug=True)
    else:
        print("Flask API is running at "+host)
        serve(app,host=host,port=5100)