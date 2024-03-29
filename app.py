from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import win32com.client
from datetime import datetime
import pythoncom
import io
import base64
from waitress import serve
import twain
import tempfile
import pyinsane2
import asyncio

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

@app.route('/')
def index():
    return "API is running on port 5100"

@app.route('/api/scannerList', methods=['GET'])
def get_scanner_list():
    try:
        scanner_list = get_connected_scanners()
        return jsonify({"scannerList": scanner_list})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": "Error occurred while fetching scanner list."})

def get_connected_scanners():
    pythoncom.CoInitialize()  # Initialize COM

    wia = win32com.client.Dispatch("WIA.CommonDialog")

    try:
        dev = wia.ShowSelectDevice()

        scanner_id = dev.DeviceID
        scanner_name = dev.Properties("Name").Value

        return [{'id': scanner_id, 'name': scanner_name}]

    except Exception as e:
        print(f"Error: {e}")
        return []
    finally:
        pythoncom.CoUninitialize()

@app.route('/api/scan', methods=['POST'])
def scan_document():
    try:
        scanner_id = request.json.get('scannerId')
        dpi = request.json.get('dpi')
        pages = request.json.get('pages')
        scan_result = perform_scan(scanner_id,dpi,pages)
        # print(scan_result)
        if scan_result:
            return jsonify({"status":True,"message": "Scanning complete", "scannedImageBase64": scan_result})
        else:
            return jsonify({"status":False,"error": "Error occurred during scanning."})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"status":False,"error": "Error occurred during scanning."})


def perform_scan(scanner_id,dpi,pages):
    pythoncom.CoInitialize()  # Initialize COM

    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        dev = wia.ShowSelectDevice()
        imageList = []
        if dev.DeviceID == scanner_id:
            i = 1
            while i <= pages:
                # dev.Items(1).Properties("Horizontal Resolution").Value = dpi
                # dev.Items(1).Properties("Vertical Resolution").Value = dpi

                image_format = '{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}'  # WIA_FORMAT_JPEG
                image_data = dev.Items(1).Transfer(image_format).FileData.BinaryData

                imageList.append(base64.b64encode(image_data).decode('utf-8'))
                i+=1
            return imageList  # Return the base64-encoded image
        else:
            return None
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

@app.route('/api/getProperties', methods=['POST'])
def get_Handle():
    try:
        scanner_id = request.json.get('scannerId')
        dpi = request.json.get('dpi')
        pages = request.json.get('pages')
        scan_result = getProperties(scanner_id,dpi,pages)
        if scan_result:
            return jsonify({"Properties": scan_result})
        else:
            return jsonify({"error": "Scanner is offline."})
    except Exception as e:
        print(f"Error: {e}")
        return False

def getProperties(scanner_id,dpi,pages):
    pythoncom.CoInitialize()  # Initialize COM
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        dev = wia.ShowSelectDevice()
        if dev.DeviceID == scanner_id:
            properties_info = []
            for prop in dev.Properties:
                property_info = {
                    "Name": prop.Name,
                    "PropertyID": prop.PropertyID,
                    "Type": prop.Type,
                    "Value": prop.Value,
                }
                properties_info.append(property_info)

            return properties_info
            # return dev.Properties.Exists("Document Handling Select")
        else:
            return None
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

mode = "dev"
host = "0.0.0.0"

if __name__ == '__main__':
    if mode == "dev":
        app.run(host=host,port=5100, debug=True)
    else:
        print("Flask API is running at "+host)
        serve(app,host=host,port=5100)
    # app.run(debug=True)