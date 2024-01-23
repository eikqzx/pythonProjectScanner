from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import win32com.client
from datetime import datetime
import pythoncom
import io
import base64
from waitress import serve

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/')
def index():
    return "Flask API is running"

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

        if scan_result:
            return jsonify({"message": "Scanning complete", "scannedImageBase64": scan_result})
        else:
            return jsonify({"error": "Error occurred during scanning."})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": "Error occurred during scanning."})


def perform_scan(scanner_id,dpi,pages):
    pythoncom.CoInitialize()  # Initialize COM

    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        dev = wia.ShowSelectDevice()
        if dev.DeviceID == scanner_id:
            dev.Items(1).Properties("Horizontal Resolution").Value = dpi
            dev.Items(1).Properties("Vertical Resolution").Value = dpi

            if dev.Properties("Document Handling Select").Value == 1:  # Automatic Document Feeder (ADF)
                dev.Items(1).Properties("Document Handling Select").Value = pages

            image_format = '{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}'  # WIA_FORMAT_JPEG
            image_data = dev.Items(1).Transfer(image_format).FileData.BinaryData

            return base64.b64encode(image_data).decode('utf-8')  # Return the base64-encoded image
        else:
            return None
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

@app.route('/api/documentHandlingOptions', methods=['GET'])
def get_document_handling_options_route():
    try:
        scanner_id = request.args.get('scannerId')
        options = get_document_handling_options(scanner_id)

        if options is not None:
            return jsonify({"documentHandlingOptions": options})
        else:
            return jsonify({"error": "Error occurred while fetching document handling options."})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": "Error occurred while fetching document handling options."})

def get_document_handling_options(scanner_id):
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        dev = wia.ShowSelectDevice()

        # Assuming 'scanner_id' is used to identify the selected scanner
        if dev.DeviceID == scanner_id:
            # Get available document handling options
            document_handling_options = []
            if dev.Properties.Exists("Document Handling Select"):
                for item in dev.Properties("Document Handling Select").Items:
                    document_handling_options.append(item.Value)

            return document_handling_options
        else:
            return None  # Return None if the selected scanner is not found
    except Exception as e:
        print(f"Error: {e}")
        return None

mode = "production"
host = "0.0.0.0"
if __name__ == '__main__':
    if mode == "dev":
        app.run(host=host,port=5100, debug=True)
    else:
        print("Flask API is running at "+host)
        serve(app,host=host,port=5100)
    # app.run(debug=True)