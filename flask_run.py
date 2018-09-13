from flask import Flask
import subprocess
import sys
from datetime import datetime
import json

app = Flask(__name__)

@app.route('/')
def hello_world():
    return 'Hello, World!'

@app.route('/test')
def test_sub_call():
    write_ts = str(datetime.now())
    data = {
        "operation": "request_fields",
        "write_ts": write_ts
    }
    with open("task.json", "w") as write_file:
        json.dump(data, write_file)

    subprocess.run([sys.executable, "excelSub.py"])
    return "done"

@app.route('/hello')
def hello():
    return 'Hello, World'

app.run(host="localhost", port='8000')