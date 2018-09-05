from flask import Flask
import subprocess


app = Flask(__name__)

@app.route('/')
def hello_world():
    return 'Hello, World!'

@app.route('/test')
def test_sub_call():
    subprocess.run("excelSub.py")
    return "done"

@app.route('/hello')
def hello():
    return 'Hello, World'

app.run(host="localhost", port='8000')