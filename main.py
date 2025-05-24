from flask import Flask
import threading
import p9

app = Flask(__name__)


@app.route("/")
def home():
    return "Hello from PythonAnywhere!"

def delayed_call():
    import time
    print('akshat')
    time.sleep(5)
    p9.call()

@app.route('/run')
def run_script():
    thread=threading.Thread(target=delayed_call)
    thread.start()
    #p9.call()
    return "Akshat's Job Started"

app.run(host="0.0.0.0",port=81)
