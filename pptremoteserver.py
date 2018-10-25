# This is a script to remote control PowerPoint presentations on Windows from your smartphone. 
# Usage:
#  * Run pptremoteserver.py on a server accessible from the internet
#  * Run pptremoteagent.py on the computer where you have PowerPoint running
#  * Open the pptremoteserver's IP address with your smartphone's browser and control the slideshow
#
# Notes:
#  * By default pptremoteagent.py polls localhost:5000. To have it poll a different address, run:
#       pptremoteagent.py -s <serverip>:<serverport>
#  * You will need to install the following Python 3 modules:
#       On server: flask
#       On agent: keyboard, pywin32, requests

import sys
from flask import Flask, jsonify, render_template

COMMAND     = None
SERVER_PORT = '8080' #modify this to set server to run on a different port

app = Flask(__name__)
app.config['SECRET_KEY'] = 'my-powerpoint-game-is-strong' 

@app.route("/", methods=['GET']) 

@app.route("/index/", methods=['GET'])
def index():
    global COMMAND
    
    return render_template('index.html')  

@app.route("/command/")
def command():
    global COMMAND
    
    returnvalue = jsonify({'command': COMMAND})
    COMMAND = None
    
    return (returnvalue)
    
@app.route("/setcmdnext/")
def setcmdnext():
    global COMMAND
    COMMAND = 'next'
    return jsonify({'command': COMMAND})
    
@app.route("/setcmdback/")
def setcmdback():
    global COMMAND
    COMMAND = 'back'
    return jsonify({'command': COMMAND})
    
def main(argv):
    app.run(host='0.0.0.0', port=SERVER_PORT)

if __name__ == '__main__':
    main(sys.argv[1:])  