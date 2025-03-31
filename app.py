from flask import Flask, render_template, send_from_directory
import subprocess
import os

app = Flask(__name__, static_folder="static", template_folder="templates")

@app.route('/')
def home():
    return render_template("index.html")  # Flask now correctly finds index.html

@app.route('/run-script')
def run_script():
    subprocess.Popen(["python", "start.py"], close_fds=True if os.name != 'nt' else False)
    return "Script started!"

if __name__ == '__main__':
    app.run(debug=True)
