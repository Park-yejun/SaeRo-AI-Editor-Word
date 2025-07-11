from flask import Flask
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/')
def hello():
    return "<h1>Cloud Run Deployment Successful!</h1>"