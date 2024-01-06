from flask import Flask, redirect, render_template, request
from argo_functions import *
from io import StringIO
from dotenv import load_dotenv # python-dotenv
# Load environment variables from .env file
load_dotenv()
from waitress import serve

app = Flask(__name__)

@app.route(f'/{os.getenv("SECRETPATH")}', methods=['GET', 'POST'])
def index():
    return render_template('index.html', secret_path=os.getenv("SECRETPATH"))

@app.route(f'/{os.getenv("SECRETPATH")}/<search_term>', methods=['GET'])
def search_pptx(search_term):
    output = StringIO()

    search_pptx_files(search_term, output_stream=output, http_mode=True, endpoint = os.getenv("ENDPOINT"))

    return output.getvalue()

if __name__ == '__main__':
    serve(app, listen='0.0.0.0:8080', url_scheme='http')