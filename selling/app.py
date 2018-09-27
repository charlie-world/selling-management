# -*- coding: utf-8 -*-

from flask import Flask, render_template, request
from . import __version__
from .parse.parse_app import save_data, init, make_html

def create_app():

    app = Flask(__name__)

    @app.route('/')
    def hello():
        return f"Selling management system (version: {__version__})"

    # render html
    @app.route('/upload')
    def render_upload():
        return render_template('render.html')

    @app.route('/upload-file', methods=['POST'])
    def uploadFile():
        f = request.files['file']
        file_name = f.filename
        f.save(f"./upload/{file_name}")

        data = save_data(f.filename)

        init()

        return make_html(data)

    return app
