#!/usr/bin/env/python3

import os
from flask import Flask, request, render_template, flash, redirect, send_file
from werkzeug.utils import secure_filename
from tcsa import make_workbook
from UploadManager import UploadManager, allowed_file

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UploadManager().get_uploadfolder()
app.secret_key = b'C\n\xe4\x14?\xe5\x94\xa0\x7f\x17\x19Ej\x02l\x9e'

@app.route("/", methods=["GET"])
def root():
    return render_template('upload_form.html')

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        flash("no file present")
        return redirect('/')
    file = request.files['file']
    uploader = UploadManager(file)
    if file.filename == '':
        flash("no file selected")
        return redirect('/')
    if file and allowed_file(file.filename):
        # save file
        print('saving file....')
        file.save(uploader.get_fullpath())

        return send_file(
            make_workbook(uploader.get_fullpath()),
            mimetype="application/vnd.ms-excel"
        )
