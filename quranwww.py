import os
import tempfile

from flask import Flask
from flask import request
from flask import render_template

import quran

app = Flask(__name__)

@app.route("/", methods=['GET', 'POST'])
def get_quran():
    if request.method == 'POST':
        sura = request.form['sura']
        start_ayat = None
        end_ayat = None
        arabic_font = "Calibri"
        try:
            start_ayat = int(request.form['start'])
            end_ayat = int(request.form['end'])
            arabic_font = request.form["arabic_font"]
        except KeyError:
            pass
        _, filename = tempfile.mkstemp()
        quran.create_presentation(int(sura), filename, start_ayat, end_ayat, arabic_font)
        bytes = None
        with open(filename, "rb") as f:
            bytes = f.read()
        os.unlink(filename)
        return (bytes, 200, {"Content-Type": "application/vnd.oasis.opendocument.presentation"})
    else:
        return render_template('quranform.html')

if __name__ == '__main__':
    app.run(debug=True)