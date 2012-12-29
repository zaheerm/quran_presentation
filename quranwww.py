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
        try:
            start_ayat = request.form['start']
            end_ayat = request.form['end']
        except KeyError:
            pass
        quran.create_presentation(sura, "/tmp/test.odp", start, end)
        bytes = None
        with open("/tmp/test.odp", "rb") as f:
            bytes = f.read()
        return (bytes, 200, {"Content-Type": "application/vnd.oasis.opendocument.presentation"})
    else:
        return render_template('quranform.html')

if __name__ == '__main__':
    app.run()