from flask import Flask, request
import docx2txt
import re

app = Flask(__name__)

@app.route("/")
def info():
    return '''
    <body style="background-color:powderblue;">
        <center>
        <h1>Upload New File</h1>
        <form method="post" action="/uploadInfo" enctype="multipart/form-data">
        <br />
            <input type="file" name="files" multiple accept=".doc,.docx">
            <br /> 
            <br />
            <input type="submit">
            </center>
        </form>
    </body>
    '''
@app.route("/uploadInfo", methods=["POST"])
def uploadInfo():
    emails = ""
    for file in request.files.getlist("files"):
        try:
            myText = docx2txt.process(file)
            pattern = re.compile(r'[\w\.-]+@[a-z0-9\.-]+')
            matches = pattern.finditer(myText)
            for match in matches:
                emails += str(match.group(0)) + "<br />"
        except Exception as e:
            print(e)
    return '''
            <body style="background-color:powderblue;">
            <center>
                <p>{}</p>
            </center>
            </body>
            '''.format(emails)