from flask import Flask, render_template

app = Flask(__name__)


@app.route('/')
def hello_world():
    return render_template('index.html', username = 'Elena')

@app.route('/<uname>')
def hello_world2(uname):
    return render_template('index.html', username = uname)


if __name__ == '__main__':
    app.run()
