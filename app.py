import os
from flask import Flask, render_template, make_response

app = Flask(__name__)

app.secret_key = 'hhtt_unified_system'

from hhtt_fujian import fujian_app
from hhtt_manage import manage_app

app.register_blueprint(fujian_app, url_prefix='/fujian')
app.register_blueprint(manage_app, url_prefix='/manage')

@app.route('/')
def index():
    response = make_response(render_template('index.html'))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)