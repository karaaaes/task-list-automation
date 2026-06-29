import os
from flask import Flask, render_template, request, Response, jsonify, stream_with_context
import requests

app = Flask(__name__)

BACKEND_URL = os.environ.get('BACKEND_URL', 'http://localhost:8888')


@app.route('/')
def home():
    return render_template('home.html', active='home')


@app.route('/tasks')
def tasks_page():
    return render_template('tasks.html', active='list_task')


@app.route('/tasks/create')
def create_task_page():
    return render_template('create_task.html', active='create_task')


@app.route('/tasks/generate-ppt')
def generate_ppt_page():
    return render_template('generate_ppt.html', active='generate_ppt')


# ---- Proxy all /api/* calls to backend ----
@app.route('/api/<path:subpath>', methods=['GET', 'POST', 'PUT', 'PATCH', 'DELETE'])
def proxy(subpath):
    url = f"{BACKEND_URL}/api/{subpath}"
    method = request.method

    # Pass through query string
    params = request.args.to_dict(flat=True)

    # Special handling for ppt download (binary streaming)
    headers = {k: v for k, v in request.headers if k.lower() not in ('host', 'content-length')}

    try:
        if method == 'GET':
            r = requests.get(url, params=params, headers=headers, stream=True, timeout=180)
        else:
            body = request.get_data()
            r = requests.request(method, url, params=params, headers=headers, data=body, stream=True, timeout=180)
    except requests.RequestException as e:
        return jsonify({'error': f'Backend unreachable: {e}'}), 502

    # Build response
    excluded = {'content-encoding', 'content-length', 'transfer-encoding', 'connection'}
    resp_headers = [(k, v) for k, v in r.raw.headers.items() if k.lower() not in excluded]
    return Response(stream_with_context(r.iter_content(chunk_size=8192)), status=r.status_code, headers=resp_headers)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8889, debug=False)
