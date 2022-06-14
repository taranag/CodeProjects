import datetime
import sqlite3
from flask import Flask, jsonify, render_template, request, send_file, url_for, flash, redirect
from werkzeug.exceptions import abort
from DownloadDataPPTX import *
from LearnDataPPTX import *

def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn


def get_post(post_id):
    conn = get_db_connection()
    post = conn.execute('SELECT * FROM posts WHERE id = ?',
                        (post_id,)).fetchone()
    conn.close()
    if post is None:
        abort(404)
    return post


app = Flask(__name__)
app.config['SECRET_KEY'] = b'|+\xb5E\xb7\x1b\xef\xbf~H\x14\x96\x81\xd8q)'
app.config['APPLICATION_ROOT'] = ''

prefix = app.config['APPLICATION_ROOT']
@app.route(prefix + '/')
def index():
    conn = get_db_connection()
    posts = conn.execute('SELECT * FROM posts').fetchall()
    conn.close()
    return render_template('index.html', posts=posts)

@app.route(prefix + '/<int:post_id>')
def post(post_id):
    post = get_post(post_id)
    return render_template('post.html', post=post)


@app.route('/create', methods=('GET', 'POST'))
def create():
    if request.method == 'POST':
        title = request.form['title']
        content = request.form['content']

        if not title:
            flash('Title is required!')
        else:
            conn = get_db_connection()
            conn.execute('INSERT INTO posts (title, content) VALUES (?, ?)',
                         (title, content))
            conn.commit()
            conn.close()
            return redirect(url_for('index'))

    return render_template('create.html')

@app.route('/generateReport', methods=('GET', 'POST'))
def generateReport():
    if request.method == 'POST':
        companyID = request.form['companyID']
        #type = request.form['type']
        type = request.form['options']
        groupBy = request.form['groupBy']
        
        todayDate = datetime.datetime.now()

        if not companyID or not type or not groupBy:
            flash('Please make sure all fields are filled out!')
        else:
            fileName = "Company" + companyID + type.capitalize() + "By" + groupBy.capitalize() + str(todayDate.date())
            url = ""
            if (type == "download"):
                url = generatePPTXDownloadData(companyID, fileName, groupBy)
            # elif (type == "value"):
            #     url = generatePPTXValueData(companyID, fileName, groupBy)
            elif (type == "learn"):
                 url = generatePPTXLearnData(companyID, fileName, groupBy)
            if (url != "" or url != None):
                return send_file(url)
            flash("Error: Could not generate report!")

    return render_template('generateReport.html')

@app.route('/<int:id>/edit', methods=('GET', 'POST'))
def edit(id):
    post = get_post(id)

    if request.method == 'POST':
        title = request.form['title']
        content = request.form['content']

        if not title:
            flash('Title is required!')
        else:
            conn = get_db_connection()
            conn.execute('UPDATE posts SET title = ?, content = ?'
                         ' WHERE id = ?',
                         (title, content, id))
            conn.commit()
            conn.close()
            return redirect(url_for('index'))

    return render_template('edit.html', post=post)

@app.route('/<int:id>/delete', methods=('POST',))
def delete(id):
    post = get_post(id)
    conn = get_db_connection()
    conn.execute('DELETE FROM posts WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash('"{}" was successfully deleted!'.format(post['title']))
    return redirect(url_for('index'))

@app.route('/get_number')
def get_number():
    value1 = request.args.get('val1')
    value2 = request.args.get('val2')
    value3 = int(value1) + int(value2)
    return jsonify({'data':f'The result is: {value3}'})

@app.route('/get_PPTX')
def get_PPTX():
    companyID = request.args.get('companyID')
    filename = request.args.get('filename')
    groupBy = request.args.get('groupBy')
    url = generatePPTXDownloadData(companyID, filename, groupBy)
    return jsonify({'url':url})

@app.route('/file/<fileName>', methods=['GET'])
def getFile(fileName):
    return send_file(fileName)

@app.route('/getPPTXdownload')
def getPPTXdownload():
    companyID = request.args.get('companyID')
    filename = request.args.get('filename')
    groupBy = request.args.get('groupBy')
    url = generatePPTXDownloadData(companyID, filename, groupBy)
    return send_file(url)