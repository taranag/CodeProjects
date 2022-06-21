import datetime
import sqlite3
from flask import Flask, jsonify, render_template, request, send_file, url_for, flash, redirect
from werkzeug.exceptions import abort
from DownloadDataPPTX import *
from LearnDataPPTX import *
from DataGenerator import generateFullReport, getCompanies

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

def get_report(report_id):
    conn = get_db_connection()
    report = conn.execute('SELECT * FROM reports WHERE id = ?',
                        (report_id,)).fetchone()
    conn.close()
    if report is None:
        abort(404)
    return report



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

@app.route('/reports')
def reports():
    conn = get_db_connection()
    reports = conn.execute('SELECT * FROM reports order by date desc').fetchall()
    conn.close()
    return render_template('reports.html', reports=reports)

@app.route(prefix + '/reports/<int:report_id>')
def report(report_id):
    report = get_report(report_id)
    return render_template('report.html', report=report)



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
        options = [0, 0, 0, 0]
        companyID = request.form['companyID']
        groupBy = request.form['groupBy']
        startDate = request.form['startDate']
        endDate = request.form['endDate']
        
        todayDate = datetime.datetime.now()

        optionsString = ''

        if request.form.get('titlePage'):
            options[0] = 1
            optionsString += 'Title Page, '
        if request.form.get('download'):
            options[1] = 1
            optionsString += 'Download, '
        if request.form.get('learn'):
            options[2] = 1
            optionsString += 'Learn, '
        if request.form.get('value'):
            options[3] = 1
            optionsString += 'Value, '

        optionsString = optionsString[:-2]
        
        print(options)
    
        if not endDate:
            endDate = todayDate.strftime("%Y-%m-%d")

        if not companyID or not groupBy or not startDate:
            flash('Please make sure all fields are filled out!')
        else:
            fileName = "Company" + companyID + "By" + groupBy.capitalize() + "From" + startDate + "To" + endDate
            url = ""
            url = generateFullReport(companyID, fileName, groupBy, startDate, endDate, options)
            if (url != "" or url != None):
                conn = get_db_connection()
                conn.execute('INSERT INTO reports (companyID, options, groupBy, startDate, endDate) VALUES (?, ?, ?, ?, ?)',
                            (companyID, optionsString, groupBy, startDate, endDate))
                conn.commit()
                conn.close()
                return send_file(url)
            flash("Error: Could not generate report!")

    companies = getCompanies()
    return render_template('generateReport.html', companies=companies)

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