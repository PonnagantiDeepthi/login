from flask import Flask,render_template,request,redirect,url_for, Response
from flask_mysqldb import MySQL
from fpdf import FPDF
import io

import xlwt
import pymysql


app = Flask(__name__)


app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = 'Ajay_111'
app.config['MYSQL_DB'] ='flaskapp1'



mysql = MySQL()
mysql.init_app(app)

@app.route("/", methods=['GET','POST'])
def home():
    return render_template('index2.html')

@app.route("/details",methods=['GET','POST'])
def filldetails():
    if request.method=='POST':
        userdetails=request.form
        name=userdetails['name']
        email=userdetails['email']
        gender=userdetails['gender']
        pno=userdetails['pno']
        dob=userdetails['dob']
        ssc=userdetails['ssc']
        inter=userdetails['inter']
        grad=userdetails['grad']
        pyear=userdetails['pyear']
        ad=userdetails['ad']
        
        cur = mysql.connection.cursor()
        cur.execute("INSERT INTO users(name,email,gender,pno,dob,ssc,inter,grad,pyear,ad )VALUES(%s, %s,%s,%s,%s,%s,%s,%s,%s,%s)",(name,email,gender,pno,dob,ssc,inter,grad,pyear,ad
       
        ))
        mysql.connection.commit()
        cur.close()
        
    return render_template('details.html')





@app.route('/download/report/excel')
def download_excel():
    cur1 = mysql.connection.cursor()
    
    cur1.execute("SELECT * FROM users")
    result = cur1.fetchall()
    output = io.BytesIO()
    workbook = xlwt.Workbook()
    sh = workbook.add_sheet('EMPLOYEE REPORT')
    sh.write(0, 0, 'name')
    sh.write(0, 1, 'email')
    sh.write(0, 2, 'gender')
    sh.write(0, 3, 'pno')
    sh.write(0, 4, 'dob')
    sh.write(0, 5, 'ssc')
    sh.write(0, 6, 'inter')
    sh.write(0, 7, 'grad')
    sh.write(0, 8, 'pyear')
    sh.write(0, 9, 'ad')
    idx = 0
    for row in result:
        sh.write(idx+1, 0, row[0])
        sh.write(idx+1, 1, row[1])
        sh.write(idx+1, 2, row[2])
        sh.write(idx+1, 3, row[3])
        sh.write(idx+1, 4, row[4])
        sh.write(idx+1, 5, row[5])
        sh.write(idx+1, 6, row[6])
        sh.write(idx+1, 7, row[7])
        sh.write(idx+1, 8, row[8])
        sh.write(idx+1, 9, row[9])
        idx += 1
    workbook.save(output)
    output.seek(0)
    cur1.close()
        
    return Response(output, mimetype="application/ms-excel", headers={"Content-Disposition":"attachment;filename=employee.xls"})

@app.route('/column',methods=['GET','POST'])
def columns_pdf():
    if request.method=='POST':
        a=request.form.getlist('bike')
        print(a)
        cur2 = mysql.connection.cursor() 
        cur2.execute("SELECT * FROM users")
        result = cur2.fetchall()
        a=request.form.getlist('details')
        pdf = FPDF()
        pdf.add_page()
        page_width = pdf.w - 2 * pdf.l_margin
        pdf.set_font('Times','B',14.0) 
        pdf.cell(page_width, 0.0, 'User Data', align='C')
        pdf.ln(10)
        pdf.set_font('Times', 'B', 12)
        col_width = page_width/5
        pdf.ln(1)
        th = pdf.font_size
        if("name" in a):
            pdf.cell(col_width, th, 'Name', border=1)
        if("email" in a):
            pdf.cell(col_width, th, 'Email', border=1)
        if("gender" in a):
            pdf.cell(col_width, th, 'Gender' ,border=1)
        if("pno" in a):
            pdf.cell(col_width, th, 'Pno', border=1)
        if("dob" in a):
            pdf.cell(col_width, th, 'DOB' ,border=1)
        if("ssc" in a):
            pdf.cell(col_width, th, 'SSC' ,border=1)
        if("inter" in a):
            pdf.cell(col_width, th, 'Inter' ,border=1)
        if("grad" in a):
            pdf.cell(col_width, th, 'Graduation' ,border=1)
        if("pyear" in a):
            pdf.cell(col_width, th, 'Pass out year' ,border=1)
        if("ad" in a):
            pdf.cell(col_width, th, 'Address' ,border=1)
    
    
        

        pdf.ln(th)

        for row in result:
            if("name" in a):
                pdf.cell(col_width, th, row[0], border=1)
            if("email" in a):
                pdf.cell(col_width, th, row[1], border=1)
            if("gender" in a):
                pdf.cell(col_width, th, row[2], border=1)
            if("pno" in a):
                 pdf.cell(col_width, th, row[3], border=1)
            if("dob" in a):
                 pdf.cell(col_width, th, row[4], border=1)
            if("ssc" in a):
                pdf.cell(col_width, th, row[5], border=1)
            if("inter" in a):
                pdf.cell(col_width, th, row[6], border=1)
            if("grad" in a):
                pdf.cell(col_width, th, row[7], border=1)
            if("pyear" in a):
                pdf.cell(col_width, th, row[8], border=1)
            if("ad" in a):
                pdf.cell(col_width, th, row[9], border=1)
            pdf.ln(th)
        pdf.ln(10)
        pdf.set_font('Times','',10.0) 
        pdf.cell(page_width, 0.0, '- end of report -', align='C')
        cur2.close() 
     
        return Response(pdf.output(dest='S').encode('latin-1'), mimetype='application/pdf',headers={'Content-Disposition':'attachment;filename=employee_report.pdf'})       
    
    return render_template('pdfh.html')


if __name__ == "__main__":
    app.run(debug=True)