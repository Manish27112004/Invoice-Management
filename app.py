# app.py

from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
import xlsxwriter
import io
import os
import pdfkit

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change this to a secure secret key

# Function to create a connection to the database
def get_db_connection():
    conn = sqlite3.connect('invoices.db')
    conn.row_factory = sqlite3.Row
    return conn

# Function to create the database tables
def create_tables():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS invoices
                      (id INTEGER PRIMARY KEY AUTOINCREMENT,
                       customer_name TEXT,
                       total_amount REAL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS invoice_items
                      (id INTEGER PRIMARY KEY AUTOINCREMENT,
                       invoice_id INTEGER,
                       item_name TEXT,
                       quantity INTEGER,
                       amount REAL)''')
    conn.commit()
    conn.close()

create_tables()

# Main home page with links to different functionalities
@app.route('/')
def home():
    if 'username' in session:
        return render_template('home.html')
    else:
        return redirect(url_for('login'))

# Route for login page
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''SELECT * FROM employees WHERE username = ? AND password = ?''', (username, password))
        employee = cursor.fetchone()
        conn.close()
        
        if employee:
            session['username'] = username
            return redirect(url_for('home'))
        else:
            return render_template('login.html', error='Invalid username or password')
    
    return render_template('login.html')

# Route to create a new invoice
@app.route('/create_invoice', methods=['GET', 'POST'])
def create_invoice():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        customer_name = request.form['customer_name']
        item_names = request.form.getlist('item_name')
        quantities = request.form.getlist('quantity')
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''INSERT INTO invoices (customer_name, total_amount)
                          VALUES (?, ?)''', (customer_name, 0))
        invoice_id = cursor.lastrowid
        
        total_amount = 0
        
        for item_name, quantity in zip(item_names, quantities):
            amount = get_item_price(item_name) * int(quantity)
            total_amount += amount
            cursor.execute('''INSERT INTO invoice_items (invoice_id, item_name, quantity, amount)
                              VALUES (?, ?, ?, ?)''', (invoice_id, item_name, quantity, amount))
        
        cursor.execute('''UPDATE invoices SET total_amount = ? WHERE id = ?''', (total_amount, invoice_id))
        
        conn.commit()
        conn.close()
        
        return redirect(url_for('view_invoice', invoice_id=invoice_id))
    
    return render_template('create_invoice.html')

# Route to view a specific invoice
@app.route('/view_invoice/<int:invoice_id>')
def view_invoice(invoice_id):
    if 'username' not in session:
        return redirect(url_for('login'))
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''SELECT * FROM invoices WHERE id = ?''', (invoice_id,))
    invoice = cursor.fetchone()
    
    if not invoice:
        return 'Invoice not found', 404
    
    cursor.execute('''SELECT * FROM invoice_items WHERE invoice_id = ?''', (invoice_id,))
    items = cursor.fetchall()
    
    conn.close()
    
    return render_template('view_invoice.html',invoice_id=invoice_id, invoice=invoice, items=items)

# Route to view all invoices
@app.route('/view_invoices')
def view_invoices():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''SELECT * FROM invoices''')
    invoices = cursor.fetchall()
    conn.close()
    
    return render_template('view_invoices.html', invoices=invoices)

# Main route to generate a report on all sales
@app.route('/generate_report')
def generate_report():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    # Generate Excel report
    workbook = xlsxwriter.Workbook('sales_report.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Write headers
    worksheet.write(0, 0, 'Invoice ID')
    worksheet.write(0, 1, 'Customer Name')
    worksheet.write(0, 2, 'Total Amount')
    
    # Fetch data from database and write to Excel
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''SELECT * FROM invoices''')
    row = 1
    for invoice in cursor.fetchall():
        worksheet.write(row, 0, invoice['id'])
        worksheet.write(row, 1, invoice['customer_name'])
        worksheet.write(row, 2, invoice['total_amount'])
        row += 1
    conn.close()
    
    workbook.close()
    
    # Send the generated Excel file as an attachment
    return send_file('sales_report.xlsx', as_attachment=True)

# Function to get the price of an item (replace this with your own item pricing logic)
def get_item_price(item_name):
    item_prices = {
        'Smoodh': 10,
        'Amul_Lassi': 25,
        'Lays': 10,
        'Snikers': 10,
        'Cold_Coffee': 40,
        'Bread': 40,
        'Protein_Bar': 50
    }
    return item_prices.get(item_name, 0)

# Main route to generate PDF invoice for a given invoice
@app.route('/generate_pdf_invoice/<int:invoice_id>')
def generate_pdf_invoice(invoice_id):
    if 'username' not in session:
        return redirect(url_for('login'))
    
    # Fetch invoice data from the database
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''SELECT * FROM invoices WHERE id = ?''', (invoice_id,))
    invoice = cursor.fetchone()
    
    if not invoice:
        conn.close()
        return 'Invoice not found', 404
    
    cursor.execute('''SELECT * FROM invoice_items WHERE invoice_id = ?''', (invoice_id,))
    items = cursor.fetchall()
    
    conn.close()
    
    # Render HTML template for the invoice
    rendered_html = render_template('invoice_template.html', invoice=invoice, items=items)
    
    # Path to the wkhtmltopdf executable
    path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'  # Modify this path as per your installation
    
    # Configure pdfkit to use wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    
    # Generate PDF from the rendered HTML
    pdf_filename = f'invoice_{invoice_id}.pdf'
    pdfkit.from_string(rendered_html, pdf_filename, configuration=config)
    
    # Send the generated PDF file as an attachment
    return send_file(pdf_filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)