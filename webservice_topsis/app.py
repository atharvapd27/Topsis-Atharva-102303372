import os
import smtplib
import io
import csv
import math
from flask import Flask, render_template, request, flash
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import openpyxl # Lightweight Excel reader

app = Flask(__name__)
app.secret_key = "supersecretkey"

# --- CREDENTIALS ---
# Set these in Vercel Environment Variables for security
SENDER_EMAIL = os.environ.get("atharvapandey100@gmail.com")
SENDER_PASSWORD = os.environ.get("azlp blsr czrx lhrp")

def calculate_topsis_lite(file_obj, filename, weights, impacts):
    try:
        data_rows = []
        headers = []
        file_extension = os.path.splitext(filename)[1].lower()

        # 1. READ DATA (Manual parsing to avoid Pandas)
        if file_extension == '.csv':
            # Decode bytes to string
            stream = io.StringIO(file_obj.read().decode("latin1"), newline=None)
            csv_reader = csv.reader(stream)
            headers = next(csv_reader)
            for row in csv_reader:
                data_rows.append(row)
                
        elif file_extension == '.xlsx':
            wb = openpyxl.load_workbook(file_obj, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            headers = list(rows[0])
            data_rows = [list(r) for r in rows[1:] if r[0] is not None] # Skip empty rows
            
        else:
            return None, "Error: Unsupported file format."

        if len(headers) < 3:
            return None, "Error: Input file must contain at least 3 columns."

        # 2. SEPARATE ID COLUMN AND DATA
        # Assuming Col 0 is Name/ID, Cols 1+ are numeric criteria
        names = [row[0] for row in data_rows]
        
        # Convert numeric data to floats
        matrix = []
        try:
            for row in data_rows:
                numeric_row = [float(x) for x in row[1:]]
                matrix.append(numeric_row)
        except ValueError:
            return None, "Error: Columns 2 onwards must contain numeric values only."

        num_rows = len(matrix)
        num_cols = len(matrix[0])

        if num_cols != len(weights):
            return None, f"Error: File has {num_cols} numeric columns, but {len(weights)} weights provided."

        # 3. TOPSIS ALGORITHM (Pure Math)
        
        # Step A: Normalize Matrix
        # Calculate denominator (sqrt of sum of squares) for each column
        denominators = []
        for j in range(num_cols):
            sum_squares = sum(matrix[i][j] ** 2 for i in range(num_rows))
            denominators.append(math.sqrt(sum_squares))

        # Step B: Calculate Weighted Normalized Matrix
        normalized_matrix = []
        for i in range(num_rows):
            norm_row = []
            for j in range(num_cols):
                val = (matrix[i][j] / denominators[j]) * weights[j]
                norm_row.append(val)
            normalized_matrix.append(norm_row)

        # Step C: Determine Ideal Best and Worst
        ideal_best = []
        ideal_worst = []
        
        for j in range(num_cols):
            col_values = [normalized_matrix[i][j] for i in range(num_rows)]
            if impacts[j] == '+':
                ideal_best.append(max(col_values))
                ideal_worst.append(min(col_values))
            else: # impact is '-'
                ideal_best.append(min(col_values))
                ideal_worst.append(max(col_values))

        # Step D: Euclidean Distances
        scores = []
        for i in range(num_rows):
            dist_best_sq = sum((normalized_matrix[i][j] - ideal_best[j]) ** 2 for j in range(num_cols))
            dist_worst_sq = sum((normalized_matrix[i][j] - ideal_worst[j]) ** 2 for j in range(num_cols))
            
            dist_best = math.sqrt(dist_best_sq)
            dist_worst = math.sqrt(dist_worst_sq)
            
            # Step E: Topsis Score
            if (dist_best + dist_worst) == 0:
                score = 0
            else:
                score = dist_worst / (dist_best + dist_worst)
            scores.append(score)

        # 4. RANKING
        # Combine names, original data, and scores
        final_results = []
        for i in range(num_rows):
            # Create a dictionary or list for the row
            row_data = data_rows[i] # Original row
            row_data.append(scores[i]) # Add Score
            final_results.append(row_data)

        # Sort by score (Descending) to determine rank
        # We store index to keep track of original order if needed, but here we just sort
        final_results.sort(key=lambda x: x[-1], reverse=True)

        # Assign Ranks
        for rank, row in enumerate(final_results, 1):
            row.append(rank) # Add Rank Column
            
        # Update Headers
        headers.append("Topsis Score")
        headers.append("Rank")

        return headers, final_results, None

    except Exception as e:
        return None, None, f"Calculation Error: {str(e)}"

def send_email(receiver_email, result_filename, headers, data_rows):
    msg = MIMEMultipart('alternative')
    msg['From'] = SENDER_EMAIL
    msg['To'] = receiver_email
    msg['Subject'] = "Topsis Results (Table & File)"

    # --- GENERATE HTML TABLE MANUALLY ---
    table_html = '<table style="border-collapse: collapse; width: 100%; border: 2px solid #333; font-family: Arial;">'
    
    # Header Row
    table_html += '<thead><tr style="background-color: #009879; color: white;">'
    for h in headers:
        table_html += f'<th style="border: 1px solid #000; padding: 8px;">{h}</th>'
    table_html += '</tr></thead><tbody>'

    # Data Rows
    for idx, row in enumerate(data_rows):
        bg_color = "#f2f2f2" if idx % 2 == 0 else "#ffffff"
        table_html += f'<tr style="background-color: {bg_color};">'
        for cell in row:
            # Format floats to 4 decimals for cleaner look
            val = f"{cell:.4f}" if isinstance(cell, float) else cell
            table_html += f'<td style="border: 1px solid #000; padding: 8px; text-align: center;">{val}</td>'
        table_html += '</tr>'
    table_html += '</tbody></table>'

    html_content = f"""
    <html>
      <body>
        <h2 style="color: #2e7d32; font-family: Arial;">Here are your TOPSIS Results</h2>
        <p style="font-family: Arial;">Please find the result file attached.</p>
        <br>
        {table_html}
        <br>
        <p style="font-family: Arial; font-size: 12px; color: #777;">Generated by Topsis Webservice</p>
      </body>
    </html>
    """
    msg.attach(MIMEText(html_content, 'html'))

    # --- GENERATE CSV IN MEMORY ---
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(headers)
    writer.writerows(data_rows)
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(csv_buffer.getvalue().encode('utf-8'))
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {result_filename}")
    msg.attach(part)

    # SMTP Connection
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())
    server.quit()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash("No file part")
            return render_template('index.html')
            
        file = request.files['file']
        weights_str = request.form['weights']
        impacts_str = request.form['impacts']
        email_id = request.form['email']

        if file.filename == '':
            flash("No selected file")
            return render_template('index.html')

        try:
            weights = [float(w) for w in weights_str.split(',')]
            impacts = [i.strip() for i in impacts_str.split(',')]
        except ValueError:
            flash("Invalid format for weights/impacts.")
            return render_template('index.html')

        if len(weights) != len(impacts):
            flash("Counts mismatch: Weights vs Impacts.")
            return render_template('index.html')
        
        # Calculate using Lite Version
        headers, result_rows, error = calculate_topsis_lite(file, file.filename, weights, impacts)
        
        if error:
            flash(error)
            return render_template('index.html')

        result_filename = f"result_{os.path.splitext(file.filename)[0]}.csv"

        try:
            send_email(email_id, result_filename, headers, result_rows)
            flash(f"Success! Check {email_id} for results.")
        except Exception as e:
            flash(f"Error sending email: {e}")

    return render_template('index.html')

# Vercel entry point
app = app