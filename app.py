import time
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, abort
import pandas as pd
from io import BytesIO
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Spacer
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, ListFlowable, ListItem
from reportlab.lib.units import mm
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MERGED_FOLDER'] = 'merged'
app.secret_key = 'supersecretkey'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['MERGED_FOLDER'], exist_ok=True)


counter = 1


def insert_column(merged_df, position, base_name='id'):
    global counter

    base_name2 = 'name_x'
    base_name3 = 'name_y'
    new_column_name = f"{base_name}{counter}"
    new_column_name2 = f"{base_name2}{counter}"
    new_column_name3 = f"{base_name3}{counter}"

    merged_df.insert(position, new_column_name3, merged_df['name_y'])
    merged_df.insert(position, new_column_name2, merged_df['name_x'])
    merged_df.insert(position, new_column_name, merged_df['user_id'])


    counter += 1



def add_page_number(canvas, doc):
    page_num = canvas.getPageNumber()
    text = f"Page {page_num}"
    canvas.setFont("Helvetica", 9)
    width, height = letter
    canvas.drawString(650, 50, "IT Security team")
    canvas.drawCentredString(width / 2+100, 20, text)  # Position at bottom center of page


def merge_csv(users_file, details_file):
    try:
        users_df = pd.read_excel(users_file,dtype=str)
        details_df = pd.read_excel(details_file,dtype=str)
        merged_df = pd.merge(users_df, details_df, on='user_id', how='inner')
        print(merged_df)




        merged_df = merged_df.sort_index()
        return merged_df
    except Exception as e:
        print(f"Error merging CSV files: {e}")
        return None


#def get_timestamped_filename(original_filename):
#    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
#    name, ext = os.path.splitext(original_filename)
#    return f"{name}_{timestamp}{ext}"

def get_timestamped_filename(filename):
    safe_filename = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)  # 只允许字母、数字、下划线、破折号和点
    return f"{int(time.time())}_{safe_filename}"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'csv', 'xlsm', 'pdf'}



def save_merged_file(merged_df, file_type, title, title_2, note1):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    merged_filename = f"merged_data_{timestamp}.{file_type}"
    merged_file_path = os.path.join(app.config['MERGED_FOLDER'], merged_filename)

    if file_type == 'csv':
        merged_df.to_csv(merged_file_path, index=False)
    elif file_type == 'xlsx':
        with pd.ExcelWriter(merged_file_path, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='MergedData')
    elif file_type == 'pdf':
        if merged_df.empty:
            print("The merged DataFrame is empty.")
            return None  # Early return if there's no data

        try:
            # PDF Generation
            columns = merged_df.columns.tolist()

            new_columns = []
            modified_columns = []
            for i, col in enumerate(merged_df.columns):
                if 'Unnamed' in col and i > 0:  # 确保不是第一列
                    new_columns.append(new_columns[i - 1])  # 替换为左侧相邻列名
                    modified_columns.append(i + 1)  # 记录修改列的位置（从1开始计数）
                else:
                    new_columns.append(col)


            merged_df.columns = new_columns

            last_in_intervals = []
            for i in range(1, len(modified_columns)):

                if modified_columns[i] != modified_columns[i - 1] + 1:
                    last_in_intervals.append(modified_columns[i - 1])

            last_in_intervals.append(modified_columns[-1])

            print(f"test:{last_in_intervals}")


            print("hear")
            print(merged_df)


            new_header = merged_df.columns.tolist()
            merged_df.loc[-1] = new_header
            merged_df.index = merged_df.index + 1

            merged_df = merged_df.sort_index()

            modified_columns = []


            doc = SimpleDocTemplate(merged_file_path, pagesize=landscape(letter),leftMargin=0*inch, rightMargin=0*inch,
            topMargin=0*inch,bottomMargin=0*inch)
            story = []
            styles = getSampleStyleSheet()


            logo_path = 'static/images/logo.jpg'
            logo = Image(logo_path, width=2 * inch, height=1 * inch)


            text_line1 = ""
            text_line2 = "Confidential"


            text_paragraph1 = Paragraph(text_line1, styles['Normal'])
            text_paragraph2 = Paragraph(text_line2, styles['Normal'])


            spacer = Spacer(0, 0.35 * inch)


            text_block = [text_paragraph1, spacer, text_paragraph2]


            table_data = [[logo, text_block]]


            table1 = Table(table_data, colWidths=[8 * inch, None])
            table1.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 1 * mm),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0 * mm),
                #('LEFTPADDING', (0, 0), (0, 0), 0),
                #('RIGHTPADDING', (1, 0), (1, 0), 0),
            ]))


            #story.append(table1)



            print(merged_df)
            merged_df['user_id'] = merged_df['user_id']
            #merged_df.insert(8, 'id1', merged_df['user_id'])
            #merged_df.insert(16, 'id2', merged_df['user_id'])
            #insert_column(merged_df, 8)
            #insert_column(merged_df, 16)
            #merged_df.columns = merged_df.columns.str.replace('id1', 'user_id')
            #merged_df.columns = merged_df.columns.str.replace('id2', 'user_id')
            print(merged_df)


            # Add some space below the title
            story.append(Spacer(1, 12))

            # Convert dataframe to list of lists
            data = [merged_df.columns.tolist()] + merged_df.values.tolist()





            # Define margins and usable width
            left_margin = 0.5 * inch
            right_margin = 0.5 * inch
            usable_width = landscape(letter)[0] - left_margin - right_margin  # Total width minus margins

            # Calculate dynamic column widths based on content length
            col_widths = []
            for col in merged_df.columns:
                max_len = max(merged_df[col].astype(str).apply(len).max(), len(str(col)))
                len_num = max(1.3 * inch, min(1.0 * inch, max_len * 0.10 * inch))
                col_widths.append(len_num)
            print(col_widths)
            # Check if col_widths is empty
            if not col_widths:
                print("Column widths are empty.")
                return None  # Early return if column widths are not determined

            total_col_width = 0
            max_cols_per_page = 0
            max11 = []
            for width in col_widths:
                total_col_width += width
                if total_col_width <= usable_width:
                    max_cols_per_page += 1
                else:
                    total_col_width -= width
                    max11.append(max_cols_per_page)
                    print(total_col_width)
                    total_col_width = width
                    max_cols_per_page = 1
            max11.append(max_cols_per_page)
            print(usable_width)
            print(max11)

            # Calculate max columns per page based on usable width
            total_col_width = 0
            max_cols_per_page = 0
            max1 = []
            p = 0
            print(col_widths[0])

            for width in col_widths:
                total_col_width += width
                if total_col_width <= usable_width:
                    max_cols_per_page += 1
                    p = p + 1
                else:
                    max1.append(max_cols_per_page)
                    print(p)
                    insert_column(merged_df, p)
                    p = p + 4
                    total_col_width = width + col_widths[0] + col_widths[1] + col_widths[2]
                    max_cols_per_page = 4
            max1.append(max_cols_per_page)

            print("123344567890087655")
            print(merged_df)
            print("------")
            print(max1)

            print(merged_df)
            data = [merged_df.columns.tolist()] + merged_df.values.tolist()
            print(data)

            col_widths = []
            for col in merged_df.columns:
                max_len = max(merged_df[col].astype(str).apply(len).max(), len(str(col)))
                #len_num = max(1.0 * inch, min(2.0 * inch, max_len * 0.10 * inch))
                len_num = max(0.3 * inch, min(2.0 * inch, max_len * 0.10 * inch))
                if len_num > 1.5 * inch:
                    len_num = 1.5 * inch
                col_widths.append(len_num)
            print(col_widths)
            # Check if col_widths is empty
            if not col_widths:
                print("Column widths are empty.")
                return None  # Early return if column widths are not determined

            # Ensure we have columns to display
            if max_cols_per_page == 0:
                print("No columns fit on the page.")
                return None  # Early return if no columns can be displayed

            # Define row handling logic as before
            max_rows_per_page = 15
            rows = len(data)
            cols = len(data[0])
            num_row_pages = (rows // max_rows_per_page) + 1

            num_col_pages = len(max1)
            print(num_col_pages)
            print(len(max1))



            print("++++++++++++++++")
            for row_page in range(num_row_pages):
                start_row = row_page * max_rows_per_page
                if(start_row == 0):
                    start_row = 1
                end_row = min(start_row + max_rows_per_page, rows)
                end_col = 0
                for col_page in range(num_col_pages):
                    #start_col = col_page * max1[col_page]
                    start_col = end_col
                    print("----")
                    print(col_page)
                    print(start_col)
                    end_col = end_col + max1[col_page]
                    print(end_col)


                    # Prepare page data and create the table
                    page_data = [row[start_col:end_col] for row in data[start_row+1:end_row]]


                    header_row = data[1][start_col:end_col]
                    page_data.insert(0, header_row)
                    if start_row!=1:
                        page_data.insert(1, data[2][start_col:end_col])

                    if not page_data or not page_data[0]:  # Check if page_data is empty
                        print("Page data is empty.")
                        continue
                    story.append(table1)
                    titles = Paragraph(title, styles['Title'])
                    story.append(titles)

                    titles_2 =Paragraph(title_2, styles['Title'])
                    story.append(titles_2)

                    styleN = styles["BodyText"]
                    wrapped_data = []
                    for row in page_data:
                        wrapped_row = []
                        for item in row:
                            paragraph = Paragraph(str(item), styleN)
                            wrapped_row.append(paragraph)
                        wrapped_data.append(wrapped_row)
                    table = Table(wrapped_data, colWidths=col_widths[start_col:end_col])
                    print("Table created.")
                    print("================================")
                    print(max1[col_page])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        #('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        #('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('TOPPADDING', (0, 0), (-1, -1), 1 * mm),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 0 * mm),
                        ('SPAN', (0, 0), (0, 1)),
                        ('SPAN', (1, 0), (1, 1)),
                        ('SPAN', (2, 0), (2, 1)),
                        ('SPAN', (3, 0), (max1[col_page]-1, 0)),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ]))
                    story.append(table)

                    # Add page break if not the last row page
                    if row_page < num_row_pages - 1 or col_page < num_col_pages - 1:
                        story.append(PageBreak())


            # Create a spacer to position the footer content
            story.append(Spacer(1, 24))


            footer_content = [
                [Paragraph("Confirmed by (BM/HoD):", styles['Normal'])],
            ]

            #custom_style = ParagraphStyle(name='CustomStyle', fontSize=12, wordWrap='CJK', alignment=1)
            left_content = [
                Paragraph("Content: This is the left section."),
                Spacer(0,12),
                Paragraph("Signature: ________________"),
                Paragraph("Date: October 27, 2024"),
            ]


            #user_note = sanitize_input(user_note)


            right_content = [
                Paragraph("Note: This is the right section."),
                Paragraph("Name: ______________"),
                Paragraph(note1),
                #Paragraph(user_note, custom_style),
            ]


            combined_data = [[
                Table([[line] for line in left_content], colWidths=[3 * inch]),
                Table([[line] for line in right_content], colWidths=[3 * inch]),
            ]]

            combined_table = Table(combined_data, colWidths=[3.5 * inch, 3.5 * inch])
            combined_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))

            story.append(combined_table)

            # Add page number
            doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        except Exception as e:
            print(f"Error generating PDF: {e}")

    return merged_file_path


@app.before_request
def restrict_url():

    path = request.path

    if path != '/' and not path.startswith('/static/') or request.query_string:

        abort(404)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':

        title = request.form.get('title', 'merged_data')  # Default title if none provided
        title_2 = request.form.get('title2', 'merged_data')  # Default title if none provided
        note1 = request.form.get('note1', 'merged_data')  # Default title if none provided


        if 'users_file' in request.files and 'details_file' in request.files:
            users_file = request.files['users_file']
            details_file = request.files['details_file']

            if users_file and allowed_file(users_file.filename) and details_file and allowed_file(details_file.filename):
                users_file_name = get_timestamped_filename(users_file.filename)
                details_file_name = get_timestamped_filename(details_file.filename)

                users_file_path = os.path.join(app.config['UPLOAD_FOLDER'], users_file_name)
                details_file_path = os.path.join(app.config['UPLOAD_FOLDER'], details_file_name)

                users_file.save(users_file_path)
                details_file.save(details_file_path)

                merged_df = merge_csv(users_file_path, details_file_path)
                if merged_df is None:
                    flash("Error merging CSV files.", "error")
                    return redirect(url_for('index'))

                if 'download_csv' in request.form:
                    buffer = BytesIO()
                    merged_df.to_csv(buffer, index=False)
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.csv', mimetype='text/csv')

                elif 'download_excel' in request.form:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='MergedData')
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.xlsx',
                                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                elif 'download_pdf' in request.form:
                    merged_file_path = save_merged_file(merged_df, 'pdf', title, title_2, note1)
                    buffer = BytesIO()
                    with open(merged_file_path, 'rb') as f:
                        buffer.write(f.read())
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.pdf',
                                     mimetype='application/pdf')

                elif 'show_data' in request.form:
                    return render_template('index.html', tables=[merged_df.to_html(classes='data')],
                                           titles=merged_df.columns.values)

        return redirect(url_for('index'))

    uploaded_files = os.listdir(app.config['UPLOAD_FOLDER'])
    merged_files = os.listdir(app.config['MERGED_FOLDER'])

    return render_template('index.html', uploaded_files=uploaded_files, merged_files=merged_files)

@app.route('/download/<folder>/<filename>', methods=['GET'])
def download_file(folder, filename):
    folder_path = app.config.get(f'{folder.upper()}_FOLDER')
    file_path = os.path.join(folder_path, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash(f"The file {filename} does not exist.", "error")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(port= 5000, debug=True)
