# type: ignore

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def read_excel(file_name):
    try:
        return pd.read_excel(file_name)
    except FileNotFoundError:
        print(f"Error: File '{file_name}' not found.")
        exit(1)
    except Exception as e:
        print(f"Error reading the file: {e}")
        exit(1)

def validate_data(df):
    required_columns = {'Student ID', 'Name', 'Subject', 'Score'}
    if not required_columns.issubset(df.columns):
        print(f"Error: Excel file must contain the columns: {required_columns}")
        exit(1)

    if df.isnull().values.any():
        print("Error: Excel file contains missing values.")
        exit(1)

def generate_pdf_report(student_id, name, subject_scores, total_score, avg_score):
    file_name = f"report_card_{student_id}.pdf"
    doc = SimpleDocTemplate(file_name, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(f"<b>Student Name:</b> {name}", styles['Normal']))
    elements.append(Paragraph(f"<b>Student ID:</b> {student_id}", styles['Normal']))
    elements.append(Paragraph(f"<b>Total Score:</b> {total_score}", styles['Normal']))
    elements.append(Paragraph(f"<b>Average Score:</b> {avg_score:.2f}", styles['Normal']))

    table_data = [["Subject", "Score"]] + [[subject, score] for subject, score in subject_scores.items()]
    table = Table(table_data, colWidths=[200, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)

    try:
        doc.build(elements)
        print(f"Report card generated: {file_name}")
    except Exception as e:
        print(f"Error generating PDF for {name}: {e}")

def main():
    file_name = "student_scores.xlsx"
    df = read_excel(file_name)

    validate_data(df)

    grouped = df.groupby(['Student ID', 'Name'])
    for (student_id, name), group in grouped:
        subject_scores = dict(zip(group['Subject'], group['Score']))
        total_score = group['Score'].sum()
        avg_score = group['Score'].mean()

        generate_pdf_report(student_id, name, subject_scores, total_score, avg_score)

if __name__ == "__main__":
    main()
