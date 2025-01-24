import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate,Table,TableStyle,Paragraph,Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import math

def read_excel(file_name):
    try:
        data = pd.read_excel(file_name)
        if not {'Student ID','Name','Subject','Score'}.issubset(data.columns):
            raise ValueError("Missing required columns in the Excel file.")
        return data
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

def process_data(data):
    try:
        group = data.groupby(['Student ID','Name'])
        total_score = group['Score'].sum().reset_index()
        average_score = group['Score'].mean().reset_index()

        summary = pd.merge(total_score,average_score,on=['Student ID','Name'])
        summary.columns = ['Student ID','Name','Total Score','Average Score']

        return summary,group
    except Exception as e:
        print(f"Error processing data: {e}")
        return None, None

def generate_pdf_report(student_id,name,scores,total,average):
    try:
        file_name = f"report_card_{student_id}.pdf"
        doc = SimpleDocTemplate(file_name,pagesize=letter,leftMargin=50,rightMargin=50,topMargin=0,bottomMargin=5)
        elements = []
        styles = getSampleStyleSheet()
        styles['Normal'].fontSize = 20
        styles['Normal'].spaceBefore = 20  
        styles['Normal'].spaceAfter = 20 
        styles['Title'].fontSize = 25
        image_path = "logo.png"
        img = Image(image_path,width=600,height=100) 
        elements.append(img)
        title = Paragraph("Report Card",styles['Title'])
        elements.append(title)
        about = Paragraph(f'Name:{name}<br/><br/>Student id :{student_id}',styles['Normal'])
        elements.append(about)
        summary = Paragraph(f"Total Score: {total}<br/><br/>Average Score: {average:.2f}",styles['Normal'])
        elements.append(summary)
        # Initializing table data with header
        table_data = [["Subject", "Score", "Status", "Grades"]]

        # Initializing fail counter
        fail_counter = 0

        # Looping through each score and calculate status and grade
        for score in scores:
            # Assigning status based on the score
            if math.isnan(score[1]):  # Check if score is NaN
                score[1] = "Missing Data or Absent"
                status = "N/A"
                grade = "N/A"
            else:
                status = "pass" if score[1] > 40 else "fail"
                
                # Incrementing fail_counter if the student failed
                if status == "fail":
                    fail_counter += 1

                # Assigning grade based on the score
                if score[1] > 85:
                    grade = "A"
                elif score[1] >= 70:
                    grade = "B"
                elif score[1] >= 50:
                    grade = "C"
                else:
                    grade = "D"
            # Append the row with Subject, Score, Status, and Grade
            table_data.append([score[0], score[1], status, grade])
        table = Table(table_data,hAlign='CENTER')
        table.setStyle(TableStyle([
            # Header row style
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),  # Dark blue background for the header row
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # White text for the header row
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center align all cells
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for header
            ('FONTSIZE', (0, 0), (-1, 0), 16),  # Font size for header
            ('PADDING', (0, 0), (-1, 0), 25),  # Padding for the header row
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),  # Extra bottom padding for the header
            
            # Body rows style
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),  # Light blue background for body rows
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),  # Black text for body rows
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),  # Center align all body cells
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Regular font for body rows
            ('FONTSIZE', (0, 1), (-1, -1), 14),  # Font size for body rows
            ('PADDING', (0, 1), (-1, -1), 8),  # Padding for body rows
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),  # Bottom padding for body rows
            
            # Alternating row background color
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),  # Alternating row background color
            
            # Grid style
            ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Thin grid line for all cells
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),  # Thicker line above the header row
            ('LINEBELOW', (0, -1), (-1, -1), 2, colors.black),  # Thicker line below the last row
        ]))
        elements.append(table)
        if fail_counter==0:
            result = Paragraph("Result : passed",styles['Normal'])
        else:
            result = Paragraph(f"fail in {fail_counter} subjects",styles['Normal'])
        elements.append(result)
        text = Paragraph(f"Wish you all the best {name}",styles['Title'])
        elements.append(text)
        doc.build(elements)
        print(f"Report card generated: {file_name}")
    except Exception as e:
        print(f"Error generating PDF for {name}: {e}")

def main():
    excel_file = "student_scores.xlsx"
    data = read_excel(excel_file)

    if data is not None:
        summary,group = process_data(data)

        if summary is not None:
            for _, row in summary.iterrows():
                student_id = row['Student ID']
                name = row['Name']
                total = row['Total Score']
                average = row['Average Score']

                # Get subject-wise scores for the student
                student_data = group.get_group((student_id, name))
                scores = student_data[['Subject', 'Score']].values.tolist()

                generate_pdf_report(student_id,name,scores,total,average)

if __name__ == "__main__":
    main()
