from django.shortcuts import render
from django.http import HttpResponse
from .forms import InstitutionForm, StudentForm
from .models import Student


def generate_docx(request):
    if request.method == "POST":

        institution_form = InstitutionForm(request.POST)
        student_form = StudentForm(request.POST)

        if institution_form.is_valid() and student_form.is_valid():
            # Process the forms and generate the docx file
            institution_data = institution_form.cleaned_data
            student_data = student_form.cleaned_data

            # Create a new docx document
            from docx import Document

            doc = Document()

            # Add institution details to the document
            doc.add_heading("Institution Details", level=1)
            doc.add_paragraph(f'Name of the Institution: {institution_data["name"]}')
            doc.add_paragraph(f'Place: {institution_data["place"]}')
            doc.add_paragraph(f'District: {institution_data["district"]}')
            doc.add_paragraph(f'Phone number: {institution_data["phone_no"]}')
            doc.add_paragraph(f'Email Id: {institution_data["email"]}')

            # Add student details to a table in the document
            doc.add_heading("Student Details", level=1)
            table = doc.add_table(rows=1, cols=3)
            table.style = "TableGrid"
            table.cell(0, 0).text = "STUDENT NAME"
            table.cell(0, 1).text = "CLASS"
            table.cell(0, 2).text = "IFSC CODE"

            # Add student data to the table
            row_cells = table.add_row().cells
            row_cells[0].text = student_data['student_name']
            row_cells[1].text = student_data['student_class']
            row_cells[2].text = student_data['student_ifsc']

            # Save the document with the institution name as the filename
            filename = institution_data["name"].replace(" ", "_") + ".docx"
            doc.save(filename)

            # Prepare the response to trigger download
            with open(filename, "rb") as file:
                response = HttpResponse(
                    file.read(),
                    content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                response["Content-Disposition"] = f"attachment; filename={filename}"
            return response
    else:
        institution_form = InstitutionForm()
        student_form = StudentForm()

    return render(
        request,
        "generate_docx.html",
        {"institution_form": institution_form, "student_form": student_form},
    )
