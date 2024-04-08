import os
import re
from django.shortcuts import render
from django.http import HttpResponse
from .resume_extractor import process_resumes  # Corrected import statement

def home(request):
    if request.method == 'POST' and request.FILES.getlist('resume_files'):
        resume_files = request.FILES.getlist('resume_files')
        resume_dir = 'uploaded_resumes'
        os.makedirs(resume_dir, exist_ok=True)

        for resume_file in resume_files:
            with open(os.path.join(resume_dir, resume_file.name), 'wb') as f:
                for chunk in resume_file.chunks():
                    f.write(chunk)

        process_resumes(resume_dir)

        output_filename = os.path.join(resume_dir, "extracted_data.xlsx")
        with open(output_filename, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{os.path.basename(output_filename)}"'
            return response

    return render(request, 'home.html')
