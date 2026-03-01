import os
from django.shortcuts import render
from django.conf import settings
from .forms import UploadFileForm
from .utils import process_file_by_extension

def upload_file(request):
    context = {}
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            upload = request.FILES['file']
            name = upload.name
            ext = os.path.splitext(name)[1].lower()

            if ext not in ['.pdf', '.docx', '.ppt', '.pptx']:
                context['error'] = 'sirf PDF, DOCX ya PPT/PPTX file upload karo.'
            else:
                input_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
                os.makedirs(input_dir, exist_ok=True)
                input_path = os.path.join(input_dir, name)

                with open(input_path, 'wb+') as dest:
                    for chunk in upload.chunks():
                        dest.write(chunk)

                output_dir = os.path.join(settings.MEDIA_ROOT, 'processed')
                os.makedirs(output_dir, exist_ok=True)

                safe_name = name.replace(' ', '_')
                base = os.path.splitext(safe_name)[0]
                out_name = f"highlighted_{base}.pdf"
                output_path = os.path.join(output_dir, out_name)

                try:
                    summary, keywords = process_file_by_extension(
                        input_path, output_path
                    )
                    download_url = settings.MEDIA_URL + 'processed/' + out_name
                    context['download_url'] = download_url
                    # Summary is now ONLY in the PDF, not shown here
                    context['message'] = f'✅ Summary PDF ready! Total {len(keywords)} key points found.'
                except Exception as e:
                    context['error'] = f'processing me error: {e}'
    else:
        form = UploadFileForm()

    context['form'] = context.get('form', form)
    return render(request, 'core/upload.html', context)
