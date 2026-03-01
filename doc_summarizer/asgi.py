"""
ASGI config for doc_summarizer project.
"""

import os

from django.core.asgi import get_asgi_application

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'doc_summarizer.settings')

application = get_asgi_application()
