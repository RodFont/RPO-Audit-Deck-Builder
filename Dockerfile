FROM python:3.12-slim

WORKDIR /srv/app

# System deps (python-pptx/openpyxl are pure-python + lxml wheels, no extra libs needed)
COPY app/requirements.txt ./app/requirements.txt
RUN pip install --no-cache-dir -r app/requirements.txt

# Copy the existing CLI script unmodified, plus the API wrapper.
COPY audit_to_deck.py ./audit_to_deck.py
COPY app ./app

RUN useradd --create-home --uid 1000 appuser \
    && chown -R appuser:appuser /srv/app
USER appuser

EXPOSE 8000
ENV MAX_UPLOAD_BYTES=26214400

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
