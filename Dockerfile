# This packages the app together with everything it needs to run
# (Python + Streamlit + pandas + openpyxl) into one self-contained image.
# End users only need a web browser - nothing to install on their side.

FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (speeds up rebuilds when only app code changes)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Streamlit's default port
EXPOSE 8501

# Basic container health check
HEALTHCHECK CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')" || exit 1

ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
