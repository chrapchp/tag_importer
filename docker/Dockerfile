FROM python:3.7.5
WORKDIR /usr/src/tag_importer
RUN apt-get update && apt-get install -y xclip
COPY requirements.txt ./
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt



