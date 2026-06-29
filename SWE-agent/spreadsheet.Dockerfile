# docker build -f spreadsheet.Dockerfile -t spreadsheetbench-v2 .
FROM python:3.11.10-bullseye  

ARG DEBIAN_FRONTEND=noninteractive  
ENV TZ=Etc/UTC

WORKDIR /

# Install swe-rex for faster startup
RUN pip install pipx
RUN pipx install swe-rex
RUN pipx ensurepath
ENV PATH="$PATH:/root/.local/bin/"

# Install any extra dependencies
RUN pip install flake8
RUN pip install openpyxl numpy pandas matplotlib xlsxwriter
RUN sed -i 's|deb.debian.org|deb.debian.org|g' /etc/apt/sources.list
RUN apt-get update
RUN apt-get install -y libreoffice libreoffice-script-provider-python

# Add LibreOffice UNO library to Python path
ENV PYTHONPATH="/usr/lib/libreoffice/program:/usr/lib/python3/dist-packages:${PYTHONPATH}"

SHELL ["/bin/bash", "-c"]