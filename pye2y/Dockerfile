ARG PYTHON_VER

FROM python:${PYTHON_VER}

USER root

RUN mkdir /src

WORKDIR /src

COPY src/requirements.txt /src/requirements.txt
RUN pip install -r requirements.txt

