FROM python:3-slim-bullseye

WORKDIR /usr/src/app

# install dependency
RUN apt-get update &&\
    apt-get install -y curl gnupg1 g++

# install microsoft ODBC driver
RUN curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - &&\
    curl https://packages.microsoft.com/config/debian/11/prod.list\
      > /etc/apt/sources.list.d/mssql-release.list &&\
    apt-get update &&\
    ACCEPT_EULA=Y apt-get install -y msodbcsql17 mssql-tools &&\
    apt-get install -y unixodbc-dev libgssapi-krb5-2
ENV PATH="$PATH:/opt/mssql-tools/bin"

RUN poetry env use python3.9
RUN pip install poetry

COPY pyproject.toml /usr/src/app

RUN poetry install

COPY src/* /usr/src/app/

ENTRYPOINT poetry run python src/gen-report.py\
  --host         [database host server]\
  --db           [database name]\
  --username     [username]\
  --password     [password]\
  --query        [SQL Query]\
  --template     [template filename.docx]\
  --output       [filename without extension]\
  --pdf_password [password for the generated pdf file]
