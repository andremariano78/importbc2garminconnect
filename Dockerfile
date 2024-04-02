FROM python:3

WORKDIR /app

copy requirements.txt ./

RUN pip install --no-cache-dir --upgrade pip \
  && pip install --no-cache-dir -r requirements.txt

copy . .

CMD ["python", "app/ImportBodyComposition2GarminConnect.py"]