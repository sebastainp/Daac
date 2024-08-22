FROM python:3.9-slim-buster 
WORKDIR /tmp
COPY . . 
RUN pip install -r requirements.txt
CMD ["python" , "sp_test.py"]
