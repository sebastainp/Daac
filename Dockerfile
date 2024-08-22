FROM python:3.8-slim-buster 
WORKDIR /tmp
COPY . . 
RUN pip install -r requirements.txt
CMD ["python" , "sp_test.py"]
