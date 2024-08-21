FROM python:3.8-slim-buster 
WORKDIR /tmp
COPY . . 
RUN pip install -r requirements.txt
CMD ["python3" , "sp_test.py"]
