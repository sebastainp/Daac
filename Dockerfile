FROM python:3.8-slim-buster 
WORKDIR /tmp
COPY . . 

CMD ["python" , "sp_test.py"]
