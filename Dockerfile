FROM python:3.9-slim-buster 
WORKDIR /var/lib/jenkins/workspace/python_daac/spdocker
COPY . . 
RUN pip install -r requirements.txt
CMD ["python" , "sp_test.py"]
