FROM public.ecr.aws/lambda/python:3.11

RUN pip3 install PyMuPDF boto3

# Copy function code

COPY app.py ./

CMD [ "app.handler" ]