# Use an official Python base image
FROM python:3.9-slim

# Copy the main module (nr_module) and pyvolt into the container
COPY main/main /app/main
COPY pyvolt /app/main/src/pyvolt

# Install dependencies (if any)
RUN pip install --no-cache-dir -r /app/main/requirements.txt
RUN pip install --no-cache-dir /app/main/src/pyvolt/.

#For Case1 (Undervoltage condition)
CMD ["python", "/app/main/Case1.py"]

#For Case2 (Overvoltage condition)
#CMD ["python", "/app/main/Case2.py"]

#For Case3 
#CMD ["python", "/app/main/Case3.py"]