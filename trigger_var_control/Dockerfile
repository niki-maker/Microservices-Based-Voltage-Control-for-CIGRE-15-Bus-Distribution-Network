# Use an official Python base image
FROM python:3.9-slim

# Copy the main module (nr_module) and pyvolt into the container
COPY trigger_var_control/trigger_var_control /app/trigger_var_control
COPY pyvolt /app/trigger_var_control/src/pyvolt

# Install dependencies (if any)
RUN pip install --no-cache-dir -r /app/trigger_var_control/requirements.txt
RUN pip install --no-cache-dir /app/trigger_var_control/src/pyvolt/.

EXPOSE 4004
# Set the default command to run the script
CMD ["python", "/app/trigger_var_control/trigger_var_control.py"]