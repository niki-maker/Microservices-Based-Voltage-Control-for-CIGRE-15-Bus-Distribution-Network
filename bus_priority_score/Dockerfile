# Use official Python base image
FROM python:3.10-slim

# Set working directory


# Copy your app code into the container
COPY bus_priority_score/bus_priority_score /app/bus_priority_score
COPY pyvolt /app/bus_priority_score/src/pyvolt

# Install dependencies
RUN pip install --no-cache-dir -r /app/bus_priority_score/requirements.txt
RUN pip install --no-cache-dir /app/bus_priority_score/src/pyvolt/.

# Expose the port your app runs on
EXPOSE 4003

# Run the Flask app
CMD ["python", "/app/bus_priority_score/bus_priority_score.py"]
