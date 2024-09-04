# Use the official Python image as a base
FROM python:3.10-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements.txt file and install dependencies
COPY requirements.txt .

# Install any necessary dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code to the working directory
COPY . .

# Set environment variables (if needed)
# ENV EXAMPLE_ENV_VAR=example_value

# Run the main.py script when the container launches
CMD ["python", "create_connection.py"]
