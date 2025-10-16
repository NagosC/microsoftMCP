# Use an official Python runtime as a parent image
FROM python:3.12-slim-bullseye

# Set the working directory in the container
WORKDIR /app

# Install poetry
RUN pip install poetry

# Copy all the code
COPY . .

# Install dependencies and the project
RUN poetry install --no-interaction --no-ansi --only main

# Expose the port the app runs on
EXPOSE 8000

# Define the command to run the application
CMD ["poetry", "run", "microsoft-mcp"]