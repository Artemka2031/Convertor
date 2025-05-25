FROM convertor-base:latest
WORKDIR /app
COPY . /app
EXPOSE 8000
CMD ["python", "-m", "api.server"]