# 공식 Python 3.11 이미지를 기반으로 합니다.
FROM python:3.11-slim

# 환경 변수를 설정합니다.
ENV APP_HOME /app
ENV PORT 8080
WORKDIR $APP_HOME

# requirements.txt를 복사하고 라이브러리를 설치합니다.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 애플리케이션 코드를 복사합니다.
COPY . .

# 컨테이너가 시작될 때 웹 서버(gunicorn)를 실행합니다.
# Cloud Run이 지정하는 포트($PORT)에서 요청을 받도록 설정합니다.
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 main:app
