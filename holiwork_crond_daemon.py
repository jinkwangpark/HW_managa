import os
import sys
import time
import subprocess
import logging
from pathlib import Path
from datetime import datetime, timedelta

# 로그 디렉토리 설정
LOG_DIR = Path("/tmp/holiwork_daemon_log")
LOG_DIR.mkdir(parents=True, exist_ok=True)

# 현재 날짜 기반 로그 파일 생성 (YYYY-MM)
current_month = datetime.now().strftime("%Y-%m")
log_file_path = LOG_DIR / f"holiwork_daemon_{current_month}.log"

# 로그 설정
logger = logging.getLogger("HoliworkDaemon")
handler = logging.FileHandler(log_file_path, encoding="utf-8")
formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)

# 6개월 지난 로그 삭제 함수
def clean_old_logs():
    six_months_ago = datetime.now() - timedelta(days=180)
    for log_file in LOG_DIR.glob("holiwork_daemon_*.log"):
        try:
            log_date_str = log_file.stem.split("_")[-1]  # YYYY-MM 추출
            log_date = datetime.strptime(log_date_str, "%Y-%m")
            if log_date < six_months_ago:
                os.remove(log_file)
                logger.info(f"Deleted old log file: {log_file}")
        except Exception as e:
            logger.error(f"Failed to delete log {log_file}: {e}")

# 데몬화 함수
def daemonize():
    try:
        pid = os.fork()
        if pid > 0:
            sys.exit(0)  # 부모 프로세스 종료
    except OSError as e:
        sys.stderr.write(f"Fork failed: {e}\n")
        sys.exit(1)

    os.setsid()
    os.umask(0)
    os.chdir("/")

    sys.stdout.flush()
    sys.stderr.flush()
    sys.stdin.close()

    while True:
        try:
            logger.info("Daemon is running...")

            # 외부 스크립트 실행
            venv_python = Path(__file__).resolve().parent / ".venv" / "bin" / "python3"
            script_path = Path(__file__).resolve().parent / "holiwork_manager.py"

            if not venv_python.exists():
                raise FileNotFoundError(f"Python 가상 환경 경로 없음: {venv_python}")
            if not script_path.exists():
                raise FileNotFoundError(f"스크립트 경로 없음: {script_path}")

            result = subprocess.run(
                [str(venv_python), str(script_path)],
                capture_output=True,
                text=True,
                check=True,
            )

            logger.info(f"Script output:\n{result.stdout}")

        except subprocess.CalledProcessError as e:
            logger.error(f"Script 실행 오류:\n{e.stderr}\nReturn Code: {e.returncode}")
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")

        # 6개월 지난 로그 정리 실행
        clean_old_logs()

        # 1시간(3600초) 대기 후 재실행
        time.sleep(3600)

if __name__ == "__main__":
    daemonize()
