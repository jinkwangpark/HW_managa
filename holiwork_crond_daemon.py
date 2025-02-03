import os
import sys
import time
import subprocess
from pathlib import Path
from datetime import datetime


def daemonize():
    # 데몬 프로세스 생성
    try:
        pid = os.fork()
        if pid > 0:
            sys.exit(0)  # 부모 프로세스 종료
    except OSError as e:
        sys.stderr.write(f"Fork failed: {e}\n")
        sys.exit(1)

    # 새로운 세션 생성
    os.setsid()
    os.umask(0)
    os.chdir("/")

    # 표준 입출력 닫기
    sys.stdout.flush()
    sys.stderr.flush()
    sys.stdin.close()

    # 무한 루프로 주기적 작업 실행
    with open("/tmp/holiwork_daemon.log", "a") as log_file:
        while True:
            try:
                # 1. 로그 기록
                log_entry = f"[{datetime.now()}] Daemon is running...\n"
                log_file.write(log_entry)
                log_file.flush()

                # 2. 외부 스크립트 실행
                venv_python = (
                    Path(__file__).resolve().parent / ".venv" / "bin" / "python3"
                )
                script_path = Path(__file__).resolve().parent / "holiwork_manager.py"

                if not venv_python.exists():
                    raise FileNotFoundError(
                        f"Python 가상 환경 경로 없음: {venv_python}"
                    )
                if not script_path.exists():
                    raise FileNotFoundError(f"스크립트 경로 없음: {script_path}")

                result = subprocess.run(
                    [str(venv_python), str(script_path)],
                    capture_output=True,
                    text=True,
                    check=True,
                )

                # 3. 실행 결과 로깅
                log_file.write(f"[{datetime.now()}] Script output:\n{result.stdout}\n")

            except subprocess.CalledProcessError as e:
                error_msg = f"[{datetime.now()}] Script 실행 오류:\n{e.stderr}\nReturn Code: {e.returncode}\n"
                log_file.write(error_msg)
            except Exception as e:
                error_msg = f"[{datetime.now()}] Unexpected error: {str(e)}\n"
                log_file.write(error_msg)

            # 4. 30초 대기 후 재실행
            time.sleep(5)


if __name__ == "__main__":
    daemonize()
