import os
import sys
import time
import subprocess  # 외부 스크립트 실행용
from pathlib import Path
from datetime import datetime


def daemonize():
    # 부모 프로세스를 종료하고 자식 프로세스 실행
    pid = os.fork()
    if pid > 0:
        sys.exit()  # 부모 프로세스 종료

    # 새로운 세션 시작
    os.setsid()

    # 다시 포크하여 완전히 독립된 프로세스 생성
    pid = os.fork()
    if pid > 0:
        sys.exit()

    # 작업 디렉토리 변경
    os.chdir("/")
    os.umask(0)

    # 표준 입출력 스트림 닫기
    sys.stdout.flush()
    sys.stderr.flush()
    sys.stdin.close()

    # 백그라운드 작업
    with open("/tmp/holiwork_daemon.log", "a") as log_file:
        while True:
            # 로그 기록
            log_file.write(f"[{datetime.now()}] Daemon is running...\n")
            log_file.flush()

            # 외부 스크립트 실행
            try:
                venv_python = f"{Path(__file__).resolve().parent}/.venv/bin/python3.9"
                script_path = f"{Path(__file__).resolve().parent}/holiwork_manager.py"
                print("venv_python : ", venv_python)
                print("script_path : ", script_path)

                result = subprocess.run(
                    [venv_python, script_path],
                    capture_output=True,
                    text=True,
                    check=True,
                )
                log_file.write(f"[{datetime.now()}] Script output:\n{result.stdout}\n")
            except subprocess.CalledProcessError as e:
                log_file.write(
                    f"[{datetime.now()}] Error while executing your_script.py:\n{e.stderr}\n"
                )
                log_file.write(f"Return code: {e.returncode}\n")

            time.sleep(5)


if __name__ == "__main__":
    daemonize()
