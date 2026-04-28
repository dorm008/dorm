import os
import psutil
import winapps
from msoffice_crypt import OfficeFile

class PCButler:
    # 1. 확장자 일괄 변경
    def change_extensions(self, folder_path, old_ext, new_ext):
        for filename in os.listdir(folder_path):
            if filename.endswith(old_ext):
                base = os.path.splitext(filename)[0]
                os.rename(
                    os.path.join(folder_path, filename),
                    os.path.join(folder_path, base + new_ext)
                )
        print(f"✅ {old_ext} 파일들이 {new_ext}로 변경되었습니다.")

    # 2. 설치 프로그램 목록 출력
    def list_installed_apps(self):
        print("\n--- [설치된 프로그램 목록] ---")
        for app in winapps.list_installed():
            print(f"- {app.name} (버전: {app.version})")

    # 3. 실행 중인 프로세스 목록 출력
    def list_running_processes(self):
        print("\n--- [실행 중인 프로세스 목록] ---")
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                print(f"PID: {proc.info['pid']} | Name: {proc.info['name']}")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

    # 4. Office 파일 비밀번호 설정
    def protect_office_file(self, file_path, password):
        try:
            file = OfficeFile(open(file_path, "rb"))
            file.encrypt(password, open(file_path, "wb"))
            print(f"🔒 {file_path}에 비밀번호가 설정되었습니다.")
        except Exception as e:
            print(f"❌ 오류 발생: {e}")

# 실행 예시
if __name__ == "__main__":
    butler = PCButler()
    # butler.list_installed_apps() # 필요한 기능의 주석을 풀어서 실행하세요.
