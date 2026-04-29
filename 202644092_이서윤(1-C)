import os
import hashlib
import psutil
import winapps
from msoffice_crypt import OfficeFile

class PCButler:
    def __init__(self):
        print("--- AI Agent 기반 PC 관리 도구 실행 ---")

    # [추가 기능] 중복 파일 찾기
    def find_duplicate_files(self, folder_path):
        if not os.path.exists(folder_path):
            print("❌ 경로를 찾을 수 없습니다.")
            return

        hashes = {}  # {hash_value: [file_paths]}
        duplicates = []

        print(f"🔍 {folder_path} 내 중복 파일을 검색 중입니다...")

        for root, dirs, files in os.walk(folder_path):
            for filename in files:
                path = os.path.join(root, filename)
                
                # 파일의 해시값 계산 (파일이 클 경우를 대비해 블록 단위로 읽기)
                hasher = hashlib.md5()
                try:
                    with open(path, 'rb') as f:
                        buf = f.read(65536)
                        while len(buf) > 0:
                            hasher.update(buf)
                            buf = f.read(65536)
                    
                    file_hash = hasher.hexdigest()

                    if file_hash in hashes:
                        hashes[file_hash].append(path)
                        duplicates.append(path)
                    else:
                        hashes[file_hash] = [path]
                except (PermissionError, OSError):
                    continue

        if not duplicates:
            print("✨ 중복된 파일이 없습니다!")
        else:
            print(f"\n📢 총 {len(duplicates)}개의 중복 파일 그룹을 발견했습니다.")
            for h, paths in hashes.items():
                if len(paths) > 1:
                    print(f"\n[중복 그룹 - Hash: {h}]")
                    for p in paths:
                        print(f" > {p}")
        
        return duplicates

    # (기존 기능들: 확장자 변경, 프로그램 목록, 프로세스 목록 등은 동일하게 유지)
    def list_running_processes(self):
        print("\n[현재 실행 중인 프로세스 목록]")
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                print(f"PID: {proc.info['pid']:<6} | Name: {proc.info['name']}")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

# 실행부
if __name__ == "__main__":
    butler = PCButler()
    
    # 1. 중복 파일 찾기 테스트 (경로는 본인의 폴더로 수정하세요)
    # target_dir = r"C:\Users\YourName\Documents"
    # butler.find_duplicate_files(target_dir)
    
    # 2. 실행 중인 프로세스 출력
    butler.list_running_processes()
