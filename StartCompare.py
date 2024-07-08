import glob
import os
import subprocess
import argparse

def main(folder_path_1, folder_path_2):
    # 첫 번째 폴더 경로 출력
    print(f"첫 번째 폴더 경로: {folder_path_1}")

    # 두 번째 폴더 경로 출력
    print(f"두 번째 폴더 경로: {folder_path_2}")

    # 폴더 1의 PNG 파일 리스트 가져오기
    image_files_1 = sorted(glob.glob(os.path.join(folder_path_1, '*.png')))
    # 폴더 2의 PNG 파일 리스트 가져오기
    image_files_2 = sorted(glob.glob(os.path.join(folder_path_2, '*.png')))

    # 두 폴더의 파일 이름 비교
    image_files_1_names = [os.path.basename(f) for f in image_files_1]
    image_files_2_names = [os.path.basename(f) for f in image_files_2]

    # 두 폴더에 동일한 파일 이름이 있는지 확인하고 짝짓기
    image_pairs = [(os.path.join(folder_path_1, f), os.path.join(folder_path_2, f)) for f in image_files_1_names if f in image_files_2_names]

    # ImageCompare.py 스크립트 반복 실행
    for first_image, second_image in image_pairs:
        subprocess.run(['python', 'ImageCompare.py', '--first', first_image, '--second', second_image])

if __name__ == '__main__':
    # argparse를 사용하여 커맨드 라인 인자 파싱
    parser = argparse.ArgumentParser(description='두 폴더의 이미지를 비교합니다.')
    parser.add_argument('folder1', type=str, help='첫 번째 폴더의 경로')
    parser.add_argument('folder2', type=str, help='두 번째 폴더의 경로')
    args = parser.parse_args()

    # main 함수 호출
    main(args.folder1, args.folder2)
