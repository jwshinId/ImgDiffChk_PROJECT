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
    image_files_1_names = {os.path.basename(f): f for f in image_files_1}
    image_files_2_names = {os.path.basename(f): f for f in image_files_2}

    # 두 폴더에 동일한 파일 이름이 있는지 확인하고 짝짓기
    image_pairs = []
    
    # hashmap 생성
    hm_befImage = {}
    hm_aftImage = {}
    
    # 둘 다 변경점이 있을 경우
    for name in image_files_1_names:
        if name in image_files_2_names:
            image_pairs.append((image_files_1_names[name], image_files_2_names[name]))
            hm_befImage[name] = True
            hm_aftImage[name] = True
    
    # 변경 전 화면만 있을 경우
    for name in image_files_1_names:
        if not name in hm_befImage:
            image_pairs.append((image_files_1_names[name], "Empty"))
            hm_befImage[name] = True

    # 변경 후 화면만 있을 경우
    for name in image_files_2_names:
        if not name in hm_aftImage:
            image_pairs.append(("Empty", image_files_2_names[name]))
            hm_aftImage[name] = True
            
    # ImageCompare.py 스크립트 반복 실행
    for index, (first_image, second_image) in enumerate(image_pairs):
        pptx_path = 'D:\\BeforeAfterPicture\\FMCS_BeforeAfter_Sample.pptx'  # 여기에 PPTX 파일 경로 설정
        save_path = 'D:\\BeforeAfterPicture\\FMCS_BeforeAfter_20240708.pptx'  # 여기에 출력 경로 설정
        try:
            if index == len(image_pairs)-1:
                subprocess.run(['python', 'ImageCompare.py', '--first', first_image, '--second', second_image, '--pptx', pptx_path, '--output', save_path, '--isend', "True", '--index', str(index+2)], check=True)
            else:
                subprocess.run(['python', 'ImageCompare.py', '--first', first_image, '--second', second_image, '--pptx', pptx_path, '--output', save_path, '--isend', "False", '--index', str(index+2)], check=True)    
        except subprocess.CalledProcessError as e:
            print(f"ImageCompare.py 실행 중 오류 발생: {e}")
            # 오류 처리 추가

if __name__ == '__main__':
    # argparse를 사용하여 커맨드 라인 인자 파싱
    parser = argparse.ArgumentParser(description='두 폴더의 이미지를 비교합니다.')
    parser.add_argument('folder1', type=str, help='첫 번째 폴더의 경로')
    parser.add_argument('folder2', type=str, help='두 번째 폴더의 경로')
    args = parser.parse_args()

    # 폴더 경로 유효성 검사 (필요시)
    if not os.path.isdir(args.folder1) or not os.path.isdir(args.folder2):
        print("폴더 경로가 올바르지 않습니다.")
        exit(1)

    # main 함수 호출
    main(args.folder1, args.folder2)
