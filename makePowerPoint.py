# startcompare.py

import glob
import os
import argparse
import cv2
from skimage.metrics import structural_similarity as compare_ssim
import imutils
from win32com.client import Dispatch

def compare_images(args):
    # 이미지 파일 명칭 추출
    imgAname = os.path.basename(args["first"])
    imgBname = os.path.basename(args["second"])
    
    # 첫 번째와 두 번째 이미지 로드
    if imgAname != "Empty":
        imageA = cv2.imread(args["first"])
    if imgBname != "Empty":
        imageB = cv2.imread(args["second"])
    
    # 첫 번째 이미지 사이즈 측정 (기준)
    if imgAname != "Empty":
        height, width = imageA.shape[:2]
    else:
        height, width = imageB.shape[:2]
    
    if imgAname != "Empty" and imgBname != "Empty":
        if imageA.shape[:2] != imageB.shape[:2]:
            print(f"사이즈가 다릅니다. 이미지를 리사이즈합니다.")
            imageA = cv2.resize(imageA, (width, height), interpolation=cv2.INTER_AREA)
            imageB = cv2.resize(imageB, (width, height), interpolation=cv2.INTER_AREA)
    
    print(f"첫 번째 이미지 크기: width={width}, height={height}")
    
    if height < width:
        height = 203
        width = 405
        f_left = 55
        f_top = 143
        s_left = 491
        s_top = 143
    else:
        height = 300
        width = 241
        f_left = 142
        f_top = 98
        s_left = 580
        s_top = 98    

    if imgAname == imgBname:
        # 이미지를 그레이스케일로 변환
        grayA = cv2.cvtColor(imageA, cv2.COLOR_BGR2GRAY)
        grayB = cv2.cvtColor(imageB, cv2.COLOR_BGR2GRAY)

        # Structural Similarity Index (SSIM) 계산
        (score, diff) = compare_ssim(grayA, grayB, full=True)
        diff = (diff * 255).astype("uint8")
        print("SSIM: {}".format(score))

        # 차이 이미지를 이진화하여 컨투어 추출
        thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = imutils.grab_contours(cnts)

    # PowerPoint 파일이 존재하는지 확인
    pptx_path = args["pptx"]
    powerpoint = Dispatch("PowerPoint.Application")

    if args.get("index") == '2':
        presentation = powerpoint.Presentations.Open(pptx_path)
        powerpoint.Visible = True # PowerPoint 창을 보이도록 설정
    else:
        # 이미 열린 프레젠테이션 확인
        for open_presentation in powerpoint.Presentations:
            if open_presentation.FullName.lower() == pptx_path.lower():
                presentation = open_presentation
                break

    # 첫 번째 슬라이드 인덱스
    first_slide_index = 1
    
    # 다음 슬라이드 인덱스 지정
    slide_index = int(args["index"])

    # 첫 번째 슬라이드 가져오기
    first_slide = presentation.Slides(first_slide_index)
    first_slide.Copy()
    
    # 다음 슬라이드로 붙여넣기
    # Paste는 붙여넣어진 슬라이드의 배열을 반환합니다
    new_slides = presentation.Slides.Paste(Index=slide_index)

    # 붙여넣어진 슬라이드 참조
    slide = new_slides[0]

    if imgAname != "Empty":
        # 첫 번째 이미지 추가
        shape_first = slide.Shapes.AddPicture(FileName=args["first"], LinkToFile=0, SaveWithDocument=-1, Left=-1, Top=-1)

        # 첫 번째 이미지 사이즈 조정
        shape_first.LockAspectRatio = 0 # 비율 해제
        shape_first.Width = width   # 새로운 너비 설정
        shape_first.Height = height # 새로운 높이 설정
        shape_first.left = f_left   # 새로운 왼쪽 여백
        shape_first.top = f_top     # 새로운 위쪽 여백

    if imgBname != "Empty":
        # 두 번째 이미지 추가
        shape_second = slide.Shapes.AddPicture(FileName=args["second"], LinkToFile=0, SaveWithDocument=-1, Left=-1, Top=-1)
        
        # 두 번째 이미지 사이즈 조정
        shape_second.LockAspectRatio = 0 # 비율 해제
        shape_second.Width = width   # 새로운 너비 설정
        shape_second.Height = height # 새로운 높이 설정
        shape_second.left = s_left   # 새로운 왼쪽 여백
        shape_second.top = s_top     # 새로운 위쪽 여백

    if imgAname == imgBname:
        # 차이를 강조하기 위해 컨투어를 반복하면서 슬라이드에 사각형 추가
        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)

            left_rect = s_left + (x / imageA.shape[1]) * width  # 이미지 B에서의 왼쪽 시작점
            top_rect = s_top + (y / imageA.shape[0]) * height    # 이미지 B에서의 위쪽 시작점
            
            width_rect = (w / imageA.shape[1]) * width               # 사각형의 너비
            height_rect = (h / imageA.shape[0]) * height              # 사각형의 높이

            # 슬라이드에 사각형 추가
            shape = slide.Shapes.AddShape(1, left_rect, top_rect, width_rect, height_rect)  # 1은 msoShapeRectangle
            shape.Line.ForeColor.RGB = RGB(255, 0, 0)  # RGB 함수를 사용하여 색상 설정
            shape.Line.Weight = 2  # 선 두께 설정
            
            # 채우기 없음 설정
            shape.Fill.Transparency = 1.0  # 채우기의 투명도를 1로 설정하여 채우기 없음으로 만듦

    if args.get("isend") == "True":
        # 베이스 슬라이드 삭제
        presentation.Slides(1).Delete()
        
        # 수정된 프레젠테이션을 새 파일로 저장
        presentation.SaveAs(args["output"])

        # PowerPoint 애플리케이션 종료
        powerpoint.Quit()

def RGB(red, green, blue):
    # RGB 값을 반환하는 함수
    return (blue << 16) + (green << 8) + red

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
            
    # ImageCompare 기능 직접 호출
    for index, (first_image, second_image) in enumerate(image_pairs):
        pptx_path = 'D:\\BeforeAfterPicture\\FMCS_BeforeAfter_Sample.pptx'  # 여기에 PPTX 파일 경로 설정
        save_path = 'D:\\BeforeAfterPicture\\FMCS_BeforeAfter_20240708.pptx'  # 여기에 출력 경로 설정
        args = {
            "first": first_image,
            "second": second_image,
            "pptx": pptx_path,
            "output": save_path,
            "isend": "True" if index == len(image_pairs)-1 else "False",
            "index": str(index + 2)
        }
        try:
            compare_images(args)
        except Exception as e:
            print(f"ImageCompare 기능 실행 중 오류 발생: {e}")

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
