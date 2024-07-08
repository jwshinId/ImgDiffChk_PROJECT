import win32com.client

def add_picture_and_save_ppt(pptx_path, image_path, output_path):
    # PowerPoint 애플리케이션 열기
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    
    # PowerPoint 파일 열기
    prs = ppt_app.Presentations.Open(pptx_path)
    
    # 첫 번째 슬라이드 선택 (1부터 시작)
    slide = prs.Slides(1)
    print("이미지 경로:", image_path)
    # 이미지 추가
    left = 100  # 이미지의 왼쪽 위치
    top = 100   # 이미지의 위쪽 위치
    width = 400  # 이미지의 너비
    height = 300  # 이미지의 높이
    shape = slide.Shapes.AddPicture(image_path, left, top, 5, 3)
    
    # 이미지 사이즈 조정
    shape.Width = 405  # 새로운 너비 설정
    shape.Height = 203 # 새로운 높이 설정
    shape.left = 55
    shape.top = 143
    
    # PowerPoint 파일 다른 이름으로 저장
    prs.SaveAs(output_path)
    
    # PowerPoint 애플리케이션 종료
    ppt_app.Quit()

if __name__ == "__main__":
    pptx_path = r"D:/BeforeAfterPicture/FMCS_BeforeAfter_Sample.pptx"  # 원본 PowerPoint 파일 경로
    image_path = r"D:\\BeforeAfterPicture\\2040708\\before\\SCR_C1ELEC_BUSWAY_DTS.png"  # 삽입할 이미지 파일 경로
    output_path = r"D:/BeforeAfterPicture/FMCS_BeforeAfter_20240708_new.pptx"  # 저장할 PowerPoint 파일 경로
    add_picture_and_save_ppt(pptx_path, image_path, output_path)
