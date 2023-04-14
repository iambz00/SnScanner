import os, sys, re, time
import cv2, pytesseract
import numpy as np
import openpyxl
from PIL import Image as PILImage

_VERSION = "20230414"
TESSERACT_PATH = "C:/Program Files/Tesseract-OCR/tesseract.exe"
'''
삼성 시리얼번호 규칙

예시) R54T1067RRR
해당 S/N의 형식은 새로운 삼성 테블릿 모델의 S/N 규칙에 따라 구성됩니다. 따라서 다음과 같은 의미를 가지게 됩니다:

R은 "Region/Country Code"를 나타내며, 해당 제품이 한국에서 제조된 것을 나타냅니다.
54는 "Location Code"를 나타내며, 해당 제품이 한국 내에서 어떤 지역에서 생산되었는지를 나타냅니다.
T는 "Year Code"를 나타내며, 해당 제품이 2021년에 생산된 것을 나타냅니다.
10은 "Month Code"를 나타내며, 해당 제품이 10월에 생산된 것을 나타냅니다.
67은 "Production Code"를 나타내며, 해당 제품이 생산된 라인을 나타냅니다.
따라서, 해당 S/N인 R54T1067RRR은 2021년 10월 한국 내에서 생산된 삼성 테블릿 제품이며, 해당 제품의 일련번호는 RRR입니다.

'''
class TextScanner:
    def __init__(self, work_dir, sn_only=False):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        self.dir = work_dir
        self.sn_only = sn_only
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active

        # 결과 파일명
        date_time = time.strftime("%Y%m%d%H%M%S")
        self.output_file = f"output_{date_time}.xlsx"

    def scan(self):
        files = os.listdir(self.dir)
        self.worksheet.append(["파일명", "이미지", "1차인식", "2차인식", "특이사항"])
        self.workrow = 2    # 엑셀 현재 행
        for file in files:
            if not self.sn_only:
                print("* " + file)
            self.scan_tesseract(file)

        # 컬럼 폭 조정
        for column in self.worksheet.columns:
            length = max(len(str(cell.value)) for cell in column)
            self.worksheet.column_dimensions[column[0].column_letter].width = length * 1.2
        self.worksheet.column_dimensions['B'].width = 24

        self.workbook.save(self.output_file)
    def scan_tesseract(self, file):
        '''
        1차 인식: 전체 그림에서 시리얼 넘버로 추정되는 문자열을 인식
        2차 인식: 1차 인식된 영역만 잘라내서 전처리 후 다시 인식
        '''
        alpha = 2.0
        # cv2.imread 는 한글 파일을 읽지 못함
        image_array = np.fromfile(self.dir + '\\' + file, np.uint8)
        image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
        # 그레이스케일로 변환
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # 1차 인식
        result = pytesseract.image_to_data(gray, lang="eng+kor", output_type=pytesseract.Output().DICT, config='--oem 3 --psm 11')
        texts = result['text']

        sn_candidate = []
        for i in range(len(texts)):
            # 시리얼 형식에 맞는 결과 검색
            match = re.search(r'R[A-Z0-9]{10}', texts[i])
            if match:
                text = match.group(0).replace("O","0") # 모든 O는 0으로 강제 변환
                #if text[1] == "S":  # 두번째 나오는 S는 5로 강제 변환
                #    text = text[:1] + "5" + text[2:]

                # Bounding box를 구한다 - left, top, width, height
                l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
                print(text + " > ", end="", flush=True)

                # Bounding box 외부로 여유공간이 있어야 인식이 잘 되므로
                # 좌 4px, 상 4px 만큼 바운딩박스를 키우고
                crop = gray[max(0, t-4):t+h, max(0, l-4):l+w]
                rows, cols = crop.shape
                # 좌상단 픽셀과 같은 색으로 외부를 32픽셀씩 둘러싸 준다
                newpage = crop[0,0] * np.ones( (rows+64)*(cols+64), np.uint8).reshape(rows+64, cols+64)
                for y in range(rows):
                    for x in range(cols):
                        newpage[y+32, x+32] = crop[y, x]
                # 대비(Contrast)를 준다
                cont = np.clip((1+alpha)*newpage - 128*alpha, 0, 255).astype(np.uint8)
                # 2차 인식
                newtext = pytesseract.image_to_string(cont, lang="eng", config='--oem 3 --psm 7').strip().replace("O","0").replace("\n","")
                if len(newtext) > 2 and newtext[1] == "S":
                    newtext = newtext[:1] + "5" + newtext[2:]
                newmatch = re.search(r'R[S5][A-Z0-9]{9}', newtext)
                if newmatch:
                    newtext = newmatch.group(0)
                different = (text != newtext) and "1차 2차 불일치" or ""

                # 출력
                print(newtext, different)
                self.worksheet.append([file, "", text, newtext, different])
                pilImage = self.toPILImage(crop)
                if pilImage:
                    xlImage = openpyxl.drawing.image.Image(pilImage)
                    self.worksheet.add_image(xlImage, 'B' + str(self.workrow))
                else:
                    self.worksheet['B' + str(self.workrow)] = "직접 확인해 주세요"
                self.workrow += 1

                sn_candidate.append(newtext)
        if len(sn_candidate) < 1:
            # 아예 인식이 안된 경우
            print("미인식 - 확인 필요")
            self.worksheet.append([file, "직접 확인해주세요", "", "", "미인식"])
            self.workrow += 1
    def toPILImage(self, img):
        pil_image = PILImage.fromarray(img)
        w, h = pil_image.width, pil_image.height
        # Resize to w < 240, h < 32
        if w < 1200 and h < 160:
            ratio = min(160 / w, 22 / h)
            pil_image = pil_image.resize((int(ratio * pil_image.width), int(ratio * pil_image.height)))
            # openpyxl 버그 회피용 코드, 이렇게 하지 않으면 저장 시 오류가 난다
            pil_image.fp = openpyxl.drawing.image.BytesIO()
            pil_image.save(pil_image.fp, format="png")
            return pil_image
        else:
            return None
    def imwrite(self, img, file):
        result, encoded_img = cv2.imencode(os.path.splitext(file)[1], img)
        if result:
            f = open(self.output_dir + "/확인_" + file, mode="w+b")
            encoded_img.tofile(f)
            f.close()

def s(*imgs, title='test'):
    i = 1
    for img in imgs:
        print(i)
        cv2.imshow(title + str(i), img)
        i += 1
    cv2.waitKey(0)
    cv2.destroyAllWindows()

if __name__ == "__main__":
    image_dir = "img"
    sn_only = False
    if len(sys.argv) == 1:
        print("SnScanner [-s] [directory]")
    else:
        for i in range(1, len(sys.argv)):
            if sys.argv[i] == "-s":
                sn_only = True
            else:
                image_dir = sys.argv[i]
        if os.path.exists(image_dir):
            ts = TextScanner(image_dir, sn_only=sn_only)
            ts.scan()
        else:
            print(f"[{image_dir}] 폴더가 존재하지 않습니다.")
