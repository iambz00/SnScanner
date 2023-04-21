import os, sys, re, time
import cv2, pytesseract
import numpy as np
import openpyxl
import argparse
from PIL import Image as PILImage
from PIL import ImageFont, ImageDraw

_VERSION = "20230421"
TESSERACT_PATH = "C:/Program Files/Tesseract-OCR/tesseract.exe"
OUTPUT_FILE = f"output_{time.strftime('%Y%m%d%H%M%S')}.xlsx"
SERIAL_PATTERN = r'R[A-Z0-9]{10}'

class SnScanner:
    def __init__(self, work_dir, tesseract_path="", output_file="", pattern="", interact=False):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.dir = work_dir
        self.pattern = pattern
        self.interact = interact
        if self.interact:
            from PIL import ImageGrab
            self.screenW, self.screenH = ImageGrab.grab().size
            del ImageGrab
            file_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
            self.font = ImageFont.truetype(os.path.join(file_dir, 'D2Coding-01.ttf'), 16)

        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active

        # 결과 파일명
        self.output_file = output_file
    def scan(self):
        self.worksheet.append(["파일명", "이미지", "1차인식", "2차인식", "특이사항"])
        self.workrow = 2    # 엑셀 현재 행
        for entry in os.scandir(self.dir):
            if entry.is_file() and not entry.name.startswith('.') and not entry.name.startswith('~'):
                print("* " + entry.name)
                self.scan_tesseract(entry.name)

        # 컬럼 폭 조정
        for column in self.worksheet.columns:
            length = max(len(str(cell.value)) for cell in column)
            self.worksheet.column_dimensions[column[0].column_letter].width = length * 1.2
        self.worksheet.column_dimensions['B'].width = 24

        # 시리얼번호는 고정폭인 Consolas로 설정
        consolas = openpyxl.styles.Font(name="Consolas")
        for cell in self.worksheet['C:C']: cell.font = consolas
        for cell in self.worksheet['D:D']: cell.font = consolas
        self.workbook.save(self.output_file)
    def scan_tesseract(self, file):
        '''
        1차 인식: 전체 그림에서 시리얼 넘버로 추정되는 문자열을 인식
        2차 인식: 1차 인식된 영역만 잘라내서 전처리 후 다시 인식
        '''
        #alpha = 2.0
        
        try:
            # cv2.imread 는 한글 파일을 읽지 못함
            image_array = np.fromfile(self.dir + '\\' + file, np.uint8)
            image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
        except:
            print("파일 열기 오류")
            return

        if type(image).__name__.lower() == "NoneType".lower():
            print("사진이 아닙니다.")
            return

        # 그레이스케일로 변환
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        # Gaussian Blur 적용 - 태블릿 화면의 격자를 완화
        gray = cv2.GaussianBlur(gray, (5,5), 0)
        # Bilaternal Filter 적용 - 격자 완화 + 경계선 강화
        gray = cv2.bilateralFilter(gray, 15, 15, 15)

        ''' 메모 - 아래 전처리는 인식율을 오히려 떨어트린다.
        1. Contrast 증가
        gray = np.clip((1+alpha)*gray - 128*alpha, 0, 255).astype(np.uint8)

        2. Thresholding
        r, gray = cv2.threshold(gray, -1, 255, cv2.THRESH_OTSU)

        3. Resizing
        w, h = gray.shape[1], gray.shape[0]
        gray = cv2.resize(gray, (w//2, h//2), cv2.INTER_LANCZOS4)
        '''
        # 1차 인식
        result = pytesseract.image_to_data(gray, lang="eng+kor", output_type=pytesseract.Output().DICT, config='--oem 3 --psm 11')
        texts = result['text']

        # Bounding box 표시
        if self.interact:
            # 화면 크기의 80%보다 큰 경우 축소
            WIDTH, HEIGHT = int(0.8 * self.screenW), int(0.8 * self.screenH)
            h, w = gray.shape
            f = min(HEIGHT / h, WIDTH / w, 1)
            if f < 1:
                shot = cv2.resize(gray, (int(f * w), int(f * h)), cv2.INTER_LANCZOS4)
            else:
                shot = gray.copy()
            shot = cv2.cvtColor(shot, cv2.COLOR_GRAY2BGR)
            for i in range(len(texts)):
                # Confidence 30% 이상인 항목에 대해서만 표시
                if texts[i] and result['conf'][i] > 30:
                    l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
                    box = np.array([[l,t], [l+w,t], [l+w,t+h], [l,t+h]])
                    # 축소한 경우 Bounding box 도 축소
                    box = np.array(f * box, np.int32)
                    # Confidence에 따라 색상 변경 - 높을수록 빨갛게
                    conf = result['conf'][i] / 100
                    b, g, r = int((1-conf)*128), int((1-conf)*128), int(conf*255)
                    cv2.polylines(shot, [box], True, (b, g, r), 1)
                    # 한글은 PIL로 그려야 함
                    pil_shot = PILImage.fromarray(shot)
                    draw = ImageDraw.Draw(pil_shot)
                    draw.text((box[0,0], box[0,1]), texts[i], font=self.font, fill=(b,g,r,0))
                    shot = np.array(pil_shot, np.uint8)
            self.imshow(shot, title=file)

        # 2차 인식
        sn_candidate = []
        for i in range(len(texts)):
            # 시리얼 형식에 맞는 결과 검색
            match = re.search(self.pattern, texts[i])
            if match:
                text = match.group(0).replace("O","0") # 모든 O는 0으로 강제 변환

                # Bounding box를 구한다 - left, top, width, height
                l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
                print(text + " > ", end="", flush=True)

                # Bounding box 외부로 여유공간이 있어야 인식이 잘 되므로
                # 상하좌우 4px 씩 더 가져온 후
                image_height, image_width = gray.shape
                crop = gray[max(0, t-4):min(t+h+4, image_height-1), max(0, l-4):min(l+w+4, image_width-1)]
                rows, cols = crop.shape
                # 좌상단 픽셀과 같은 색으로 외부를 32픽셀씩 둘러싸 준다
                newpage = crop[0,0] * np.ones( (rows+64)*(cols+64), np.uint8).reshape(rows+64, cols+64)
                for y in range(rows):
                    for x in range(cols):
                        newpage[y+32, x+32] = crop[y, x]
                # 추가로 전처리를 할 수록 인식율이 떨어지므로 가만 놔두자...
                # 대비(Contrast)를 준다
                #cont = np.clip((1+alpha)*newpage - 128*alpha, 0, 255).astype(np.uint8)
                # OTSU 방식으로 Thresholding
                #r, thres = cv2.threshold(newpage, -1, 255, cv2.THRESH_OTSU)
                cont = newpage
                # 인식
                newtext = pytesseract.image_to_string(cont, lang="eng", config='--oem 3 --psm 7').strip().replace("O","0").replace("\n","")
                if len(newtext) > 2 and newtext[1] == "S":
                    newtext = newtext[:1] + "5" + newtext[2:]
                newmatch = re.search(self.pattern, newtext)
                if newmatch:
                    newtext = newmatch.group(0)
                different = (text != newtext) and "1차/2차 불일치" or ""

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
    def imshow(self, *imgs, title='test'):
        i = 1
        for img in imgs:
            if title == "test":
                title += str(i)
            cv2.imshow(title, img)
            i += 1
        cv2.waitKey(0)
        cv2.destroyAllWindows()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(usage="%(prog)s [옵션] 경로")
    parser.add_argument("path", metavar="경로", help="대상 이미지가 있는 폴더")
    parser.add_argument("-t", metavar="Tesseract경로", help=f"Teserract 위치. 기본값: {TESSERACT_PATH}", default=TESSERACT_PATH,
                        dest="tesseract_path")
    parser.add_argument("-o", metavar="파일명", help="출력 파일명. 기본값: ouput_날짜시간.xlsx", default=OUTPUT_FILE,
                        dest="output_file")
    parser.add_argument("-p", metavar="패턴", help=f"검출 패턴(Python 정규식) 기본값: 태블릿 시리얼 검출용 '{SERIAL_PATTERN}'",
                        default=SERIAL_PATTERN, dest="pattern")
    parser.add_argument("-i", help="각 파일마다 인식 영역과 문자를 확인하면서 넘어갑니다.", action="store_true",
                        dest="interact")
    args = parser.parse_args(args=None if sys.argv[1:] else ['-h'])

    tesseract_path = args.tesseract_path.replace("\\", "\\\\")
    output_file = args.output_file.replace("\\", "\\\\")
    pattern = args.pattern.replace("\\", "\\\\")

    if not os.path.exists(args.path):
        print(f"[{args.path}] 폴더가 존재하지 않습니다.")
        exit(1)
    if not os.path.exists(tesseract_path):
        print(f"[{tesseract_path}] 경로에 Tesseract가 없습니다.")
        exit(2)

    ss = SnScanner(args.path, tesseract_path=tesseract_path, output_file=output_file, pattern=pattern, interact=args.interact)
    ss.scan()
