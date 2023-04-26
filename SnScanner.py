import os, sys, re, time
import cv2, pytesseract
import numpy as np
import openpyxl
import argparse
from PIL import Image as PILImage
from PIL import ImageFont, ImageDraw

_VERSION = "20230426"
TESSERACT_PATH = "C:/Program Files/Tesseract-OCR/tesseract.exe"
OUTPUT_FILE = f"output_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
SERIAL_PATTERN = r'R[A-Z0-9]{10}'

class SnScanner:
    def __init__(self, work_dir, tesseract_path="", output_file="", pattern="", interact=False, samsung=False):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.dir = work_dir
        self.pattern = pattern
        self.interact = interact
        self.samsung = samsung
        if self.interact:
            from PIL import ImageGrab
            self.screenW, self.screenH = ImageGrab.grab().size
            del ImageGrab
            file_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
            self.font = ImageFont.truetype('malgun.ttf', 12)
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        # 결과 파일명
        self.output_file = output_file

    def scan(self):
        #self.worksheet.append(["파일명", "이미지", "1차인식", "2차인식", "1차신뢰도", "2차신뢰도", "특이사항"])
        self.worksheet.append(["파일명", "이미지", "시리얼", "특이사항"])
        self.workrow = 2    # 엑셀 현재 행
        for entry in os.scandir(self.dir):
            if entry.is_file() and not entry.name.startswith('.') and not entry.name.startswith('~'):
                print("* " + entry.name)
                self.scan_file(entry.name)

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
    def scan_file(self, file):
        '''
        1차 인식: 전체 그림에서 시리얼 넘버로 추정되는 문자열을 인식
        2차 인식: 1차 인식된 영역만 잘라내서 전처리 후 다시 인식
        '''
        # cv2.imread 는 한글 파일을 읽지 못함. np.fromfile 이용
        try:
            image_array = np.fromfile(self.dir + '\\' + file, np.uint8)
            original_image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
        except:
            print("파일 열기 오류")
            return

        if type(original_image).__name__.lower() == "NoneType".lower():
            print("사진이 아닙니다.")
            return

        ''' 메모 - 전처리는 오히려 인식율을 떨어트리니 안하는 게 낫다.
        1. Contrast 증가
        alpha = 1.0
        image = np.clip((1+alpha)*image - 128*alpha, 0, 255).astype(np.uint8)

        2. Thresholding
        r, image = cv2.threshold(image, -1, 255, cv2.THRESH_OTSU)

        3. Resizing
        w, h = image.shape[1], image.shape[0]
        image = cv2.resize(image, (w//2, h//2), cv2.INTER_LANCZOS4)

        4. Grayscale 변환도 필요없다.
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        '''
        # Gaussian Blur 적용 - 태블릿 화면의 격자를 완화
        image = cv2.GaussianBlur(original_image, (5,5), 0)
        # Bilaternal Filter 적용 - 격자 완화 + 경계선 강화
        image = cv2.bilateralFilter(image, 15, 15, 15)

        # 1차 인식
        result = pytesseract.image_to_data(image, lang="eng+kor", output_type=pytesseract.Output().DICT, config='--oem 3 --psm 11')
        texts, confs = result['text'], result['conf']

        # 패턴에 부합하는 결과만 정리
        sn_candidate = []
        for i in range(len(texts)):
            match = re.search(self.pattern, texts[i])
            if match:
                text = self.refineSN(texts[i])
                conf = confs[i]
                sn_candidate.append((i, text, conf))

        # 인식 결과가 없으면 그레이스케일로 변환해서 재인식
        if not sn_candidate:
            result = pytesseract.image_to_data(cv2.cvtColor(image, cv2.COLOR_BGR2GRAY), lang="eng+kor", output_type=pytesseract.Output().DICT, config='--oem 3 --psm 11')
            texts, confs = result['text'], result['conf']
            for i in range(len(texts)):
                match = re.search(self.pattern, texts[i])
                if match:
                    text = self.refineSN(texts[i])
                    conf = confs[i]
                    sn_candidate.append((i, text, conf))

        # Bounding box 표시 - 원본 이미지에 작업
        if self.interact:
            hit_index = []
            for (i, _, _) in sn_candidate:
                hit_index.append(i)
            # 큰 이미지는 화면 크기의 80%로 축소
            WIDTH, HEIGHT = int(0.8 * self.screenW), int(0.8 * self.screenH)
            h, w, _ = original_image.shape
            f = min(HEIGHT / h, WIDTH / w, 1)
            if f < 1:
                box_image = cv2.resize(original_image, (int(f * w), int(f * h)), cv2.INTER_LANCZOS4)
            else:
                box_image = original_image.copy()
            overlay = box_image.copy()
            overlay_alpha = 0.8

            for i in range(len(texts)):
                # Confidence 30% 이상인 항목만 표시
                if texts[i] and confs[i] > 30:
                    l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
                    box = np.array([[l,t], [l+w,t], [l+w,t+h], [l,t+h]])
                    # 축소한 경우 Bounding box 도 축소
                    box = np.array(f * box, np.int32)
                    # SN은 빨갛게, 나머지는 녹색
                    (b, g, r) = (0, 0, 224) if hit_index.count(i) else (64, 255, 0)
                    cv2.polylines(overlay, [box], True, (b, g, r), 1)
                    # 한글은 PIL로 그려야 함
                    pil_box_image = PILImage.fromarray(overlay)
                    draw = ImageDraw.Draw(pil_box_image)
                    draw.text((box[0,0], box[3,1]), texts[i], font=self.font, fill=(b,g,r,0))
                    overlay = np.array(pil_box_image, np.uint8)
            # overlay layer에 그려서 합친다 = 투명효과
            box_image = cv2.addWeighted(overlay, overlay_alpha, box_image, 1 - overlay_alpha, 0)
            self.imshow(box_image, title=file)

        # 2차 인식
        for (i, text, conf) in sn_candidate:
            # Bounding box를 구한다 - left, top, width, height
            l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
            print(text + " > ", end="", flush=True)

            # Bounding box 외부로 여유공간이 있어야 인식이 잘 되므로
            # 상하좌우 4px 씩 더 가져온 후
            image_height, image_width, _ = image.shape
            crop = image[max(0, t-4):min(t+h+4, image_height-1), max(0, l-4):min(l+w+4, image_width-1)]
            original_crop = original_image[max(0, t-4):min(t+h+4, image_height-1), max(0, l-4):min(l+w+4, image_width-1)]
            rows, cols, _ = crop.shape

            # 좌상단 픽셀과 같은 색으로 외부를 32픽셀씩 둘러싸준다
            newpage = np.ones( (rows+64)*(cols+64)*3, np.uint8).reshape(rows+64, cols+64, 3)
            newpage[:] = crop[0,0]
            newpage[32:32+rows, 32:32+cols, :] = crop

            # 인식
            #newtext = pytesseract.image_to_string(newpage, lang="eng", config='--oem 3 --psm 7').strip().replace("O","0").replace("\n","")
            result2 = pytesseract.image_to_data(newpage, lang="eng", output_type=pytesseract.Output().DICT, config='--oem 3 --psm 7')
            text_index = result2['word_num'].index(1)
            newtext = self.refineSN(result2['text'][text_index])
            newconf = result2['conf'][text_index]

            fittext = newtext
            note = ""
            if text != fittext:
                if len(text) < len(fittext):
                    fittext = text
                #note = "확인 필요"

            # 출력
            print(fittext, note)
            #self.worksheet.append([file, "", text, newtext, conf, newconf, note])
            self.worksheet.append([file, "", fittext, note])
            pilImage = self.toPILImage(original_crop)
            if pilImage:
                xlImage = openpyxl.drawing.image.Image(pilImage)
                self.worksheet.add_image(xlImage, 'B' + str(self.workrow))
            else:
                self.worksheet['B' + str(self.workrow)] = "직접 확인해 주세요"
            self.workrow += 1

        # 그레이스케일로도 인식이 안되었으면
        if not sn_candidate:
            print("미인식 - 확인 필요")
            self.worksheet.append([file, "확인 필요", "", "미인식"])
            self.workrow += 1
    def toPILImage(self, img):
        pil_image = PILImage.fromarray(img)
        w, h = pil_image.width, pil_image.height
        # 엑셀파일 안에 잘 들어가도록 크기조절
        if w < 1200 and h < 160:
            ratio = min(160 / w, 22 / h)
            pil_image = pil_image.resize((int(ratio * pil_image.width), int(ratio * pil_image.height)))
            # openpyxl 버그 회피용 코드, BytesIO를 미리 작업해줘야 오류가 안남
            pil_image.fp = openpyxl.drawing.image.BytesIO()
            pil_image.save(pil_image.fp, format="png")
            return pil_image
        else:
            return None
    def refineSN(self, text):
        # 모든 O를 0으로 바꾸고 
        text = text.strip().replace("O","0").replace("\n","")
        # 앞에 내용을 뗀다(S/N: 등)
        match = re.search(self.pattern + ".*", text)
        text = match.group(0) if match else text

        # 삼성태블릿 인식 시 보완
        if self.samsung:
            # RS4 -> R54
            if len(text) > 2 and text[1] == "S":
                text = text[:1] + "5" + text[2:]
            # Tesseract 버그 추정 RS54, R554 -> R54
            if len(text) > 11 and text.startswith(('R55','R5S')):
                text = text[:1] + "5" + text[3:]
        return text

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
    parser.add_argument("--samsung", help="삼성 태블릿인 경우 인식율 강화", action="store_true", dest="samsung")
    #parser.add_argument("-v", help="출력 파일에 상세 내용 표시", action="store_true", dest="verbose")
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

    ss = SnScanner(args.path, tesseract_path=tesseract_path, output_file=output_file,
                    pattern=pattern, interact=args.interact, samsung=args.samsung)
    ss.scan()
