import os, sys, re, time
import cv2, pytesseract
import numpy as np

_VERSION = "20230413"
TESSERACT_PATH = "C:/Program Files/Tesseract-OCR/tesseract.exe"

class TextScanner:
    def __init__(self, work_dir, sn_only=False):
        #self.reader = easyocr.Reader(['en', 'ko'])
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        self.tesseract_config = ('--oem 3 --psm 11')
        self.dir = work_dir
        self.sn_only = sn_only
        date_time = time.strftime("%Y%m%d%H%M%S")
        # 결과 파일명 및 디렉토리명
        self.output_dir = f"output_{date_time}"
        if not os.path.exists(self.output_dir):
            os.mkdir(self.output_dir)
        self.output_file = open(self.output_dir + "/output.csv", "w")
    def scan(self):
        files = os.listdir(self.dir)
        self.output_file.write("파일명,1차인식,2차인식,불일치\n")
        for file in files:
            if not self.sn_only:
                print("* " + file)
            self.scan_tesseract(file)
    def scan_tesseract(self, file):
        alpha = 2.0
        # cv2.imread 는 한글 파일을 읽지 못함
        image_array = np.fromfile(self.dir + '\\' + file, np.uint8)
        image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
        # 그레이스케일로 변환
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # 1차 인식
        result = pytesseract.image_to_data(gray, lang="eng+kor", output_type=pytesseract.Output().DICT, config=self.tesseract_config)
        texts = result['text']

        sn_candidate = []
        for i in range(len(texts)):
            match = re.search(r'R[S5]4[A-Z0-9]{8}', texts[i])
            if match:
                text = match.group(0).replace("O","0") # 모든 O는 0으로 강제 변환
                if text[1] == "S":
                    text = text[:1] + "5" + text[2:]
                # 일련번호 형식에 맞는 결과를 찾아서 Bounding box를 구한다 - left, top, width, height
                l,t,w,h = result['left'][i], result['top'][i], result['width'][i], result['height'][i]
                print(text + " > ", end="", flush=True)

                # Bounding box 외부로 약간 여유공간이 있어야 인식이 잘 된다
                crop = gray[t:t+h, l:l+w]
                rows, cols = crop.shape
                newpage = 255 * np.ones( (rows+128)*(cols+128), np.uint8).reshape(rows+128, cols+128)
                for y in range(rows):
                    for x in range(cols):
                        newpage[y+64, x+64] = crop[y, x]
                # 대비(Contrast)를 준다
                cont = np.clip((1+alpha)*newpage - 128*alpha, 0, 255).astype(np.uint8)
                # 2차 인식
                newtext = pytesseract.image_to_string(cont, lang="eng", config=self.tesseract_config).strip().replace("O","0").replace("\n","")
                if newtext[1] == "S":
                    newtext = newtext[:1] + "5" + newtext[2:]
                newmatch = re.search(r'R[S5][A-Z0-9]{9}', newtext)
                if newmatch:
                    newtext = newmatch.group(0)
                different = (text != newtext) and "불일치" or ""

                # 출력
                print(newtext, different)
                self.output_file.write(f"{file},{text},{newtext},{different}\n")
                if different:
                    # 1차와 2차 결과가 불일치한 경우 확인용 그림파일 저장
                    self.imwrite(cont, file)
                sn_candidate.append(newtext)
        if len(sn_candidate) < 1:
            # 아예 인식이 안된 경우도 확인용 그림파일 저장
            print("미인식 - 확인 필요")
            self.output_file.write(f"{file},미인식,,확인필요\n")
            self.imwrite(gray, file)
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
    if sys.argv == 1:
        print("SnScanner [-s] [directory]")
    else:
        for i in range(1, len(sys.argv)):
            if sys.argv[i] == "-s":
                sn_only = True
            else:
                image_dir = sys.argv[i]
        ts = TextScanner(image_dir, sn_only=sn_only)
        ts.scan()
