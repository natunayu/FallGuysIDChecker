import sys
import PIL.Image
from google.cloud import vision
import os
from PIL import Image
import pyocr
import numpy as np
import io
import cv2
import pyautogui
from time import sleep
from win32gui import GetWindowText, GetForegroundWindow




#tesseractのデータ参照とVisionAPiの秘密鍵へのパス
path=''
os.environ["GOOGLE_APPLICATION_CREDENTIALS"]=""


#pyocrへOCRデータを参照
os.environ['PATH'] = os.environ['PATH'] + path
pyocr.tesseract.TESSERACT_CMD = path + 'tesseract.exe'
tools = pyocr.get_available_tools()
tool = tools[0]


round_count = 1 #ラウンドのカウントを初期化
player_list = [] #メンバーリストの初期化
player_score = [] #メンバーのスコアリストの初期化
flag = 100 #0=fall-guys未起動, 1=リザルト画面外, 2=リザルト画面 3=ファイナルラウンド終了(ゲームリセット) 5=アプリを終了
match_name = ""


#文字間の齟齬を計算して文字数の差を数値で排出(レーベンシュタイン距離)
def w_distance(a, b):
    l1, l2 = len(a), len(b)
    dp1, dp2 = [i for i in range(l1 + 1)], [0] * (l1 + 1)
    for i in range(1, l2 + 1):
        dp2[0] = i
        for j in range(1, l1 + 1):
            if a[j - 1] == b[i - 1]:
                dp2[j] = dp1[j - 1]
            else:
                dp2[j] = min(dp1[j - 1] + 1, min(dp1[j] + 1, dp2[j - 1] + 1))
        tmp = dp1
        dp1 = dp2
        dp2 = tmp
    return dp1[l1]


#numpy配列のスライス
def trim(array: list, x: int, y: int, width: int, height: int):
    return array[y:y + height, x:x + width]


#csvへの出力書き込み
def write_csv(p_list, p_score):
    global match_name
    global round_count
    global flag
    s=''
    print("マッチ結果をファイルに出力しています…")
    for i in range(len(p_list)):
        s = s + str(p_list[i]) + ", " + str(p_score[i]) + "\n"
    before = s.encode('cp932', "ignore")
    after = before.decode('cp932')
    with open("./result/" + match_name + "_log.csv", mode='w') as f:
        f.write(after)

    print("出力が完了しました！")
    flag = 5


#PIL->CV2
def pil2cv(image: PIL.Image.Image):
    new_image = np.array(image, dtype=np.uint8)
    if new_image.ndim == 2:  # モノクロ
        pass
    elif new_image.shape[2] == 3:  # カラー
        new_image = cv2.cvtColor(new_image, cv2.COLOR_RGB2BGR)
    elif new_image.shape[2] == 4:  # 透過
        new_image = cv2.cvtColor(new_image, cv2.COLOR_RGBA2BGRA)
    return new_image


#CV2->PIL
def cv2pil(image):
    new_image = image.copy()
    if new_image.ndim == 2:  # モノクロ
        pass
    elif new_image.shape[2] == 3:  # カラー
        new_image = cv2.cvtColor(new_image, cv2.COLOR_BGR2RGB)
    elif new_image.shape[2] == 4:  # 透過
        new_image = cv2.cvtColor(new_image, cv2.COLOR_BGRA2RGBA)
    new_image = Image.fromarray(new_image)
    return new_image


#CV2Image縦連結
def vconcat_resize_min(im_list, interpolation=cv2.INTER_CUBIC):
    w_min = min(im.shape[1] for im in im_list)
    im_list_resize = [cv2.resize(im, (w_min, int(im.shape[0] * w_min / im.shape[1])), interpolation=interpolation)
                      for im in im_list]
    return cv2.vconcat(im_list_resize)


#OCR用にフォールガイズのリザルトスクリーンショットを返す
def get_result_image(result_image):
    #スクリーンショットに置き換え
    img_array = np.array(result_image)

    x = 398
    y = 183
    width = 1124
    height = 500

    im_trim2 = trim(img_array, x, y, width, height)

    im_trim3 = trim(im_trim2, int(width / 12 * 0), 0, int(width / 12), 500)
    imt1 = pil2cv(Image.fromarray(im_trim3))

    #セルのIDをfor文とOCRで検証 セルマス12x5
    for i in range(11):
        im_trim3 = trim(im_trim2, int(width / 12 * (i + 1)), 0, int(width / 12), 500)
        material_img = pil2cv(Image.fromarray(im_trim3))

        imt1 = vconcat_resize_min([imt1, material_img])

    return cv2pil(imt1)


#VisionAPIでプレイヤーIDの検出 -> list
def player_determining(result_image):

    resized_img = result_image
    img_bytes = io.BytesIO()
    resized_img.save(img_bytes, format='PNG')
    img_bytes = img_bytes.getvalue()

    try:
        client = vision.ImageAnnotatorClient()
        image = vision.Image(content=img_bytes)
        response = client.text_detection(image=image)

        stack1 = str(response.text_annotations)[28:]
        for w in range(len(stack1)):
            if stack1[w] == '"':
                stack2 = stack1[:w]
                return stack2.replace('「', '').replace('」', '').replace('―', '').replace('一', '').split('\\n')
                break
            else:
                pass
    except Exception as e:
        print("VisionAPIへの接続エラー:", e)


#indexを返す 無い場合-1
def index_check(l, x):
    return l.index(x) if x in l else -1


#クリアプレイヤーのリストを引数にスコアを集計
def score_calc(clear_player_list):
    global flag
    global round_count
    global player_score
    global player_list
    print("スコアを集計しています…")

    #1ラウンド終了時
    if round_count == 1:
        player_list = clear_player_list.copy()
        for i in range(len(player_list)):
            player_score.append(1)

    #最終ラウンド終了時
    elif flag == 3:
        for player in clear_player_list:
            index = index_check(player_list, player)

            #見つからなかった場合の処理(最も距離が短い要素を正として判定)
            if index < 0:
                w_d_min = 100
                for i in range(len(player_list)):
                    w_d = w_distance(player_list[i], player)
                    if w_d < w_d_min:
                        w_d_min = w_d
                        index = i
            player_score[index] += 10

        write_csv(player_list, player_score)

    #その他のマッチ終了時
    else:
        for player in clear_player_list:
            index = index_check(player_list, player)

            #見つからなかった場合の処理(最もレーベンシュタイン距離が短い要素を正として判定)
            if index < 0:
                w_d_min = 100
                for i in range(len(player_list)):
                    w_d = w_distance(player_list[i], player)
                    if w_d < w_d_min:
                        w_d_min = w_d
                        index = i
            player_score[index] += round_count + round_count

    print("スコアの集計を完了しました。")


#現在のフレームを判別
def check_frame():
    global flag
    global round_count

    #アクティブウィンドウの取得
    active_frame = GetWindowText(GetForegroundWindow())
    if active_frame != "FallGuys_client":
        if flag == 0:
            return 0
        else:
            print("0: FallGuysを起動してください")
            return 0

    #スクリーンショットをnparrayに変換
    img_array = np.array(pyautogui.screenshot(region=(0, 0, 1920, 1080)))

    #result画面のクリアの文字をトリミング
    im_trim1 = trim(img_array, 550, 40, 810, 140)
    #優勝者が排出されたかどうかのトリミング
    im_trim2 = trim(img_array, 10, 10, 500, 125)

    im_trim1 = Image.fromarray(im_trim1)
    im_trim2 = Image.fromarray(im_trim2)

    # 画像の文字を抽出
    builder = pyocr.builders.TextBuilder(tesseract_layout=6)
    text1 = tool.image_to_string(im_trim1, lang="jpn", builder=builder)
    text2 = tool.image_to_string(im_trim2, lang="jpn", builder=builder)

    #リザルト画面の判定
    if 'クリア' in text1:
        if flag == 2:
            return 2
        else:
            print('2: リザルト中')
            r_img = Image.fromarray(img_array)

            result_img = get_result_image(r_img)
            result_list = player_determining(result_img)
            score_calc(result_list)

            #ラウンドの加算
            round_count += 1
            print(result_list)
            return 2

    #ファイナルラウンドの判定
    elif 'WINNER' in text2:
        if flag == 3:
            return 3
        else:
            print('3: ファイナルラウンド終了')
            flag = 3
            im_trim3 = trim(img_array, 850, 925, 300, 30)
            im_trim3 = Image.fromarray(im_trim3)
            champion_id = player_determining(im_trim3)

            score_calc(champion_id)
            print("champion_id")
            return 3

    else:
        if flag == 1:
            return 1
        else:
            print('1: FallGuysをプレイ中')
            return 1



if __name__ == "__main__":
    match_name = str(input("ゲームマッチの名称を入力してください(例: 第1ゲーム) >"))
    print(match_name)
    try:
        while True:
            flag = check_frame()
            sleep(0.5)
            if flag == 5:
                print('マッチが終わった為アプリを終了します。')
                sys.exit()

    #Ctrl+Cでアプリの強制終了。
    except KeyboardInterrupt:
        print('アプリを終了します。')








