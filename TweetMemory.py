import tweepy
import tkinter
from tkinter.scrolledtext import ScrolledText
import datetime
import openpyxl

def TweetOnTwitter(tweet):
    API_KEY = ''
    API_SECRET = ''
    ACCESS_TOKEN = ''
    ACCESS_TOKEN_SECRET = ''

    # APIの認証(API v2)
    client = tweepy.Client(consumer_key=API_KEY, consumer_secret=API_SECRET, access_token=ACCESS_TOKEN, access_token_secret=ACCESS_TOKEN_SECRET)

    # ツイート
    client.create_tweet(text=tweet)
    print("Tweet successful.")

def SaveTweet(tweet):
    # 日時取得
    now = datetime.datetime.now()
    
    # ブック取得
    book = openpyxl.load_workbook('./tweets.xlsx')
    sheet = book["Sheet1"]
    # 書き込み行番号取得
    cell = sheet['A2']
    line = cell.value
    
    # ツイート記録
    sheet["B"+str(line)] = now      # 投稿日時保存
    sheet["C"+str(line)] = tweet    # ツイート内容保存
    sheet["C"+str(line)].alignment = openpyxl.styles.Alignment(wrapText=True)   # 改行含めてセル書き込み
    sheet['A2'] = line + 1          # 次回書き込み行番号情報更新

    # Excel更新
    book.save('./tweets.xlsx')
    print("Tweet save is successful.")

def NetaMemory():
    tweet = text.get("1.0", "end-1c")   # 第1引数はmust, 第2引数は全文取得の意味
    # ツイート投稿
    TweetOnTwitter(tweet)
    # ツイート内容保存
    SaveTweet(tweet)

# ツイート入力ウインドウの定義
root = tkinter.Tk()
root.geometry('400x120')
root.title("つぶやき")
text = ScrolledText(root, font=("", 12), height=5, width=50)    # 改行対応入力フォーム
text.pack()
button = tkinter.Button(root, text="ツイート", command=NetaMemory)
button.pack()
root.mainloop()