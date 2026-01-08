# -*- coding: utf-8 -*-
"""
Created on Thu Jan  8 10:37:51 2026

@author: kyuun
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Jan  7 14:21:12 2026

@author: kyuun
"""

import webbrowser
import os
import pandas as pd
import shutil
import datetime
import time
import pubchempy as pcp
import math
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

print("廃液伝票用csvファイルを作成します　インターネットに接続してください")
print("廃液管理フォームの回答へアクセスしてダウンロードます...")
time.sleep(1)

# URLを読んでデフォルトのブラウザで表示→DL フォーム回答とプリセット試薬←いつかやる 
webbrowser.open("https://docs.google.com/spreadsheets/d/1O45H6QQWB1CeCBSK8tUwfkXG5fpDKs1Tcy9y7WJ3DwU/export?format=xlsx")

user_folder = os.path.expanduser("~")
DL = os.path.join(user_folder, "Downloads")

# DLされたら即次の処理をするべくtry文を使用
while True:
    try:
        shutil.move(DL +"\廃液管理（回答）.xlsx", "./廃液管理（回答）.xlsx")
        dt_now = datetime.datetime.now()
        rawfilename = '廃液管理回答_'+ dt_now.strftime('%m%d_%H%M') + '.xlsx'
        os.rename('廃液管理（回答）.xlsx', rawfilename)
    except:
        time.sleep(0.01)
    else:
        break
print("Download done")

df_ans = pd.read_excel(rawfilename, keep_default_na = True)

# 質問や選択肢を置換 以後これを使用
tank = "排出先タンク"
chem1 = "投入した試薬・物質1(水以外) / The reagent, substance disposed 1, except for water"
way1 = "物質1の量 記録方法 次のページで入力します / In which form to record the amount of disposals. It will be recorded in the next page."
weight1 = "溶質1の重量 [ g ] / Weight [ g ]"
flag1w = "廃液の溶質はこの1種のみですか? / Is that all for your disposal?"
conc1 = "濃度1 [ %, ppm, mol/L ] / Concentration [ %, ppm, mol/L ]"
vol1 = "体積1 [ mL ] / Volume [ mL ]"
flag1c = "廃液の内容物はこの1種のみですか? / Is that all for your disposal?"

chem2 = "投入した試薬・物質2(水以外) / The reagent, substance disposed 2, except for water"
way2 = "物質2の量 記録方法  / In which form to record the amount of disposals"
weight2 = "溶質2の重量 [ g ]  / Weight [ g ]"
flag2w = "廃液の溶質はこの2種のみですか? / Is that all for your disposal?"
conc2 = "濃度2 [ %, ppm, mol/L ] / Concentration [ %, ppm, mol/L ]"
vol2 = "体積2 [ mL ] / Volume [ mL ]"
flag2c = "廃液の内容物はこの2種のみですか? / Is that all for your disposal?"

chem3 = "投入した試薬・物質3(水以外) / The reagent, substance disposed 3, except for water"
way3 = "物質3の量 記録方法 / In which form to record the amount of disposals"
weight3 = "溶質3の重量 [ g ]  / Weight [ g ]"
flag3w = "廃液の溶質はこの3種のみですか? / Is that all for your disposal?"
conc3 = "濃度3 [ %, ppm, mol/L ] / Concentration [ %, ppm, mol/L ]"
vol3 = "体積3 [ mL ] / Volume [ mL ]"
flag3c = "廃液の内容物はこの3種のみですか? / Is that all for your disposal?"

chem4 = "投入した試薬・物質4(水以外) / The reagent, substance disposed 4, except for water"
way4 = "物質4の量 記録方法 / In which form to record the amount of disposals"
weight4 = "溶質4の重量 [ g ]  / Weight [ g ]"
flag4w = "廃液の溶質はこの4種のみですか? / Is that all for your disposal?"
conc4 = "濃度4 [ %, ppm, mol/L ] / Concentration [ %, ppm, mol/L ]"
vol4 = "体積4 [ mL ] / Volume [ mL ]"
flag4c = "廃液の内容物はこの4種のみですか? / Is that all for your disposal?"

other = "溶質の重量, もしくは濃度と体積 / Weight [ g ] or concentration [ %, ppm ] or [ mol/L ] and volume [ mL ]"

q_list = [tank, chem1, way1, weight1, flag1w, conc1, vol1, flag1c, chem2, way2, weight2, flag2w, conc2, vol2, flag2c, chem3, way3, weight3, flag3w, conc3, vol3, flag3c, chem4, way4, weight4, flag4w, conc4, vol4, flag4c, other]

g = "溶質の重量 [ g ] で記録 / Record with weight [ g ]"
pct = "溶質の濃度 [ % ] と溶液の体積 [ mL ] で記録 / Record with concentration [ % ] and volume [ mL ]"
ppm = "溶質の濃度 [ ppm ] と溶液の体積 [ mL ] で記録 / Record with concentration [ ppm ] and volume [ mL ]"
mol = "溶質のモル濃度 [ mol/L ] と溶液の体積 [ mL ] で記録 / Record with molarity [ mol/L ] and volume [ mL ]"

Y = "はい / Yes"
N = "いいえ / No"

def delete_rows(tank_ID):
    #認証情報の設定
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    #サービスアカウント作成, 秘密鍵生成!
    creds = ServiceAccountCredentials.from_json_keyfile_name('./data/calcium-task-481709-t1-5c1f9c5b1842.json', scope)
    # クライアントの作成
    client = gspread.authorize(creds)
    spreadsheet = client.open('廃液管理（回答）')
    worksheet = spreadsheet.sheet1
    data = worksheet.get_all_values()
    #data[n][4]が投入イベントの廃液タンク番号
    i = 0
    while True:
        if i == len(data):
            break
        else:
            pass
        if data[i][4] == tank_ID:
            worksheet.delete_rows(i+1) #1はじまり!!
            data = worksheet.get_all_values()
        else:
            i += 1

tankIDlist = []

# 過去に入力した内容を呼び出してリストに格納
molarity_subs = [] # 入ってたascii名と入力した日本語名を一緒に記録→ファイル読み出し, 最後に保存
CASnum = []
nameEN = [] # 入ってたascii名と入力した日本語名を一緒に記録→ファイル読み出し, 最後に保存
nameJP = []
molsub = open("./data/molarity_subs.txt", 'r', encoding = 'UTF-8')
for string in molsub:
    molarity_subs.append(string.rstrip('\n'))
CASnumbers = open('./data/CASnum.txt', 'r', encoding = 'UTF-8')
for string in CASnumbers:
    CASnum.append(string.rstrip('\n'))
EN = open('./data/nameEN.txt', 'r', encoding = 'UTF-8')
for string in EN:
    nameEN.append(string.rstrip('\n'))
JP = open('./data/nameJP.txt', 'r', encoding = 'UTF-8')
for string in JP:
    nameJP.append(string.rstrip('\n'))
molsub.close()
CASnumbers.close()
EN.close()
JP.close()

# 追記モードで予め開いておく 最後に閉じる
molsub = open("./data/molarity_subs.txt", 'a', encoding = 'UTF-8')
CASnumbers = open('./data/CASnum.txt', 'a', encoding = 'UTF-8')
EN = open('./data/nameEN.txt', 'a', encoding = 'UTF-8')
JP = open('./data/nameJP.txt', 'a', encoding = 'UTF-8')

while True:
    tank_disposing = input("伝票を作成するタンクを記号で指定 なければEnter: ")
    if tank_disposing == "K学" or (tank_disposing.isascii() and len(tank_disposing) > 1):
        tankIDlist.append(tank_disposing)
        event = [i for i, tank_num in enumerate(df_ans[tank]) if tank_num == tank_disposing]
        sub_list = []
        for i in event:
            for j in range(4):
                sub = []
                sub.append(df_ans[q_list[7*j+1]][i])
                if df_ans[q_list[7*j+2]][i] == g:
                    sub.append("gram")
                    sub.append(df_ans[q_list[7*j+3]][i])
                    sub.append("")
                    sub.append("")
                elif df_ans[q_list[7*j+2]][i] == pct:
                    sub.append("pct")
                    sub.append(df_ans[q_list[7*j+5]][i])
                    sub.append(df_ans[q_list[7*j+6]][i])
                    sub.append("")
                elif df_ans[q_list[7*j+2]][i] == ppm:
                    sub.append("ppm")
                    sub.append(df_ans[q_list[7*j+5]][i])
                    sub.append(df_ans[q_list[7*j+6]][i])
                    sub.append("")
                elif df_ans[q_list[7*j+2]][i] == mol:
                    # モル濃度の場合→CAS番を入力させて格納
                    sub.append("mol")
                    sub.append(df_ans[q_list[7*j+5]][i])
                    sub.append(df_ans[q_list[7*j+6]][i])
                    if df_ans[q_list[7*j+1]][i] in molarity_subs:
                        CAS = CASnum[molarity_subs.index(df_ans[q_list[7*j+1]][i])]
                    else:
                        while True:
                            CAS = input(str(df_ans[q_list[7*j+1]][i]) + " のCAS番号? : ")
                            if re.fullmatch(r'\d+-\d+-\d+', CAS) is None:
                                continue
                            else:
                                molarity_subs.append(df_ans[q_list[7*j+1]][i])
                                CASnum.append(CAS) # 以上リストに記載
                                molsub.write(df_ans[q_list[7*j+1]][i]+"\n")
                                CASnumbers.write(CAS + "\n") # 以上ファイルに追記
                                break
                    sub.append(CAS)
                sub_list.append(sub)
                if df_ans[q_list[7*j+4]][i] == Y or df_ans[q_list[7*j+7]][i] == Y:
                    break
            if type(df_ans[q_list[29]][i]) == str:
                while True:
                    sub = []
                    while True:
                        name = input(" " + df_ans[q_list[29]][i] + " 物質名を順に入力, 終わりの場合は Enter: ")
                        if (not name.isascii()) or name == "":
                            break
                    if name == "":
                        break
                    sub.append(name)
                    print("記録方法を入力, 重量: w, 重量%濃度: c, ppm: p, モル濃度: m")
                    while True:
                        way = input("")
                        if way == 'w' or way == 'c' or way == 'p' or way == 'm':
                            break
                        else:
                            print("入力が不適です! 重量: w, 重量%濃度: c, ppm: p, モル濃度: m")
                    if way == "w":
                        sub.append("gram")
                        a = float(input("重量 [g] を値のみ入力: "))
                        sub.append(a)
                    elif way == "c":
                        sub.append("pct")
                        a = float(input("濃度 [%] を値のみ入力: "))
                        v = float(input("体積 [mL] を値のみ入力: "))
                        sub.append(a)
                        sub.append(v)
                    elif way == "p":
                        sub.append("ppm")
                        a = float(input("濃度 [ppm] を値のみ入力: "))
                        v = float(input("体積 [mL] を値のみ入力: "))
                        sub.append(a)
                        sub.append(v)
                    elif way == "m":
                        sub.append("mol")
                        a = float(input("モル濃度 [mol/L] を値のみ入力: "))
                        v = float(input("体積 [mL] を値のみ入力: "))
                        sub.append(a)
                        sub.append(v)
                        while True:
                            CAS = input(name + " のCAS番号? : ")
                            if re.fullmatch(r'\d+-\d+-\d+', CAS) is None:
                                continue
                            else:
                                break
                        sub.append(CAS)
                    sub_list.append(sub)
        df_disp = pd.DataFrame(data = sub_list, columns = ["subs", "way", "conc_gram", "vol", "CAS"])

        for i in range(len(df_disp)):
            # この順番じゃないと先に "mg" の中の "g" が置換されてオワる
            if "ug" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("ug", ''))/1000000
            elif "μg" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("μg", ''))/1000000
            elif "mg" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("mg", ''))/1000
            elif "g" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("g", ''))
            elif "mol/L" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("mol/L", ''))
            elif "mM" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("mM", ''))/1000
            elif "uM" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("uM", ''))/1000000
            elif "M" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("M", ''))
            elif "ppm" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("ppm", ''))/1000000
            elif "%" in str(df_disp["conc_gram"][i]):
                df_disp.loc[i, "conc_gram"] = float(str(df_disp["conc_gram"][i]).replace("%", ''))
            elif type(df_disp["conc_gram"][i]) == str:
                print("数値を認識できませんでした")
                df_disp.loc[i, "conc_gram"] = float(input(df_disp["conc_gram"][i] + " 数値を入力: "))

        for i in range(len(df_disp)):
            if "mL" in str(df_disp["vol"][i]):
                df_disp.loc[i, "vol"] = float(str(df_disp["vol"][i]).replace("mL", ''))
            if "ml" in str(df_disp["vol"][i]):
                df_disp.loc[i, "vol"] = float(str(df_disp["vol"][i]).replace("ml", ''))
            elif "uL" in str(df_disp["vol"][i]):
                df_disp.loc[i, "vol"] = float(str(df_disp["vol"][i]).replace("uL", ''))/1000
            elif "μL" in str(df_disp["vol"][i]):
                df_disp.loc[i, "vol"] = float(str(df_disp["vol"][i]).replace("μL", ''))/1000
            elif "L" in str(df_disp["vol"][i]):
                df_disp.loc[i, "vol"] = float(str(df_disp["vol"][i]).replace("L", ''))*1000 # 260107 リットル表記に対応
        print("UTCIMSで読み込むためのcsvファイルを作成します...")
        export = open(DL+"\\"+ tank_disposing + "_export.csv", 'w', encoding = 'UTF-8-sig') # ここ UTF-8 だとうまく走らなかった
        export.write("SourceType,CASRN,Substance,Mass\n")
        for i in range(len(df_disp)):
            if df_disp["subs"][i].isascii():
                if df_disp["subs"][i] in nameEN:
                    JPname = nameJP[nameEN.index(df_disp["subs"][i])]
                else:
                    nameEN.append(df_disp["subs"][i])
                    JPname = input(df_disp["subs"][i] + " 日本語名を入力: ")
                    nameJP.append(JPname)
                    EN.write(df_disp["subs"][i]+'\n')
                    JP.write(JPname + '\n')
                df_disp.loc[i, "subs"] = JPname
            if df_disp["way"][i] == "mol":
                try:
                    MW = float(pcp.get_properties('MolecularWeight', df_disp["CAS"][i], "name")[0]['MolecularWeight']) #CAS番検索に失敗した場合のtry文
                except:
                    print("CAS番号の検索に失敗しました")
                    MW = float(input(df_disp["subs"][i] + " の分子量を入力: "))
                    mass = df_disp["conc_gram"][i]*df_disp["vol"][i]*MW/1000
                    export.write("FreeText,," + df_disp["subs"][i] + "," + str(mass) + "\n")
                else:
                    mass = df_disp["conc_gram"][i] * df_disp["vol"][i] * MW/1000 # モル濃度×体積mL×分子量/1000
                    export.write("SubstanceMaster," + df_disp["CAS"][i] + ",," + str(mass) + "\n")
            elif df_disp["way"][i] == "gram":
                export.write("FreeText,," + df_disp["subs"][i] + "," + str(df_disp["conc_gram"][i]) + "\n")
            elif df_disp["way"][i] == "pct":
                mass = df_disp["conc_gram"][i]*df_disp["vol"][i]
                export.write("FreeText,," + df_disp["subs"][i] + "," + str(mass) + "\n")
            elif df_disp["way"][i] == "ppm":
                mass = df_disp["conc_gram"][i]*df_disp["vol"][i]
                export.write("FreeText,," + df_disp["subs"][i] + "," + str(mass) + "\n")
        export.close()
        water = math.ceil(9000 - sum(pd.read_csv(tank_disposing + "_export.csv")['Mass']))
        export = open(tank_disposing + "_export.csv", 'a', encoding = 'UTF-8-sig')
        if "A" in tank_disposing or "C" in tank_disposing or "J" in tank_disposing or "H" in tank_disposing or "D" in tank_disposing or "E" in tank_disposing:
            export.write("FreeText,,水," + str(water))
        else:
            export.write("FreeText,,水," + str(water + 9000))
        export.close()
        print("作成完了しました! ダウンロードディレクトリをご確認ください!")
        print()
    else:
        print("回答シート上のデータを削除します") # tankIDlistが空の場合の処理→そのうち
        for ID in tankIDlist:
            print("タンク " + ID + " のデータを削除しています......")
            delete_rows(ID)
        break
molsub.close()
CASnumbers.close()
EN.close()
JP.close()
print("集計してcsvへ記載した投入イベントを削除しました。")
print("削除されたイベントを確認するにはダウンロードしたファイルを参照してください。")
time.sleep(5)