# coding: utf-8
import sys
import io
import os
import pandas as pd
import shutil

def MojibakeTaisaku():
    # 出力の文字化け対策
    sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def GetExcelData():
    # pyファイルのあるフォルダに移動する
    os.chdir(os.path.dirname(os.path.abspath(__file__)))    

    # xlsファイルを開いて配列dfに格納する
    df = pd.read_excel("test.xlsx", index_col=0)
    idx = df.index.values               # 1列目が項目名として格納されている
    clm = df.columns.values             # 1行目が要素名として格納されている
    r,c= df.shape                       # r:行数、c:列数を取得

    # print(df.at['入力電流', '項目'])     # at[index名,columns名]のクロス位置の値を返す
    # print(df.iat[0,0])                 # at[index番号,columns番号]のクロス位置の値を返す    [0,0]が左上。index列の右から開始

    # print('要素は',idx)
    # print('タイトル行は',clm)
    # print(r,c)

    return df, idx, clm, r, c   #df:セル配列、idx:項目名、clm:要素名、r:行数、c:列数を返す

def ReplaceDxfStrings(df,r,c,dxfName_Ref,dxfName_New):
    enc="utf-8" #"shift_jis"

    # pyファイルのあるフォルダに移動する
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # ファイル名付与
    shutil.copy(dxfName_Ref, dxfName_New)

    with open(dxfName_New, encoding=enc) as f:
        data_lines = f.read()

    # 文字列置換
    for i in range(0,r):
        TargetStr = str(df.iat[i,0])
        ReplaceStr = str(df.iat[i,c])
        # print(i)
        # if TargetStr == 'mDk_VA':
        #     print('TargetStr:',TargetStr,type(TargetStr))
        #     print('ReplaceStr:',ReplaceStr,type(ReplaceStr))
        
        # 置換処理
        data_lines = data_lines.replace(TargetStr, ReplaceStr)

    # 同じファイル名で保存
    with open(dxfName_New, mode="w", encoding=enc) as f:
        f.write(data_lines)

if __name__ == '__main__':
    MojibakeTaisaku()
    df, idx, clm, r, c = GetExcelData()

    dxfName_Ref = 'NewSlim_1.dxf'
    c_start = 2
    for j in range(c_start,c):
        print('TargetColumnNo:',j,clm[j])
        dxfName_New = clm[j] +'.dxf'
        ReplaceDxfStrings(df,r,j,dxfName_Ref,dxfName_New)