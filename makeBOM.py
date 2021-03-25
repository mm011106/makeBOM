#!/usr/bin/python3
# -*- coding: utf-8 -*-
# 
# Eagle CAD の部品表出力を在庫部品リストと照合して、組配用部品表、構成部品表のリストを作る
#       v1.0    2021/3  miyamoto

#   出力ファイルはCSV, UTF-8, CR/LF
#       組配部品表はUTF-8(BOM)  ->エクセルで読める
#   辞書ファイル（在庫部品）はBoxの上ある。

# 実行は：
#   $ python3 makeBOM.py {PATHtoTargetFile}/EagleBOMファイル
# 引数：
#   1:EagleBOMファイルの指定（必須）
#   2:取り出すシート番号の指定  シート番号（１枚だけ指定可能）  0:全て
#       シート番号を指定した場合、指定したシートとそれ以外のシートの２つに分けて部品表が作成される
#       ファイル名に _数字 がついているのが指定したシートの部品表、
#       _R がついているのがそれ以外。


import pandas as pd
import sys
from pathlib import Path


# 定数、関数の定義
s_dic_filename = 'parts_dic_ORG.txt'
p_dic_path = Path.home() / 'Box' / 'PCB-CAD' / 'makebom' 
p_dic_file = p_dic_path / s_dic_filename
# windowsとの互換性のため、必要なファイルへのパスをpathlibで加工するようにした
# ファイル名は適宜変更のこと

def make_kumihai_partslist(df_partslist):
    global mi

    # 同じ部品を何個、どこで使っているか調べるため  部品名でグループ化しDataFrameに入れる
    # grouped_by_value.groups[item]とすると、itemで示された名前のグループを構成するindex番号のリストが得られる
    df_grouped_by_value = df_partslist.groupby('Value')
    # print(df_grouped_by_value.groups)
    
    df_kumihai_partslist = pd.DataFrame({'Mfr':[], 'Category':[], 'Partnumber':[], 'PartName':'', 'Count':[], 'NumTerminals':[], 'UnitPrice':[]})
    # appendするために基本の空データフレームを作る
    # print(df_kumihai_partslist.dtypes)
    
    for item in df_partslist['Value'].unique(): # 部品表のValue欄からユニークなリストを作ってiteration
        
        # miから部品表のCADnameの値で検索する
        if len( mi[mi['CADname']==item] )>0:
            # マッチがある場合      複数ある場合は最初を取る(iloc[0,:])
            df_kumihai_partslist.loc[item]=mi[mi['CADname']==item].iloc[0,:]
        else:
            # マッチがない場合はエラーを表示
            df_kumihai_partslist.at[item,'Category']='NO-MI'
            print("No MI Found on: Value =",item,'/ PartNumber =',df_partslist[df_partslist['Value']==item]['Part'].tolist())

        # 同じ部品が使われている部品番号をリストとして取得
        partsOnItem = df_partslist.loc[df_grouped_by_value.groups[item],'Part'].tolist()
        #部品使用数をカウントして'Count'rowに入れる
        df_kumihai_partslist.at[item,'Count']=len(partsOnItem)
        # 部品番号をリストから文字列に変換して'PartName'rowに入れる
        df_kumihai_partslist.at[item,'PartName']=', '.join(partsOnItem)

    df_kumihai_partslist = df_kumihai_partslist.astype({'Count':int})
    # print(df_kumihai_partslist.dtypes)

    return df_kumihai_partslist



s_argv_fail_message=''
if len(sys.argv)>1 : # 引数がなければエラー
    p_argv1=Path(sys.argv[1])
    sheet_extract = 0
    if p_argv1.is_file() : # 引数のファイルの実体がなければエラー
        # print(p_argv1.resolve())
        # print(p_argv1.parent.resolve())
        s_output_filename = str(p_argv1.stem)
        s_output_path = str(p_argv1.parent.resolve())+'/'
    else: 
        s_argv_fail_message = "Specified file did not exist or not a file."
    
    # 引数が２つある場合は２つ目をシート番号として取り込む
    if len(sys.argv)>2:
        sheet_extract = sys.argv[2]

else:
    s_argv_fail_message = 'No input files specified.'

# エラーがある場合は終了
if len(s_argv_fail_message)!=0 :
    print(s_argv_fail_message)
    exit()


# 出力する部品表に必要な要素の指定
columun_list_kousei_CSV = ['Category', 'Partnumber', 'Mfr',  'Count', 'UnitPrice']
columun_list_PBAN_CSV = ['Mfr', 'Category', 'Partnumber', 'PartName', 'Count', 'NumTerminals']

# MI読み込み
print('Reading MI. -- ',end='')
# MIの読み込み 必要なrowだけ読み込む コラム名を設定
mi=pd.read_table(str(p_dic_file), header=None).loc[:,[0,1,2,3,9,14]]
mi.columns=['Category','Partnumber','Mfr','CADname','NumTerminals','UnitPrice']

# 部品表読み込み
print('Reading file. -- ',end='')
df_partslist = pd.read_table(str(p_argv1), header = 4, delim_whitespace=True).loc[:,['Part','Value','Sheet']]
# 空行はカウントされないので四行目がヘッダ名の定義に相当する。
# スペース区切りのためdelim_whitespace=Trueを入れた
# ここでは、全ての行のPart,Value,Sheetのデータを抜き出す。

if sheet_extract!=0:
    print('Separate Sheets. -- ',end='')
    print('Sheet:' + sheet_extract + ' / ',end='')
    s_export_filename = s_output_path + s_output_filename 
    s_postfix = '_' + str(sheet_extract)
    df_kumihai_partslist=make_kumihai_partslist(df_partslist[df_partslist['Sheet']==int(sheet_extract)])
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_PBAN_CSV].to_csv(s_export_filename + '_PBAN' + s_postfix + '.csv', encoding='utf_8_sig', header=False, index=False)
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_kousei_CSV].to_csv(s_export_filename + '_kousei' + s_postfix + '.csv', header=False, index=False)

    print('Sheet:others :',end='')
    s_export_filename = s_output_path + s_output_filename 
    s_postfix = '_R'
    df_kumihai_partslist=make_kumihai_partslist(df_partslist[df_partslist['Sheet']!=int(sheet_extract)])
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_PBAN_CSV].to_csv(s_export_filename + '_PBAN' + s_postfix + '.csv', encoding='utf_8_sig', header=False, index=False)
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_kousei_CSV].to_csv(s_export_filename + '_kousei' + s_postfix + '.csv',header=False, index=False)

else:
    print('All Sheets. -- ',end='')
    s_export_filename = s_output_path + s_output_filename 
    s_postfix = '_A'
    df_kumihai_partslist=make_kumihai_partslist(df_partslist)
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_PBAN_CSV].to_csv(s_export_filename + '_PBAN' + s_postfix + '.csv', encoding='utf_8_sig', header=False, index=False)
    df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_kousei_CSV].to_csv(s_export_filename + '_kousei' + s_postfix + '.csv', header=False, index=False)


print('  Fine.')