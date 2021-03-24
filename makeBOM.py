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

import pandas as pd
import sys
from pathlib import Path

s_argv_fail_message=''
s_dic_filename = 'parts_dic_ORG.txt'
p_dic_path = Path.home() / 'Box' / 'PCB-CAD' / 'makebom' 
p_dic_file = p_dic_path / s_dic_filename
# windowsとの互換性のため、必要なファイルへのパスをpathlibで加工するようにした
# ファイル名は適宜変更のこと


# print(len(sys.argv))
if len(sys.argv)>1 : # 引数がなければエラー
    p_argv1=Path(sys.argv[1])
    # print(p_dic_path)
    if p_argv1.is_file() : # 引数のファイルの実体がなければエラー
        # print(p_argv1.resolve())
        # print(p_argv1.parent.resolve())
        s_output_filename = str(p_argv1.stem)
        s_output_path = str(p_argv1.parent.resolve())+'/'
    else: 
        s_argv_fail_message = "Specified files did not exists or not a file."
else:
    s_argv_fail_message = 'No input files specified.'

if len(s_argv_fail_message)!=0 :
    print(s_argv_fail_message)
    exit()

# print('Input file =', str(p_argv1))
# print('Output Dir =', s_output_path)
# print('Output File Name', s_output_filename)
# print('Dic File Name', str(p_dic_file))
# print(p_dic_file.is_file() )
    

# internal variables settings

columun_list_kousei_CSV = ['Category', 'Partnumber','Mfr',  'Count', 'UnitPrice']
columun_list_PBAN_CSV = ['Mfr','Category', 'Partnumber', 'PartName', 'Count', 'NumTerminals']

# 部品表読み込み
print('Reading file. -- ',end='')

df_partslist = pd.read_table(str(p_argv1), header = 4, delim_whitespace=True).loc[:,['Part','Value','Sheet']]
# 空行はカウントされないので四行目がヘッダ名の定義に相当する。
# スペース区切りのためdelim_whitespace=Trueを入れた
# ここでは、全ての行のPart,Value,Sheetのデータを抜き出す。
# Extract necessary part in the bom file and store into a DataFrame.

# !! 条件をつけてリストアップする
# df_partslist[df_partslist['Sheet']==1]
# とすると、シート１にある部品だけが出てくる。
# df_partslist[(df_partslist['Sheet']!=3)|(df_partslist['Part']=='C42')|(df_partslist['Part']=='C43')]
# 複数条件の時はこのようにする。　sheet3以外、もしくは　C42 C43をリストアップする

# 組み配部品表・PBAN部品表作成のため    部品名でグループ化しDataFrameに入れる
# grouped_by_value.groups[item]とすると、itemで示された名前のグループを構成するindex番号のリストが得られる
grouped_by_value = df_partslist.groupby('Value')

d_kumihai_list={}
for item in df_partslist['Value'].unique():
    # 同じ種類の部品ごとに部品番号をピックアップするためのリスト 
    l_part_names=[]
    for index in grouped_by_value.groups[item]:
        # グループ内（同じ部品）の部番をひとつずつピックアップ
        l_part_names.append(df_partslist.loc[index]['Part'])

    d_kumihai_list[item]=l_part_names

s_kumihai_list = pd.Series(d_kumihai_list)

# print(s_kumihai_list)

print('Reading MI. -- ',end='')
# MIの読み込み 必要なrowだけ読み込む コラム名を設定
mi=pd.read_table(str(p_dic_file), header=None).loc[:,[0,1,2,3,9,14]]
mi.columns=['Category','Partnumber','Mfr','CADname','NumTerminals','UnitPrice']

df_kumihai_partslist = pd.DataFrame({'Mfr':[], 'Category':[], 'Partnumber':[], 'PartName':[], 'Count':[], 'NumTerminals':[], 'UnitPrice':[]})
# appendするために基本の空データフレームを作る

print('Looking up MI. -- ')
parts = []
for item in df_partslist['Value'].unique(): # 部品表のValue欄からユニークなリストを作ってiteration
    
    # miから部品表のCADnameの値で検索する
    if len( mi[mi['CADname']==item] )>0:
        # マッチが複数ある場合は最初を取る(iloc[0,:])
        df_kumihai_partslist.loc[item]=mi[mi['CADname']==item].iloc[0,:]
    else:
        df_kumihai_partslist.at[item,'Category']='NO-MI'
        print("No MI Found on: Value =",item,'/ PartNumber =',df_partslist[df_partslist['Value']==item]['Part'].tolist())
        # マッチがない場合はエラーを表示

    df_kumihai_partslist.at[item,'Count']=len(s_kumihai_list[item])
    #部品使用数をカウント and store into row:Count
    parts.append(', '.join(s_kumihai_list[item]))
    # 部品番号をリストから文字列に変換

df_kumihai_partslist['PartName']=parts
# df_kumihai_partslist['PartName']=s_kumihai_list
df_kumihai_partslist = df_kumihai_partslist.astype({'Count':int})

df_kumihai_partslist.loc[:,columun_list_PBAN_CSV].to_csv(s_output_path+s_output_filename+'_PBAN.csv', encoding='utf_8_sig', header=False, index=False)
df_kumihai_partslist.loc[:,columun_list_kousei_CSV].to_csv(s_output_path+s_output_filename+'_kousei.csv',header=False, index=False)


