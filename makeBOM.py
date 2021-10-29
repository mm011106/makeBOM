#!/usr/bin/python3
# -*- coding: utf-8 -*-
# 
# KI-CAD の部品表出力を在庫部品リストと照合して、組配用部品表、構成部品表のリストを作る
#       v2.0    2021/10  miyamoto

#   出力ファイルはCSV, UTF-8, LF
#       組配部品表はUTF-8(BOM)  ->エクセルで読める
#   辞書ファイル（在庫部品）はBoxの上ある。

# 実行は：
#   $ python3 makeBOM.py {PATHtoTargetFile}/EagleBOMファイル
# 引数：
#   1:EagleBOMファイルの指定（必須）
# 出力：
#  {PATHtoTargetFile}の下に partslist というディレクトリが作成され、そこに出力される
#

from numpy import dtype
import pandas as pd
import sys
import os
from pathlib import Path

# 定数、関数の定義
s_dic_filename = 'FOR_KICAD_parts_dic_ORG.txt'
p_dic_path = Path.home() / 'Box' / 'PCB-CAD' / 'makebom' 
p_dic_file = p_dic_path / s_dic_filename
# windowsとの互換性のため、必要なファイルへのパスをpathlibで加工するようにした
# ファイル名は適宜変更のこと


#
#   設計部品表を作成する
#  引数：
#   1:ki-CADからのBOM（必須）
def make_sekkei_partslist(df_partslist):
    df_sekkei_partslist = pd.DataFrame({'Ref':[], 'Category':[], 'Partnumber':[],'Mfr':[]},dtype=object)
    # print(df_sekkei_partslist.dtypes)
    # dtypeで表のデータの型を前もって指定しておかないと、Floatになる。
    # 他のdataframeを代入するとその型に上書きされるけれど、要素として代入しようとすると型が合わないのでエラーになる。

    # 最初に設計部品表を作る    これを元にして組配部品表を組み直すという手順で進む

    # print(df_partslist.dtypes)
    print('Generating Sekkei Partslist....')
    for idx in df_partslist.index.values:
        # print(df_partslist.loc[idx,'Part'])
        # print(df_partslist.at[idx, 'Part'])
        # print(df_sekkei_partslist.loc[idx])
        
        cadName=df_partslist.at[idx,'Value']
        sizeCode=df_partslist.at[idx,'Size']
        solderingMethod=df_partslist.at[idx,'Package']

        # BOMの部品名でMIを検索
        matchedMI=mi[mi['CADname']==cadName]
        if len(matchedMI)==1:  # 部品名でマッチするMIが１つなら、それを採用
            # BOMの中にサイズの指定があるかどうか確認
            # サイズの指定があればサイズコードでMIを選別
            if (type(sizeCode) is str):
                # print("!! finding suitable device size of :",df_partslist.at[idx,'Size'])
                matchedMI=matchedMI[matchedMI['SizeCode']==sizeCode]
                # print(matchedMI)
                # print("  found size code in BOM")
                #
            if len(matchedMI)==1:
                # 最後まで残ったMIを設計部品表に登録
                df_sekkei_partslist.loc[idx]=matchedMI.iloc[0,:]
                # print("  .. fond MI matched with the BOM ")
        
            else:
                #   BOMのサイズコードに合うものがMIになければエラー
                df_sekkei_partslist.at[idx,'Category']='NO-suitable device'
                df_sekkei_partslist.at[idx,'Partnumber']='NO-suitable device'
                print("WARNING: No suitable device Found on: ",df_partslist.at[idx,'Part'],"  ",df_partslist.at[idx,'Value']," need to be a style of:",solderingMethod, "/",sizeCode)
            

        elif len(matchedMI)>1: # 部品名でマッチしたMIが１つ以上なら、さらに選別

            # print("SOLDERING:",solderingMethod," --  SIZE CODE:",sizeCode)
            # print(matchedMI[matchedMI['SolderingMethod']==df_partslist.at[idx,'Soldering']])
            # 部品名でマッチしたMIのリストからさらに、実装方法とサイズが回路のBOMとマッチするものを取り出す
            matchedMI=matchedMI[matchedMI['SolderingMethod']==solderingMethod]
            # print(matchedMI['SizeCode'],sizeCode)
            matchedMI=matchedMI[matchedMI['SizeCode']==sizeCode]
            
            if len(matchedMI)>0: # 実装方法でマッチしたもののMIが1つか、それ以上ならその最初の一つを選ぶ
                df_sekkei_partslist.loc[idx]=matchedMI.iloc[0,:]
                # print("-- FOUND: ", df_sekkei_partslist.loc[idx])
            else:
                # 実装方法でマッチしなかったらエラー
                # print(allCandidate)
                # print(sizeCode)
                df_sekkei_partslist.at[idx,'Category']='NO-suitable device'
                df_sekkei_partslist.at[idx,'Partnumber']='NO-suitable device'
                print("WARNING: No suitable device Found on: ",df_partslist.at[idx,'Part'],"  ",df_partslist.at[idx,'Value']," need to be a style of:",solderingMethod, "/",sizeCode)
            #   ここにマッチしたMIをさらに詳細に選別するコードを入れる
            
        else:
            # そもそもMIのなかにマッチがない場合はエラーを表示
            df_sekkei_partslist.at[idx,'Category']='NO-MI'
            df_sekkei_partslist.at[idx,'Partnumber']='NO-MI'
            print("WARNING: No MI Found on: Value =",df_partslist.at[idx,'Part']," -- ",df_partslist.at[idx,'Value'])

        # 部品番号をコピー
        df_sekkei_partslist.at[idx, 'Ref']=df_partslist.at[idx, 'Part']
        # print()
        # print(df_sekkei_partslist.loc[idx])
        # print()
        # print(mi[mi['CADname']==cadName])
    
    print('.. finished')
    print()
    # print(df_sekkei_partslist)
    return df_sekkei_partslist

#
#   組配部品表（PBAN,構成表用）を作成する
#  引数：
#   1:設計部品表（必須）
def make_kumihai_partslist(df_partslist):
    global mi
    
    df_kumihai_partslist = pd.DataFrame({'Mfr':[], 'Category':[], 'Partnumber':[], 'PartName':'', 'Count':[], 'NumTerminals':[], 'UnitPrice':[]})
    #   Partnumber ：部品の名前
    df_sekkei_partslist_groupedby_Partnumber=df_partslist.groupby('Partnumber')
   
    print('Generating Kumihai partslist...')
    for item in df_sekkei_partslist['Partnumber'].unique(): # 部品表のValue欄からユニークなリストを作ってiteration
        print(item)
        # print(mi[mi['Partnumber']==item] )

        # miから部品表のPartnmuber(型名)の値で検索する 
        # MIから作成した設計部品表を元にMIを検索しているので見つからないということはないが、
        # 設計部品表作成時にMIになかったものはNO-MIになるので、それはエラーになる。

        matchedMI=mi[mi['Partnumber']==item]
        if len( matchedMI )>0:
            # マッチがある場合      複数ある場合は最初を取る(iloc[0,:])
            df_kumihai_partslist.loc[item]=matchedMI.iloc[0,:]
        else:
            # マッチがない場合はエラーを表示
            df_kumihai_partslist.at[item,'Category']='NO-MI'
            print("WARNING: No MI Found on: Ref =",item,'/ PartNumber =',df_partslist[df_partslist['Partnumber']==item]['Ref'].tolist())

        # 同じ部品が使われている部品番号をリストとして取得
        partsOnItem = df_partslist.loc[df_sekkei_partslist_groupedby_Partnumber.groups[item],'Ref'].tolist()
        # print(partsOnItem)
        #部品使用数をカウントして'Count'rowに入れる
        df_kumihai_partslist.at[item,'Count']=len(partsOnItem)
        # 部品番号をリストから文字列に変換して'PartName'rowに入れる
        df_kumihai_partslist.at[item,'PartName']=', '.join(partsOnItem)
        # print(df_kumihai_partslist.loc[item])

    df_kumihai_partslist = df_kumihai_partslist.astype({'Count':int})
    # print(df_kumihai_partslist.dtypes)
    print('... finished.')
    return df_kumihai_partslist


# ファイルの準備
s_argv_fail_message=''
if len(sys.argv)>1 : # 引数がなければエラー
    p_argv1=Path(sys.argv[1])
    sheet_extract = 0
    if p_argv1.is_file() : # 引数のファイルの実体がなければエラー
        # print(p_argv1.resolve())
        print(p_argv1.parent.resolve())
        s_output_filename = str(p_argv1.stem)
        s_output_path = str(p_argv1.parent.resolve())+'/partslist/'
    else: 
        s_argv_fail_message = "Specified file did not exist or not a file."
    
else:
    s_argv_fail_message = 'No input files specified.'

# ファイルの準備中にエラーがある場合は終了
if len(s_argv_fail_message)!=0 :
    print(s_argv_fail_message)
    exit()

# ファイル出力のためのパスを作成    引数のBOMファイルがあるディレクトリの中にディレクトリpartlistを作る
os.makedirs(s_output_path, exist_ok=True)

# 出力する部品表に必要な要素の指定（MIのコラム名を指定する）
#   構成部品表
columun_list_kousei_CSV = ['Category', 'Partnumber', 'Mfr',  'Count', 'UnitPrice']
#   PBAN組配用部品表 
columun_list_PBAN_CSV = ['Mfr', 'Category', 'Partnumber', 'PartName', 'Count', 'NumTerminals']

# MI読み込み
print('Reading MI. -- ',end='')

#   必要なrowだけ読み込んで、コラム名を設定（ファイルからはコラム名は読み込めない）
mi=pd.read_table(str(p_dic_file), header=None).loc[:,[0,1,2,3,9,14,7,8]]
mi.columns=['Category','Partnumber','Mfr','CADname','NumTerminals','UnitPrice','SolderingMethod','SizeCode']

print('  completed.')

# BOM読み込み
print('Reading BOM file. -- ',end='')
# ki-CADのBOMデータの取り込み
# 空行はカウントされないので四行目がヘッダ名の定義に相当する。Ki-CADの方のBOM出力スクリプトに依存するので注意
# ここでは、全ての行のPart(Ref),Value,Sheet(Part), Package, Sizeのデータを抜き出す。
df_partslist = pd.read_table(str(p_argv1), header = 4, sep="\t", dtype={'Ref':'object','Value':'object','Part':'object','Package':'object','Size':'object'}).loc[:,['Ref','Value','Part','Package','Size']]   # table reading succeed.
df_partslist = df_partslist.rename(columns={'Part':'Sheet'}).rename(columns={'Ref':'Part'}) # change column name
print('  completed.')

# print(df_partslist)
# 出力するファイル名（ベース）の作成    実際にはさらにポストフィックスある
s_export_filename = s_output_path + s_output_filename 
s_postfix = ''

# BOMから設計部品表を作成
df_sekkei_partslist=make_sekkei_partslist(df_partslist)
# 設計部品表ファイルを作成
df_sekkei_partslist.sort_values('Ref').to_csv(s_export_filename + '_sekkei' + '.csv', encoding='utf_8_sig', header=False, index=False)


# 設計部品表から組配部品表を作成
df_kumihai_partslist=make_kumihai_partslist(df_sekkei_partslist)
# PBANの組配用部品表を作成
df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_PBAN_CSV].to_csv(s_export_filename + '_PBAN' + s_postfix + '.csv', encoding='shift_jis', header=False, index=False)
# 構成部品表を作成
df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_kousei_CSV].to_csv(s_export_filename + '_kousei' + s_postfix + '.csv', header=False, index=False)


print('--- All process completed.')