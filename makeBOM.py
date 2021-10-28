#!/usr/bin/python3
# -*- coding: utf-8 -*-
# 
# KI-CAD の部品表出力を在庫部品リストと照合して、組配用部品表、構成部品表のリストを作る
#       v2.0    2021/10  miyamoto

#   出力ファイルはCSV, UTF-8, CR/LF
#       組配部品表はUTF-8(BOM)  ->エクセルで読める
#   辞書ファイル（在庫部品）はBoxの上ある。

# 実行は：
#   $ python3 makeBOM.py {PATHtoTargetFile}/EagleBOMファイル
# 引数：
#   1:EagleBOMファイルの指定（必須）
#   

from numpy import dtype
import pandas as pd
import sys
from pathlib import Path

# 定数、関数の定義
s_dic_filename = 'FOR_KICAD_parts_dic_ORG.txt'
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

    df_sekkei_partslist = pd.DataFrame({'Ref':[], 'Category':[], 'Partnumber':[],'Mfr':[]},dtype=object)
    # print(df_sekkei_partslist.dtypes)
    # dtypeで表のデータの型を前もって指定しておかないと、Floatになる。
    # 他のdataframeを代入するとその型に上書きされるけれど、要素として代入しようとすると型が合わないのでエラーになる。

    # print(df_partslist.index.values)
    # print(df_partslist['Part'])
    # df_sekkei_partslist['Ref']=df_partslist['Part']

    # print(df_partslist.dtypes)
    print('Generating Sekkei Partslist....')
    for idx in df_partslist.index.values:
        # print(df_partslist.loc[idx,'Part'])
        cadName=df_partslist.at[idx,'Value']
        # print(df_partslist.at[idx, 'Part'])
        # print(df_sekkei_partslist.loc[idx])
        sizeCode=df_partslist.at[idx,'Size']
        solderingMethod=df_partslist.at[idx,'Soldering']

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
    # ファイルに落とす（テスト用のコード）  あとできちんとコードを書く
    df_sekkei_partslist.to_csv('sekkei_partslist' + '_kousei' + '' + '.csv', header=False, index=False)

    # print(df_sekkei_partslist.groupby('Partnumber').groups['RK73H1ETTP1000F'])
    df_sekkei_partslist_groupedby_Partnumber=df_sekkei_partslist.groupby('Partnumber')
    # print(df_sekkei_partslist['Partnumber'].unique())

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
            print("WARNING: No MI Found on: Ref =",item,'/ PartNumber =',df_sekkei_partslist[df_sekkei_partslist['Partnumber']==item]['Ref'].tolist())

        # 同じ部品が使われている部品番号をリストとして取得
        partsOnItem = df_sekkei_partslist.loc[df_sekkei_partslist_groupedby_Partnumber.groups[item],'Ref'].tolist()
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
mi=pd.read_table(str(p_dic_file), header=None).loc[:,[0,1,2,3,9,14,7,8]]
mi.columns=['Category','Partnumber','Mfr','CADname','NumTerminals','UnitPrice','SolderingMethod','SizeCode']
# print(mi.dtypes)
print('  completed.')

# 部品表読み込み
print('Reading BOM file. -- ',end='')
# df_partslist = pd.read_table(str(p_argv1), header = 4, delim_whitespace=True).loc[:,['Part','Value','Sheet']]
# ki-CADでのデータの取り込み
df_partslist = pd.read_table(str(p_argv1), header = 4, sep="\t", dtype={'Ref':'object','Value':'object','Part':'object','Soldering':'object','Size':'object'}).loc[:,['Ref','Value','Part','Soldering','Size']]   # table reading succeed.
df_partslist = df_partslist.rename(columns={'Part':'Sheet'}).rename(columns={'Ref':'Part'}) # change column name
print('  completed.')

# print(df_partslist)

# 空行はカウントされないので四行目がヘッダ名の定義に相当する。
# スペース区切りのためdelim_whitespace=Trueを入れた
# ここでは、全ての行のPart,Value,Sheetのデータを抜き出す。

# print('All Sheets. -- ',end='')
s_export_filename = s_output_path + s_output_filename 
s_postfix = '_A'
df_kumihai_partslist=make_kumihai_partslist(df_partslist)
df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_PBAN_CSV].to_csv(s_export_filename + '_PBAN' + s_postfix + '.csv', encoding='utf_8_sig', header=False, index=False)
df_kumihai_partslist.sort_values('Mfr').loc[:,columun_list_kousei_CSV].to_csv(s_export_filename + '_kousei' + s_postfix + '.csv', header=False, index=False)


print('--- All process completed.')