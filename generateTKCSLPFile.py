#
import pandas as pd   
import numpy as np  

import os
from glob import glob


import argparse

def getdata():

    sample_data = [286,999,1,20180101,np.nan,np.nan,np.nan,0,1111,np.nan,4111,
                    np.nan,np.nan,0,100,np.nan,0,np.nan,0,np.nan,0,0,0,"調剤薬品",np.nan,0,0,
                    np.nan,np.nan,0,0,0,0,0]
    
    return sample_data

def read(MM):


    casio_dir = "/Volumes/myShare/register"
    target_dir = os.path.join(casio_dir,MM, "*.xlsm")
    filelist = glob(target_dir)
    
    input_book_ops = lambda ex:pd.ExcelFile(ex)
    input_books = list( map( input_book_ops, filelist ) ) 
    dates = [   os.path.basename(f).split(".")[0] for f in filelist ]
    
    print(dates)
    print("- " * 40 )
    # 
    #
    #
    chozai = []
    otc = []
    kowake = []
    jika = []

    data = getdata()

    TKC_data = []
    for d,b in zip(dates, input_books):



        target_sheet = b.sheet_names[3]
        input_sheet_df = b.parse(target_sheet, header=0, 
                        skiprows = 4, skip_footer = 7, parse_cols = "I" )    
        input_sheet_df  = input_sheet_df.reset_index(drop=True) 
        col_name = input_sheet_df.columns

        chozai.append( input_sheet_df.iloc[0,0] )
        otc.append( input_sheet_df.iloc[1,0] )
        kowake.append( input_sheet_df.iloc[2,0] )
        jika.append( input_sheet_df.iloc[3,0] )        

        if chozai[-1] > 0:
            data = getdata()
            data[3] = d
            data[23] = "調剤薬品"
            data[10] = 4111
            data[14] = chozai[-1]
            TKC_data.append(data)

        if otc[-1] > 0:
            data = getdata()
            data[3] = d
            data[23] = "OTC"
            data[10] = 4112
            data[14] = otc[-1]
            TKC_data.append(data)

        if kowake[-1] > 0:
            data = getdata()
            data[3] = d
            data[23] = "医薬品小分け"
            data[10] = 4112
            data[14] = kowake[-1]
            TKC_data.append(data)

        if jika[-1] > 0:
            data = getdata()
            data[3] = d
            data[23] = "自家消費"
            data[10] = 4113
            data[14] = jika[-1]
            TKC_data.append(data)

    print(chozai)
    print(otc)
    print(kowake)
    print(jika)

    #print(TKC_data)
    df = pd.DataFrame(TKC_data)
    df.to_csv("tkc.slp",sep="\t", index=False,encoding="Shift_JISx0213",header=None)

def main():

    parser = argparse.ArgumentParser(
            prog="TKC slp file", #プログラム名
            usage="generateTKCSLPFile.py --yyyymm 201808", #プログラムの利用方法
            description="TKC slp file", #「optional arguments」の前に表示される説明文
            epilog = "", #「optional arguments」後に表示される文字列
            add_help = True #[-h],[--help]オプションをデフォルトで追加するか
            )

    parser.add_argument("--yyyymm",required = True,type=str,
                    help="yyyymm is for CASIO directory.")

    args = parser.parse_args()
    YYYYMM = args.yyyymm

    read(YYYYMM)


if __name__ == "__main__":
    main()