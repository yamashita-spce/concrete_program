#-*- Auto_excel GUI-interface-*-
#  python 3.6.5

import pandas as pd
import tkinter as tk
import numpy as np
import tkinter.filedialog as fd

global i 
global root
global file_list
global ouf
global out_idx
i, file_list,out_idx = 0, [""], 0

def main(filelist):

    print("running...")
    k,j = 0,1
    for input_file_name in file_list:
        if input_file_name == "":
            break
        info = np.zeros((3,10))
        multi_excel_sheets = pd.read_excel(input_file_name, engine="openpyxl", sheet_name = None)

        for sheet, df in multi_excel_sheets.items():
            
            if k == 0:
                df1 = df.iloc[2:12,24]
                df1 = df1.astype("object")

            if k == 3:
                info = info.T

                df1 = df1.reset_index(drop = True)
                df3 = pd.DataFrame(info)
                df1 = pd.concat([df1,df3], axis = 1)

              
                with pd.ExcelWriter(ouf, mode = "a") as writer:
                    df1.to_excel(writer, sheet_name = u"まとめ", index = None, columns = None)
                break
            #必要なカラムの取り出し
            df0 = df.iloc[2:,[14,18,17]]
            df0.columns = ["strain_gauge","laser","stress"]

            #df0 = df0[df0["strain_gauge"] > 0]
            #df0 = df0[df0["laser"] > 0]
            #df0 = df0[df0["stress"] > 0]
            
            #最大応力を検出
            stress_max = df0["stress"].max()
            max_index = df0[df0["stress"] == stress_max].index.tolist()

            strain = df0.loc[:max_index[0],"strain_gauge"]
            laser_disp = df0.loc[max_index[0]:,"laser"]

            #dropna() ←　NAN要素を排除
            
            strain = strain.dropna()
            strain_list = strain.tolist()
            laser_disp = laser_disp.dropna()
            laser_disp_list = laser_disp.tolist()

            #補正値検出
            Correction_value = float(laser_disp_list[0] - strain_list[-1])
            print("%s Correction value : %f" %(sheet, Correction_value))
            # Correction value はひずみゲージと外部変位計のずれを表す

            #補正
            for l in range(len(laser_disp_list)):
                laser_disp_list[l] = laser_disp_list[l] - Correction_value
               
            stlaser_list = strain_list + laser_disp_list

            window = np.ones(50)/50
            stlaser_np = np.convolve(stlaser_list, window, mode = "valid")
            
            #laser変位計の移動平均
            #window = np.ones(50)/50
            #laser_np = np.convolve(laser_disp_list,window,mode = "valid")
            
            #np →　series　変換
            #laser_mean = pd.Series(stlaser_np)
            #laser_mean = laser_mean.astype("object")

            strain_mean = pd.Series(stlaser_np)
            strain_mean.name = "strain"
            strain_mean = strain_mean.dropna()
            strain_mean_fin = strain_mean.reset_index(drop = True)

            #付加情報取得
            df2 = df.iloc[2:12,[24,25]]
            df2 = df2.astype("object")
            df2_info = df2.iloc[:,1]
            df2_info_np = df2_info.values
            info[k,0:] = df2_info_np

            df0 = df0.reset_index(drop = True)
            df0.insert(len(df0.columns)-1,"strain",strain_mean_fin)
            df0 = pd.concat([df0,df2], axis = 1)
            df0 = df0.reset_index(drop = True)

            #シートに書き込み
            if k == 0:
                with pd.ExcelWriter(ouf) as writer:
                    df0.to_excel(writer, sheet_name = sheet, index = None)
            else:
                with pd.ExcelWriter(ouf, mode = "a") as writer:
                    df0.to_excel(writer, sheet_name = sheet, index = None)
            k += 1
        j += 1
    
    print("\nComplete! [Enter >]")
    input()

def selct_file(self):
    global i
    if i < 3:
        inf = fd.askopenfilename(
            filetypes = [(".xlsm","*.xlsm"),(".xlsx","xlsx")]
            )
        if inf != "":
            file_list[i] = inf
            i += 1

            EditBox1 = tk.Entry(width = 60)
            EditBox1.insert(tk.END,file_list[0])
            EditBox1.place(x = 100, y = 10)
        
def output_folder(self,):
    global ouf
    global out_idx

    ouf = fd.asksaveasfilename(
        initialfile = "output_data",
        defaultextension = ".xlsx",
        filetypes=[("*.xlsx","xlsx"),("*.csv","csv")]
    )
    if ouf != "":
        out_idx = 1

    EditBox4 = tk.Entry(width = 60)
    EditBox4.insert(tk.END,ouf)
    EditBox4.place(x = 10, y = 80)

def internal_processing(self, ):
    global root
    
    if out_idx == 1 and i >= 1:

        root.destroy()
        main(file_list)

#GUIに関するコード
def gui():
    print("=============================================================")
    print("NEW11_CONCRETE.EXE  pythonversion : 3.6.5  date : 2021/09/01")
    print("Library : pandas, tkinter, openpyxl, numpy")
    print("=============================================================")
    global root
    global i
    i = 0
    root = tk.Tk()
    root.title(u"Graphic user interface")
    root.geometry("500x150")

    Button1 = tk.Button(text = u"ファイルの選択")
    Button1.place(x = 10, y = 10)
    Button1.bind("<1>",selct_file)

    Button2 = tk.Button(text = u"出力先ファイルの作成")
    Button2.place(x = 10, y = 50)
    Button2.bind("<1>",output_folder)

    Button3 = tk.Button(text = u"実行", width = 8)
    Button3.place(x = 400, y = 110)
    Button3.bind("<1>",internal_processing)

    root.mainloop()

if __name__ == "__main__":

    gui()
    exit()