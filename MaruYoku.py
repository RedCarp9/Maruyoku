import os,sqlite3,datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


# =============================================================================
#　データベース内のtable名’table_name’にデータが存在するか確認
# =============================================================================
def table_isexist(conn, cur, table_name):
    cur.execute(f"""
        SELECT COUNT(*) FROM sqlite_master 
        WHERE TYPE='table' AND name='{table_name}'
        """)
    if cur.fetchone()[0] == 0:
        return False
    return True 


# =============================================================================
# 実行ボタンを押したときの処理
# =============================================================================
def exection_click():
    
    mode_dict = {'モードを選択してください':'mode0','Step1 年間入場・車両情報の読込':'mode1',
                 'Step2 材料リスト読込・個別ファイル出力':'mode2','Step3 マルヨク出力':'mode3'}

    mode = mode_dict[mode_box.get()]
    
    # 車種一覧
    car_type = ['32','40','47','54','1000','1200','1500','185','186','2000','N2000','2600','2700',
                '5000','6000','7000','7200','8000','8600','k32','t6000','t45000','tkt8000','tkt9640']
    month = ['January','February','March','April','May','June','July','August','September','October','November','December'] 
    maruyoku_rename = {'code':'コード','name':'部品名','January':'1月','February':'2月','March':'3月','April':'4月','May':'5月','June':'6月',
                    'July':'7月','August':'8月','September':'9月','October':'10月','November':'11月','December':'12月'}
    db_rename = {'コード':'code','部品名':'name','1月':'January','2月':'February','3月':'March','4月':'April','5月':'May','6月':'June',
                 '7月':'July','8月':'August','9月':'September','10月':'October','11月':'November','12月':'December'}
    
    # データベース名’MyProperty.db’を準備
    conn = sqlite3.connect('MaruYoku_DB.db')
    # カーソルオブジェクトの作成
    cur = conn.cursor()
    
    if mode == 'mode0':
        # サブウィンドウの定義
        subwindow = tk.Toplevel()
        subwindow.title('エラー')
        subwindow.geometry("230x80+500+400")
        sub_label = tk.Label(subwindow, text="\nモードを選択してください", font=("",12))
        sub_label.pack(anchor='center')
        
    if mode == 'mode1':
        # サブウィンドウの定義
        subwindow = tk.Toplevel()
        subwindow.title('Step1_実行中')
        subwindow.geometry("230x100+500+400")
        sub_label = tk.Label(subwindow, text="\nStep1\n実行中", font=("",16))
        sub_label.pack(anchor='center')
        
        try:
            # ファイルセレクト画面
            typ = [('','*')]
            dirc = os.getcwd()
            fle = filedialog.askopenfilenames(filetypes = typ, initialdir = dirc)
            
            # 年間入場のデータを取得
            df = pd.read_excel(fle[0], dtype=str, header=2, sheet_name = "年間入場")
            df.to_sql('data_entry', conn, if_exists='replace',index=None)
            # 車両情報を取得
            df = pd.read_excel(fle[0], dtype=str, header=2, sheet_name = "車両情報")
            df.to_sql('data_car', conn, if_exists='replace',index=None)  
            
            # サブウィンドウ消去
            subwindow.destroy()
            # サブウィンドウの定義
            subwindow = tk.Toplevel()
            subwindow.title('Step1 終了')
            subwindow.geometry("230x80+500+400")
            sub_label = tk.Label(subwindow, text="\nStep1終了", font=("",16))
            sub_label.pack(anchor='center')
            
        except:
            # サブウィンドウ消去
            subwindow.destroy()
            # サブウィンドウの定義
            subwindow = tk.Toplevel()
            subwindow.title('エラー')
            subwindow.geometry("230x130+500+400")
            sub_label = tk.Label(subwindow, text="\nエラーが発生しました\n\nExcelファイルを\n閉じてください", font=("",16))
            sub_label.pack(anchor='center')
        
    
    if mode == 'mode2':
        # サブウィンドウの定義
        subwindow = tk.Toplevel()
        subwindow.title('Step2_実行中')
        subwindow.geometry("230x100+500+400")
        sub_label = tk.Label(subwindow, text="\nStep2\n実行中", font=("",16))
        sub_label.pack(anchor='center')
        
        try:
            # ファイルセレクト画面
            typ = [('','*')]
            dirc = os.getcwd()
            fle = filedialog.askopenfilenames(filetypes = typ, initialdir = dirc)
            
            # 材料情報を取得
            for i in car_type:
                try:
                    # Excelファイルの取り込み、DB登録
                    df = pd.read_excel(fle[0], dtype=str, header=1, sheet_name = i)
                    df.to_sql('data_mat_'+i, conn, if_exists='replace',index=None)
                    
                    # 各車種の材料データフレームの作成
                    df_mat = df.loc[:,['code','name']]
                    # 重複消去
                    df_mat = df_mat[~df_mat.duplicated()]
                    df_mat = df_mat.reset_index(drop=True)
                    # 新たな列の追加
                    for j in month[3:]+month[:3]: df_mat[j] = 0
                    df_mat.to_sql('data_ind_mat_'+i, conn, if_exists='replace',index=None)
                except: continue
                
            # 共通部品等の情報を取得
            df = pd.read_excel(fle[0], dtype=str, header=1, sheet_name = 'common')
            df.to_sql('data_mat_common', conn, if_exists='replace',index=None)
            
            # 各車種の材料データフレームの作成
            df_mat = df.loc[:,['code','name']]
            # 重複消去
            df_mat = df_mat[~df_mat.duplicated()]
            df_mat = df_mat.reset_index(drop=True)
            # 新たな列の追加
            for j in month[3:]+month[:3]: df_mat[j] = 0
            # DB登録
            df_mat.to_sql('data_ind_mat_common', conn, if_exists='replace',index=None)
    
    
            # 個別設定ファイルの作成
            if table_isexist(conn, cur, 'data_entry'):
                # Excelファイルフォルダ、ファイル名設定
                dirc = os.getcwd()+"/材料の個別設定"
                # フォルダ作成
                os.makedirs(dirc, exist_ok = True)
                # 個別設定用のファイル名指定
                # file_name = "/kobetsu"+str(datetime.date.today())+".xlsx"
                dt = datetime.datetime.now()
                file_name = "/kobetsu"+str(datetime.date.today())+"_"+str(dt.hour)+"_"+str(dt.minute)+".xlsx"
                dirc += file_name
                writer = pd.ExcelWriter(dirc, engine = 'xlsxwriter')
                
                for i in car_type:
                    try:
                        # 入場車情報の取り出し
                        query = f"""
                            select a_{i}, b_{i}, c_{i}
                            from data_entry
                            where a_{i} is not null
                        """
                        df_entry = pd.read_sql_query(sql = query, con = conn)
                    
                        if len(df_entry['a_'+i]) >0:
                            # DBから車両情報取り出し
                            query = f"""
                                select data_car.d_{i}, data_car.e_{i}, data_car.f_{i}
                                from data_car, (
                                    select a_{i}
                                    from data_entry
                                    where a_{i} is not null)
                                where d_{i} in (a_{i})
                            """
                            df_car_inf = pd.read_sql_query(sql = query, con = conn)
                            
                            # DBから個別の材料リスト取り出し
                            query = f"""
                                select code, name, type
                                from data_mat_{i}
                                where cond == '個別'
                                order by type, code
                            """
                            df_ind_mat = pd.read_sql_query(sql = query, con = conn)
                            
                            # DBから材料リスト取り出し
                            query = f"""
                                select *
                                from data_ind_mat_{i}
                            """
                            df_mat = pd.read_sql_query(sql = query, con = conn)
                            
                            # 個別用シートの追加、材料数を自動計算し各月に登録
                            for j in range(len(df_entry['a_'+i])):
                                # 個別リストの更新
                                df_ind_mat[str(df_entry['a_'+i][j])+'_'+month[int(df_entry['c_'+i][j])-1]] = 0
                                
                                # 自動計算されるリストの呼び出し
                                query = f"""
                                    select data_mat.code, data_mat.name, data_mat.num
                                    from (
                                        select *
                                        from data_mat_{i}
                                        where type=='{df_car_inf['f_'+i][j]}') as data_mat
                                    where data_mat.cond like '%{df_entry['b_'+i][j]}%' or 
                                        data_mat.cond == '毎回'
                                """
                                df_auto = pd.read_sql_query(sql = query, con = conn)
                                df_auto = df_auto.astype({'num':int})
                                df_auto = df_auto.rename(columns={'num':month[int(df_entry['c_'+i][j])-1]})
                                df_mat = df_mat.append(df_auto)
                                
                            df_mat = df_mat.groupby(['code','name']).sum()
                            df_mat.to_sql('data_ind_mat_'+i, conn, if_exists='replace')
                            
                            # 個別入力用のExcelシート作成
                            df_ind_mat.to_excel(writer,sheet_name=i, index=None)
                    except: continue
                
                # 共通部品の設定
                # DBから材料リスト取り出し
                query = """
                    select *
                    from data_ind_mat_common
                """
                df_mat = pd.read_sql_query(sql = query, con = conn)
                df_mat = df_mat.rename(columns=maruyoku_rename)
                # 個別入力用のExcelシート作成
                df_mat.to_excel(writer,sheet_name='common', index=None)
                
                # Excelファイルを保存
                writer.save()
                # Excelファイルを閉じる
                writer.close()
                
                # サブウィンドウ消去
                subwindow.destroy()
                # サブウィンドウの定義
                subwindow = tk.Toplevel()
                subwindow.title('Step2 終了')
                subwindow.geometry("230x130+500+400")
                sub_label = tk.Label(subwindow, text="\nStep2　終了\n\n個別リストを\n編集してください", font=("",16))
                sub_label.pack(anchor='center')
                
            else:
                # サブウィンドウ消去
                subwindow.destroy()
                # サブウィンドウの定義
                subwindow = tk.Toplevel()
                subwindow.title('エラー')
                subwindow.geometry("230x130+500+400")
                sub_label = tk.Label(subwindow, text="\nStep1を実行してください", font=("",16))
                sub_label.pack(anchor='center')
            
        except:
            # サブウィンドウ消去
            subwindow.destroy()
            # サブウィンドウの定義
            subwindow = tk.Toplevel()
            subwindow.title('エラー')
            subwindow.geometry("230x130+500+400")
            sub_label = tk.Label(subwindow, text="\nエラーが発生しました\n\nExcelファイルを\n閉じてください", font=("",16))
            sub_label.pack(anchor='center')
    
    
    if mode == 'mode3':
        # サブウィンドウの定義
        subwindow = tk.Toplevel()
        subwindow.title('Step3_実行中')
        subwindow.geometry("230x100+500+400")
        sub_label = tk.Label(subwindow, text="\nStep3\n実行中", font=("",16))
        sub_label.pack(anchor='center')
        
        try:
            # Excelファイルフォルダ、ファイル名設定
            dirc = os.getcwd()+"/マルヨク"
            # フォルダ作成
            os.makedirs(dirc, exist_ok = True)
                      
            # 個別設定用のファイル名指定
            file_name = "/maruyoku"+str(datetime.date.today())+".xlsx"
            dirc += file_name
            writer = pd.ExcelWriter(dirc, engine = 'xlsxwriter')
            
            # 個別設定Excelファイルからデータ取り込み、マルヨクの作成
            cols = ['code','name','April','May','June','July','August','September','October','November','December','January','February','March']
            df_maruyoku = pd.DataFrame(index=[], columns=cols)
            # ファイルセレクト画面
            typ = [('','*')]
            dirc = os.getcwd()+"/材料の個別設定"
            fle = filedialog.askopenfilenames(filetypes = typ, initialdir = dirc)
            
            # 車両情報を取得
            for i in car_type:
                try:
                    # ファイルを読み込み、車号の入っている部分の抜き出し
                    df = pd.read_excel(fle[0], header=0, sheet_name = i)
                    Columns = df.columns[3:]
                    
                    # DBから材料リスト取り出し
                    query = f"""
                        select *
                        from data_ind_mat_{i}
                    """
                    df_mat = pd.read_sql_query(sql = query, con = conn)
                    
                    # 各車種の材料リストに追加
                    for j in Columns:
                        mon = j.split('_')[-1]
                        df2 = df.loc[:,['code','name', j]]
                        df2 = df2.rename(columns={j:mon})
                        df_mat = df_mat.append(df2)
                    # df_maruyokuに追加
                    df_maruyoku = df_maruyoku.append(df_mat)
                    # DBに追加
                    df_mat.to_sql('data_maruyoku_'+i, conn, if_exists='replace',index=None)
                    # # df_matの編集
                    # df_mat = df_mat.groupby(['code','name']).sum()
                    # # 車種名でExcelに出力
                    # df_mat = df_mat.rename(columns=maruyoku_rename)
                    # df_mat.to_excel(writer,sheet_name=i)
                except: continue
            
            # 共通材料の追加
            df_mat = pd.read_excel(fle[0], header=0, sheet_name = 'common')
            # マルヨクに共通材料を書き出し
            df_mat.to_excel(writer,sheet_name='common', index=None)
            # 変数名変更
            df_mat = df_mat.rename(columns=db_rename)
            # df_maruyokuに追加
            df_maruyoku = df_maruyoku.append(df_mat)
            
            # DBに追加
            df_maruyoku.to_sql('data_maruyoku', conn, if_exists='replace',index=None)
            
            # マルヨクの情報を出力
            query = """
                select code, name, sum(April) as April, sum(May) as May, sum(June) as June, sum(July) as July,
                    sum(August) as August, sum(September) as September, sum(October) as October, sum(November) as November,
                    sum(December) as December, sum(January) as January, sum(February) as February, sum(March) as March
                from data_maruyoku
                group by code
            """
            df = pd.read_sql_query(sql = query, con = conn)
            df = df.rename(columns=maruyoku_rename)
            # 個別入力用のExcelシート作成
            df.to_excel(writer,sheet_name='all', index=None)
            
            for i in car_type:
                try:
                    query = f"""
                        select code, name, sum(April) as April, sum(May) as May, sum(June) as June, sum(July) as July,
                            sum(August) as August, sum(September) as September, sum(October) as October, sum(November) as November,
                            sum(December) as December, sum(January) as January, sum(February) as February, sum(March) as March
                        from data_maruyoku_{i}
                        group by code
                    """
                    df = pd.read_sql_query(sql = query, con = conn)
                    df = df.rename(columns=maruyoku_rename)
                    # 個別入力用のExcelシート作成
                    df.to_excel(writer,sheet_name=i, index=None)
                    
                except: continue
                             
            # Excelファイルを保存
            writer.save()
            # Excelファイルを閉じる
            writer.close()
            
            # サブウィンドウ消去
            subwindow.destroy()
            # サブウィンドウの定義
            subwindow = tk.Toplevel()
            subwindow.title('Step3 終了')
            subwindow.geometry("230x100+500+400")
            sub_label = tk.Label(subwindow, text="\nStep3\n終了", font=("",16))
            sub_label.pack(anchor='center')
            
        except:
            # サブウィンドウ消去
            subwindow.destroy()
            # サブウィンドウの定義
            subwindow = tk.Toplevel()
            subwindow.title('エラー')
            subwindow.geometry("230x130+500+400")
            sub_label = tk.Label(subwindow, text="\nエラーが発生しました\n\nExcelファイルを\n閉じてください", font=("",16))
            sub_label.pack(anchor='center')

# =============================================================================
# 起動
# =============================================================================
if __name__ == '__main__':

    # =============================================================================
    # 画面設定
    # =============================================================================
    ## tkinter_ウインドウの作成
    root=tk.Tk()
    root.title("Maruyoku_Ver3.01")
    root.resizable(0,0)
    root.geometry("550x160+340+200")
    
    ## フォントサイズ指定 フォントタイプはデフォルト
    fonts=("",14)
    
    ## モード選択画面の作成
    mode_list = ('Step1 年間入場・車両情報の読込','Step2 材料リスト読込・個別ファイル出力','Step3 マルヨク出力')
    
    input_mode_label = tk.Label(root, text="モード選択", font=fonts)
    input_mode_label.grid(row=1,column=1,padx=10,pady=10)
    
    mode_box = ttk.Combobox(root, values=mode_list, width=35, height=30, font=("",14))
    mode_box.grid(row=2,column=2)
    mode_box.set('モードを選択してください')
    
    ## ボタンの作成
    button=tk.Button(text="実行",font=fonts,command=exection_click)
    button.place(x=350,y=100,width=60,height=40)
    
    ## ウインドウの描画
    root.mainloop()
