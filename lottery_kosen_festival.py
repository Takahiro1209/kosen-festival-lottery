import pandas as pd
import random as rand

class lotteryKosenFestival():
    def __init__(self):
        """
        Parameters
        ----------
        self.list_number : list(int)
            学籍番号のリスト
        self.list_name : list(str)
            名前のリスト
        self.list_all_winner_number : list(int)
            全当選者のリスト
            
        Returns
        -------
        None.

        """
        
        url ='attendees.xlsx'
        # 読み取る行数の設定
        pd.set_option('display.max_rows', 10)
        # elsxファイルの読み取り
        df_default = pd.read_excel(url, header=0, index_col=0)
        # メールと名前のみ抽出
        df = df_default.loc[:, ['メール', '名前']]
        # メールでソート
        df = df.sort_values('メール')
        # indexの振り直し
        self.df = df.reset_index(drop='True')
        
        # それぞれリストに保存
        self.list_number = []
        self.list_name = []
        # indexとその行をそれぞれ取り出す
        for index, item in df.iterrows():
            self.list_number.append(int(item['メール'][1:8]))
            self.list_name.append(item['名前'])
            
        self.list_all_winner_number = []

        
    # 当選者を抽選
    def select_winner(self, winner_count):
        """
        Parameters
        ----------
        winner_count : int
            当選者の数
        list_rand_result : list(int)
            当選者の学籍番号
        count : int
            表示用のカウンター
            
        Returns
        -------
        None.

        """
        # 選ばれた当選者の数
        if winner_count > len(self.list_number):
            print('エラーメッセージ: 当選者が多すぎます')
            print('抽選を行いませんでした')
            return
        
        # 学籍番号を選出
        list_rand_result = rand.sample(self.list_number, winner_count)
        self.list_all_winner_number.extend(list_rand_result)
        self.list_all_winner_number = sorted(self.list_all_winner_number, reverse=False)

        #降順でソート
        list_rand_result = sorted(list_rand_result, reverse=False)

        list_winner_number = []
        count = 0
        #当選者を表示、リストから削除、全当選者リストに追加
        for number in list_rand_result:
            count += 1
            index = self.list_number.index(number)
            number = self.list_number.pop(index)
            name = self.list_name.pop(index)
            list_winner_number.append(number)
            #当選者を表示
            print('当選者', count, ': ', number, ' ', name, sep='')
        
        
    
    # 全当選者を表示
    def display_all_winner(self):
        print('-----全当選者-----')
        for number in self.list_all_winner_number:
            print(number)
        print('-----------------')
        
    
    # 全参加者を表示
    def display_all_attendee(self):
        print('-----全参加者--------------------')
        print(self.df)
        print('--------------------------------')