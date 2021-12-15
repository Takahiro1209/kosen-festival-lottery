import pandas as pd
import random as rand

class lotteryKosenFestival():
    """
    Parameters
    ----------
    self.list_number : list(int)
        学籍番号のリスト
    self.list_name : list(str)
        名前のリスト
    self.list_winner : list(str, str)
        全当選者のリスト
        
    Returns
    -------
    None.

    """
    
    def __init__(self):
        self.read_excel()
        

    def read_excel(self):
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
        df = df.reset_index(drop='True')
        
        # それぞれリストに保存
        self.list_attendee = []
        self.list_number = []
        self.list_name = []
        self.list_winner = []

        # indexとその行をそれぞれ取り出す
        for index, item in df.iterrows():
            number = item['メール']
            name = item['名前']
            target = '@'
            index = number.find(target)
            number = number[:index]
            if not [number, name] in self.list_winner:
                self.list_number.append(number)
                self.list_name.append(name)
                self.list_attendee.append([number, name])
            
            


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
        list_rand_result = []
        list_rand_result = rand.sample(self.list_number, winner_count)

        count = 0
        list_winner = []
        #当選者を表示、リストから削除、全当選者リストに追加
        for number in list_rand_result:
            count += 1
            index = self.list_number.index(number)
            name = self.list_name.pop(index)
            number = self.list_number.pop(index)
            
            list_winner.append([number, name])
            
        # 当選者を表示
        for number, name in list_winner:
            print(number, name)
        
        list_winner = sorted(list_winner, reverse=False)
        # 今回の当選者を全当選者リストに追加
        self.list_winner.extend(list_winner)
        self.list_winner = sorted(self.list_winner, reverse=False)
        
        
    # 全当選者を表示
    def display_winner(self):
        print('-----全当選者-----')
        for number, name in self.list_winner:
            print(number, name)
        print('-----------------')
        
    
    # 全参加者を表示
    def display_all_attendee(self):
        print('-----全参加者--------------------')
        for number, name in self.list_attendee:
            print(number, name)
        print('--------------------------------')
