import lottery_kosen_festival as lot
import powerpoint as pp
import gift

lottery = lot.lotteryKosenFestival()
gift = gift.gift()

while(1):
    command = input('コマンドを入力してください: ')
    
    if command == 'start':
        lottery.display_all_attendee()
        
        print('抽選を開始します')
        in_str = input('当選者を何人にしますか?: ')
        # 入力が整数じゃなければやり直し
        try:
            winner_count = int(in_str)
        except ValueError:
            print('エラーメッセージ: 数字を入力してください')
            continue
        
        print('-----抽選結果-----')
        lottery.select_winner(winner_count)
        
        
    # パワーポイントの作成
    elif command == 'power':
        ppoint1 = pp.powerpoint()
        ppoint2 = pp.powerpoint()

        ppoint1.add_front_cover('当選者発表')
        ppoint1.add_each_list_1(lottery.list_winner)

        ppoint2.add_front_cover('当選者発表')
        ppoint2.add_each_list_2(lottery.list_winner)

        ppoint1.save_pptx('lottery_result1')
        ppoint2.save_pptx('lottery_result2')
        continue
    
    # 全当選者の表示
    elif command == 'winner':
        lottery.display_winner()
        continue
    
    # エクセルファイルを再読込する
    elif command == 'read':
        lottery.read_excel()
        print('エクセルファイルを読み込みました')
        continue
    
    elif command == 'gift':
        continue
    
    # プログラムを終了する
    elif command == 'exit':
        print('抽選を終了します')
        break
    
    
    

    
        



