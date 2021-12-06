import lottery_kosen_festival as lot
import powerpoint as pp

lottery = lot.lotteryKosenFestival()

command = 'y'
lottery.display_all_attendee()

while(1):
    in_str = input('当選者を何人にしますか?: ')
    # 入力が整数じゃなければやり直し
    try:
        winner_count = int(in_str)
    except ValueError:
        print('エラーメッセージ: 数字を入力してください')
        continue
    
    print('-----抽選結果-----')
    lottery.select_winner(winner_count)
    #lottery.display_all_winner()
    #lottery.display_all_attendee()
    
    break
    """
    command = input('続けて抽選を行いますか？y/n: ')
    if command != 'y':
        print("抽選を終了します")
        break
    """


#パワーポイントインスタンスの作成
ppoint1 = pp.powerpoint()
ppoint2 = pp.powerpoint()

ppoint1.add_front_cover('当選者発表')
ppoint1.add_each_list_1(lottery.list_all_winner_number)

ppoint2.add_front_cover('当選者発表')
ppoint2.add_each_list_2(lottery.list_all_winner_number)

ppoint1.save_pptx('lottery_result1')
ppoint2.save_pptx('lottery_result2')