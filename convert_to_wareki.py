import datetime
from pathlib import Path

# 設定ファイル読み込み
def read_info():
    master_file = Path.cwd()/'システム設定'/'和暦.csv'
    with open(master_file, encoding='cp932') as f:
        l = f.readlines()
    d = {}
    first_loop = True
    for i in l:
        if not first_loop:
            row_list = i.split(',')
            datetime.datetime.strptime(row_list[1].replace('\n',''), '%Y/%m/%d')
            d[row_list[0]] = datetime.datetime.strptime(row_list[1].replace('\n',''), '%Y/%m/%d')
        if first_loop:
            first_loop = False
    sorted_d = sorted(d.items(), key=lambda x:x[1], reverse=True)
    return(sorted_d)

# 和暦返還 ※引数に変換したい日付を「YYYYmmdd」形式で渡すと和暦が返ってくる。引数を指定しないとシステム日付を返す
def convert_to_wareki(gappi=datetime.datetime.now().strftime('%Y%m%d')):
    target_date = datetime.datetime.strptime(gappi, '%Y%m%d')
    for i in read_info():
        if i[1] <= target_date:
            nengo = i[0]
            target_year = str((target_date.year - i[1].year) + 1)
            break
    return(f'{nengo} {target_year} 年 {target_date.month} 月 {target_date.day} 日')


# 年月日変換
def year_conversion(gappi):
    # 和暦とコードの対応表df作成
    master_dict = {
        '1':{'和暦': '明治', '西暦': '1867'},
        '2':{'和暦': '大正', '西暦': '1911'},
        '3':{'和暦': '昭和', '西暦': '1925'},
        '4':{'和暦': '平成', '西暦': '1988'},
        '5':{'和暦': '令和', '西暦': '2018'},
    }

    # 引数が「YYYYmm（6桁）」の場合は「YYYY年mm月」に変換
    if len(gappi) == 6:
        return (gappi[:4] + '年' + str(int(gappi[4:])) + '月')
    # 和暦がコードになっている場合の西暦変換
    elif len(gappi) == 5:
        target_year = master_dict[gappi[0]]['西暦']
        year = int(target_year) + int(gappi[1:3])
        return (str(year) + '年' + str(int(gappi[3:5])) + '月')
    # 和暦がコードになっている場合の和暦変換
    elif len(gappi) == 7:
        target_year = master_dict[gappi[0]]['和暦']
        return (target_year + " " + gappi[1:3] + ' 年' + str(int(gappi[3:5])) + ' 月' + str(int(gappi[5:])) + ' 日')
    # 引数が「YYYYmmdd」なら「YYYY年mm月dd日」に変換
    elif len(gappi) == 8:
        return (gappi[:4] + " 年" + str(int(gappi[4:6])) + ' 月' + str(int(gappi[6:])) + ' 日')


#
# if __name__ == '__main__':
#     print(test_year_conversion('3310202'))