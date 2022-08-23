from pathlib import Path
import os, shutil, yaml, json
import convert_to_wareki

def satei_path():
	return(Path.cwd()/'査定csv')

def uke_path():
	return(Path.cwd()/'UKE')

def chk():
	result = True
	sateifiles = list(satei_path().glob('**/*.csv'))
	ukefiles = list(uke_path().glob('**/*.UKE'))
	if not sateifiles or not ukefiles:
		result = False
	return(result)

def delete_file(target_folder):
    # 対象ファイルフォルダを空にする
    files = target_folder.iterdir()
    for file in files:
        if Path.is_dir(file):
            shutil.rmtree(file)
        elif Path.is_file(file):
            os.remove(file)
        else:
            pass

def delete():
	if satei_path().exists():
		delete_file(satei_path())
	if uke_path().exists():
		delete_file(uke_path())

def get_satei_list():
	files = satei_path().glob('**/*.csv')
	for file in files:
		openfile = file
	with open(openfile, encoding='cp932') as f:
		l = f.readlines()
	return(l)

def read_uke():
	files = uke_path().glob('**/*.UKE')
	for file in files:
		with open(file, encoding='cp932') as f:
			l = f.readlines()
	return(l)

def get_ir_info():
	tensuhyokubun_t = ''
	meisaisyokubun1_t = ''
	meisaisyokubun2_t = ''
	saishinsasyubetsu_t = ''
	iryoukikan = ''
	seikyunengetsu = ''
	for row in read_uke():
		text_list = row.split(',')
		if text_list[0] == 'IR':
			if text_list[1] == '1' or text_list[1] == 1:
				tensuhyokubun_t = ('1.医科', '3.歯科', '4.調剤', '6.訪問')
				meisaisyokubun1_t = ('１ 単独', '２ 併用', '３ 老健')
				meisaisyokubun2_t = ('1=本人入院', '2=本人外来',
										  '3=未就学者入院', '4=未就学者外来',
										  '5=家族入院', '6=家族外来', '7=高齢者入院一般',
										  '8=高齢者外来一般', '9=高齢者入院7割', '0=高齢者外来7割')
				saishinsasyubetsu_t = ('１ 一次審査', '２ 突合再審査', '３ 再 審 査')

			elif text_list[1] == '2' or text_list[1] == 2:
				tensuhyokubun_t = ('1.医科', '3.歯科', '4.調剤')
				meisaisyokubun1_t = ('1国保', '3後期', '4退職')
				meisaisyokubun2_t = ('1本入', '2本外', '3六入', '4六外', '5家入', '6家外', '7高入-', '8高外-', '9高入7', '0高外7')
				saishinsasyubetsu_t = ('1 原審査', '2 再審査')
			iryoukikan = text_list[4]
			seikyunengetsu = convert_to_wareki.year_conversion(text_list[7])

			break
	return(tensuhyokubun_t, meisaisyokubun1_t, meisaisyokubun2_t, saishinsasyubetsu_t, iryoukikan, seikyunengetsu)

def patient_detail(receipt_num, patient_id):
	result = ""
	target_row = False
	file_category = 'C'
	index_list = []
	for row in read_uke():
		text_list = row.split(',')
		if text_list[0] == 'RE':
			if text_list[1] == receipt_num and text_list[13] == patient_id:
				target_row = True
			else:
				target_row = False
		if target_row:
			result += row
			index_list.append(row.split(',')[0])
	return(result, list(set(index_list)))

def read_sateicsv():
	target_folder = Path.cwd()/'査定CSV'
	files = target_folder.glob("**/*.csv")
	for file in files:
		with open(file, encoding='cp932') as f:
			l = f.readlines()
	return (l)

def dept():
	system_file = Path.cwd()/'システム設定'/'dept_master.txt'
	with open(system_file, encoding='utf8') as f:
		l = f.readlines()

	master_dept = {}
	for row in l:
		row_list = row.split(',')
		master_dept[row_list[0]] = row_list[1].replace('\n', '')
	return(master_dept)

def dpc_or_dekidaka():
	for row in read_uke():
		row_list = row.split(',')
		if row_list[0] == 'RE' and len(row_list) == 30:
			result = 'DPC'
		elif row_list[0] == 'RE' and len(row_list) == 38:
			result = '出来高'
	return(result)

def get_uke_config():
	open_file = f'{dpc_or_dekidaka()}.yaml'
	file_path = Path.cwd()/'システム設定'/'yaml'/open_file
	with open(file_path, encoding='utf8') as y:
		config = yaml.safe_load(y.read())
	return(config)

# 社保か国保か
def syaho_or_kokuho():
	files = uke_path().glob('**/*.UKE')
	for file in files:
		with open(file, encoding='cp932') as f:
			l = f.readline()
	if l.split(',')[1] == '2' or l.split(',')[1] == 2:
		result = 2
	elif l.split(',')[1] == '1' or l.split(',')[1] == 1:
		result = 1
	return(result)

def get_hyozyunhutangaku_codelist():
	filedir = Path.cwd()/'システム設定'
	file = filedir/'標準負担額_レセ電コード.csv'
	with open(file) as f:
		rows = f.readlines()
	return [row.replace('\n', '') for row in rows]