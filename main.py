import tkinter, os, re, json, shutil, json
from pathlib import Path
from tkinter import ttk, filedialog
from tkinter.scrolledtext import ScrolledText
import import_file, convert_to_wareki
import win32com.client


class Application(ttk.Frame):
    def __init__(self, root):
        super().__init__(root,
                         width=1290, height=890,
                         borderwidth=4, relief='groove')

        def fixed_map(option):
            return [elm for elm in s.map('Treeview', query_opt=option) if elm[:2] != ('!disabled', '!selected')]

        # style
        s = ttk.Style()
        s.theme_use('default')
        s.configure(root, background='GhostWhite')
        s.configure("Treeview",
                    background="#D3D3D3",
                    foreground="black",
                    fieldbackground="#D3D3D3")
        s.map('Treeview', foreground=fixed_map('foreground'),
              background=fixed_map('background'))
        s.map('Label', foreground=fixed_map('foreground'),
              background=fixed_map('background'))
        s.configure("Vertical.TScrollbar", gripcount=0,
                    background="#75A9FF", darkcolor="LightSkyBlue",
                    lightcolor="LightSkyBlue",
                    troughcolor="GhostWhite", bordercolor="blue",
                    arrowcolor="white")
        s.configure("war.TCombobox", fieldbackground="#FFEEFF", selectbackground="#FFEEFF", selectforeground="black")
        s.configure("TCombobox", fieldbackground="white", selectbackground="white", selectforeground="black")



        self.root = root
        self.pack()
        self.pack_propagate(0)
        self.sub_window = None
        self.create_widgets()

    def create_widgets(self):
        # タブ
        self.notebook = ttk.Notebook(root, height=806, width=1275)
        self.tab_import = tkinter.Frame(self.notebook, bg="white", height=805, width=1275)
        self.tab_saishinsaseikyusyo = tkinter.Frame(self.notebook, bg="white", height=805, width=1275)
        self.notebook.add(self.tab_import, text='インポート')
        self.notebook.add(self.tab_saishinsaseikyusyo, text='再審査請求書作成')
        self.notebook.place(x=10, y=5)

        # 査定CSVインポート
        self.input_import_sateicsv = ttk.Entry(self.tab_import, width=80, font=("", 20))
        self.btn_import_sateicsv = tkinter.Button(self.tab_import,
                                                  text='査定CSVをインポート',
                                                  font=("", "20"),
                                                  width=20,
                                                  command=self.satei_file_select
                                                  )
        self.input_import_sateicsv.pack(pady=[100, 5], ipady=5)
        self.btn_import_sateicsv.pack(pady=10, ipady=5)

        # UKEインポート
        self.input_import_uke = ttk.Entry(self.tab_import, width=80, font=("", 20))
        self.btn_import_uke = tkinter.Button(self.tab_import,
                                             text='UKEをインポート',
                                             font=("", "20"),
                                             width=20,
                                             command=self.uke_file_select
                                             )
        self.input_import_uke.pack(pady=[100, 5], ipady=5)
        self.btn_import_uke.pack(pady=10, ipady=5)

        self.btn_import = tkinter.Button(self.tab_import, bd=3, bg='moccasin',
                                         text="再審査請求書作成",
                                         font=("", "20"),
                                         width=20,
                                         command=self.read_files)
        self.btn_import.pack_forget()

        # 患者一覧
        self.frame_patient_list = tkinter.Frame(
            self.tab_saishinsaseikyusyo,
            width=225,
            height=795,
            background="white"
        )
        self.frame_patient_list.place(relx=0.005, rely=0.007)
        # 患者検索テキストボックス
        self.patientid_input = ttk.Entry(self.frame_patient_list, width=19)
        self.patientid_input.place(relx=0.01, rely=0.01)

        # 患者検索ボタン
        self.btn_patient_search = tkinter.Button(self.frame_patient_list, bd=3,
                                                 bg='moccasin', text='患者ID検索',
                                                 command=self.search)
        self.btn_patient_search.place(relx=0.55, rely=0.005)
        self.frame_patient_table = tkinter.Frame(self.frame_patient_list,
                                                 width=220, height=790,
                                                 background="white")
        self.frame_patient_table.place(relx=0.01, rely=0.045)

        # フッター
        self.frame_footer = tkinter.Frame(self, width=1283,
                                          height=33,
                                          background='#2C7CFF')
        self.frame_footer.place(relx=0.002, rely=0.95)
        self.btn_close = tkinter.Button(self.frame_footer,
                                        bd=3, bg='lightskyblue',
                                        width=15,
                                        text='閉じる',
                                        command=self.root.destroy)
        self.btn_close.place(relx=0.9, rely=0.05)

        # 患者一覧
        self.tree_patient = ttk.Treeview(self.frame_patient_table, height=36)
        self.tree_patient['columns'] = ("ri", "患者ID", "氏名")
        self.tree_patient["show"] = "headings"
        self.tree_patient.bind("<<TreeviewSelect>>", self.select_record)
        self.tree_patient.column("ri", width=20)
        self.tree_patient.column("患者ID", width=70)
        self.tree_patient.column("氏名", width=120)
        self.tree_patient.heading("ri", text="ri")
        self.tree_patient.heading("患者ID", text="患者ID")
        self.tree_patient.heading("氏名", text="氏名")
        self.tree_patient['displaycolumns'] = ('患者ID', '氏名')
        self.tree_patient.place(relx=0.01, rely=0.01)
        vbar = ttk.Scrollbar(self.frame_patient_table, orient='v',
                             command=self.tree_patient.yview)
        self.tree_patient.configure(yscrollcommand=vbar.set)
        vbar.pack(side=tkinter.LEFT, fill=tkinter.Y)
        vbar.place(relx=0.9, rely=0.01, height=750)

        # 再審査請求フレーム
        self.note_saishinsaseikyusyo = ttk.Notebook(
            self.tab_saishinsaseikyusyo,
            height=780, width=1035)
        self.tab_form = tkinter.Frame(
            self.note_saishinsaseikyusyo,
            bg='white',
            height=775,
            width=1030
        )
        self.tab_satei = tkinter.Frame(
            self.note_saishinsaseikyusyo,
            bg='white',
            height=775,
            width=1030
        )
        self.tab_uke = tkinter.Frame(
            self.note_saishinsaseikyusyo,
            bg='white',
            height=775,
            width=1030
        )
        self.note_saishinsaseikyusyo.add(self.tab_form, text='再審査請求書')
        self.note_saishinsaseikyusyo.add(self.tab_satei, text='査定CSV')
        self.note_saishinsaseikyusyo.add(self.tab_uke, text='レセ電コード')
        self.note_saishinsaseikyusyo.place(relx=0.18, rely=0.001)
        # 日付
        self.frame_date = ttk.Frame(self.tab_form, relief='sunken',
                                    width=200, height=29)
        self.v_year = tkinter.StringVar()
        self.t_year = ('平成', '令和')  # yamlファイルから読み込む
        self.cb_year = ttk.Combobox(self.frame_date,
                                    textvariable=self.v_year,
                                    values=self.t_year)
        self.v_year.set(self.t_year[1])
        self.input_year = ttk.Entry(self.frame_date)
        self.label_year = ttk.Label(self.frame_date, text="年")
        self.input_month = ttk.Entry(self.frame_date)
        self.label_month = ttk.Label(self.frame_date, text="月")
        self.input_day = ttk.Entry(self.frame_date)
        self.label_day = ttk.Label(self.frame_date, text="日")
        self.frame_date.place(relx=0.01, rely=0.01)
        self.cb_year.place(relx=0.01, rely=0.1, width=60)
        self.input_year.place(relx=0.315, rely=0.1, width=30)
        self.label_year.place(relx=0.463, rely=0.17)
        self.input_month.place(relx=0.537, rely=0.1, width=30)
        self.label_month.place(relx=0.683, rely=0.17)
        self.input_day.place(relx=0.760, rely=0.1, width=30)
        self.label_day.place(relx=0.910, rely=0.17)

        # 支部
        self.frame_shibu = ttk.Frame(self.tab_form, relief='sunken',
                                     width=200, height=29)
        self.label_shibu_pre = ttk.Label(self.frame_shibu, text="支払基金")
        self.input_shibu = ttk.Entry(self.frame_shibu)
        self.label_shibu_ape = ttk.Label(self.frame_shibu, text="支部 御中")
        self.frame_shibu.place(relx=0.01, rely=0.05)
        self.label_shibu_pre.place(relx=0.01, rely=0.17)
        self.input_shibu.place(relx=0.28, rely=0.1, width=80)
        self.label_shibu_ape.place(relx=0.683, rely=0.17)

        # 医療機関入力
        self.frame_iryokikan = ttk.Frame(self.tab_form, relief='sunken', width=810, height=60)
        self.label_address = ttk.Label(self.frame_iryokikan, text="保険医療機関の住所")
        self.input_address = ttk.Entry(self.frame_iryokikan)
        self.label_hospitalname = ttk.Label(self.frame_iryokikan, text="保険医療機関の名称")
        self.input_hospitalname = ttk.Entry(self.frame_iryokikan)
        self.label_founder = ttk.Label(self.frame_iryokikan, text="開設者氏名")
        self.input_founder = ttk.Entry(self.frame_iryokikan)
        self.label_tel = ttk.Label(self.frame_iryokikan, text="電話番号")
        self.input_tel = ttk.Entry(self.frame_iryokikan)
        self.frame_iryokikan.place(relx=0.21, rely=0.01)
        self.label_address.place(relx=0.01, rely=0.13)
        self.input_address.place(relx=0.15, rely=0.1, width=380)
        self.label_hospitalname.place(relx=0.01, rely=0.55)
        self.input_hospitalname.place(relx=0.15, rely=0.54, width=380)
        self.label_founder.place(relx=0.63, rely=0.13)
        self.input_founder.place(relx=0.71, rely=0.1, width=230)
        self.label_tel.place(relx=0.63, rely=0.55)
        self.input_tel.place(relx=0.71, rely=0.54, width=230)

        self.frame_saishinsa_torisage = tkinter.Frame(self.tab_form, width=900, height=29, background='white')
        self.label_saishinsa_torisage = tkinter.Label(self.frame_saishinsa_torisage, text='下記理由により、診療報酬等明細書を',
                                                      background='white')
        self.v_saishinsa_torisage = tkinter.StringVar()
        self.t_saishinsa_torisage = ('再審査', '取下げ')
        self.v_saishinsa_torisage.set(self.t_saishinsa_torisage[0])
        self.cb_saishinsa_torisage = ttk.Combobox(self.frame_saishinsa_torisage,
                                                  textvariable=self.v_saishinsa_torisage,
                                                  values=self.t_saishinsa_torisage)
        self.label_saishinsa_torisage_ape = tkinter.Label(self.frame_saishinsa_torisage,
                                                          text='願います。', background='white')
        self.frame_saishinsa_torisage.place(relx=0.01, rely=0.09)
        self.label_saishinsa_torisage.place(relx=0.01, rely=0.2)
        self.cb_saishinsa_torisage.place(relx=0.225, rely=0.2, width=70)
        self.label_saishinsa_torisage_ape.place(relx=0.31, rely=0.2)

        # 点数表区分
        self.frame_tensuhyokubun = ttk.Frame(self.tab_form, relief="sunken", width=140, height=29)
        self.label_tensuhyokubun = ttk.Label(self.frame_tensuhyokubun, text="点数表区分")
        self.v_tensuhyokubun = tkinter.StringVar()
        self.t_tensuhyokubun = ('1.医科', '3.歯科', '4.調剤')
        self.cb_tensuhyokubun = ttk.Combobox(self.frame_tensuhyokubun,
                                             textvariable=self.v_tensuhyokubun,
                                             values=self.t_tensuhyokubun,
                                             width=7)
        self.frame_tensuhyokubun.place(relx=0.01, rely=0.13)
        self.label_tensuhyokubun.place(relx=0.02, rely=0.17)
        self.cb_tensuhyokubun.place(relx=0.5, rely=0.1)

        # 医療機関（薬局）番号
        self.frame_iryokikanbangou = ttk.Frame(self.tab_form, relief="sunken", width=315, height=29)
        self.label_iryokikanbangou = ttk.Label(self.frame_iryokikanbangou, text="医療機関(薬局)番号")
        self.input_iryokikanbangou = ttk.Entry(self.frame_iryokikanbangou)
        self.frame_iryokikanbangou.place(relx=0.15, rely=0.13)
        self.label_iryokikanbangou.place(relx=0.01, rely=0.17)
        self.input_iryokikanbangou.place(relx=0.36, rely=0.12, width=200)

        # （旧）総合病院診療科
        self.frame_shinryouka = ttk.Frame(self.tab_form, relief="sunken", width=240, height=29)
        self.label_shinryouka = ttk.Label(self.frame_shinryouka, text="(旧)総合病院 診療科")
        # self.input_shinryouka = ttk.Entry(self.frame_shinryouka)
        self.v_shinryouka = tkinter.StringVar()
        self.t_shinryouka = tuple(import_file.dept().values())
        self.cb_shinryouka = ttk.Combobox(self.frame_shinryouka,
                                          textvariable=self.v_shinryouka,
                                          values=self.t_shinryouka)
        self.frame_shinryouka.place(relx=0.455, rely=0.13)
        self.label_shinryouka.place(relx=0.01, rely=0.17)
        # self.input_shinryouka.place(relx=0.465, rely=0.12, width=125)
        self.cb_shinryouka.place(relx=0.465, rely=0.12, width=125)

        # 診療（調剤）年月
        self.frame_shinryounengetsu = ttk.Frame(self.tab_form, relief="sunken", width=150, height=29)
        self.label_shinryounengetsu = ttk.Label(self.frame_shinryounengetsu, text="診療(調剤)年月")
        self.input_shinryounengetsu = ttk.Entry(self.frame_shinryounengetsu)
        self.frame_shinryounengetsu.place(relx=0.695, rely=0.13)
        self.label_shinryounengetsu.place(relx=0.01, rely=0.17)
        self.input_shinryounengetsu.place(relx=0.55, rely=0.12, width=64)

        # 請求（調剤）年月
        self.frame_seikyuunengetsu = ttk.Frame(self.tab_form, relief="sunken", width=150, height=29)
        self.label_seikyuunengetsu = ttk.Label(self.frame_seikyuunengetsu, text="請求(調剤)年月")
        self.input_seikyuunengetsu = ttk.Entry(self.frame_seikyuunengetsu)
        self.frame_seikyuunengetsu.place(relx=0.85, rely=0.13)
        self.label_seikyuunengetsu.place(relx=0.01, rely=0.17)
        self.input_seikyuunengetsu.place(relx=0.55, rely=0.12, width=64)

        # 明細書区分
        self.frame_meisaisyokubun = ttk.Frame(self.tab_form, relief="sunken", width=330, height=29)
        self.label_meisaisyokubun = ttk.Label(self.frame_meisaisyokubun, text="明細書区分")
        self.v_meisaisyokubun1 = tkinter.StringVar()
        self.t_meisaisyokubun1 = ('1国保', '3後期', '4退職')
        self.cb_meisaisyokubun1 = ttk.Combobox(self.frame_meisaisyokubun, textvariable=self.v_meisaisyokubun1,
                                               values=self.t_meisaisyokubun1)
        self.v_meisaisyokubun2 = tkinter.StringVar()
        self.t_meisaisyokubun2 = ('1本入', '2本外', '3六入', '4六外', '5家入',
                                  '6家外', '7高入-', '8高外-', '9高入7', '0高外7')
        self.cb_meisaisyokubun2 = ttk.Combobox(self.frame_meisaisyokubun,
                                               textvariable=self.v_meisaisyokubun2,
                                               values=self.t_meisaisyokubun2)
        self.frame_meisaisyokubun.place(relx=0.01, rely=0.17)
        self.label_meisaisyokubun.place(relx=0.01, rely=0.17)
        self.cb_meisaisyokubun1.place(relx=0.33, rely=0.1, width=72)
        self.cb_meisaisyokubun2.place(relx=0.55, rely=0.1)

        # 再審査等対象種別
        self.frame_saishinsasyubetsu = ttk.Frame(self.tab_form, relief="sunken", width=330, height=29)
        self.label_saishinsasyubetsu = ttk.Label(self.frame_saishinsasyubetsu, text="再審査等対象種別")
        self.v_saishinsasyubetsu = tkinter.StringVar()
        self.t_saishinsasyubetsu = ('1 原審査', '2 再審査')
        self.cb_saishinsasyubetsu = ttk.Combobox(self.frame_saishinsasyubetsu,
                                                 textvariable=self.v_saishinsasyubetsu,
                                                 values=self.t_saishinsasyubetsu,
                                                 style='war.TCombobox')
        self.frame_saishinsasyubetsu.place(relx=0.01, rely=0.21)
        self.label_saishinsasyubetsu.place(relx=0.01, rely=0.17)
        self.cb_saishinsasyubetsu.place(relx=0.33, rely=0.1)

        # 相手方薬局又は処方箋発行医療機関
        self.frame_syohouiryoukikan = ttk.Frame(self.tab_form, relief="sunken", width=675, height=58)
        self.label_syohouiryoukikan = ttk.Label(self.frame_syohouiryoukikan, text="相手方薬局又は処方\nせん発行医療機関")
        self.label_yakkyokucode = ttk.Label(self.frame_syohouiryoukikan, text="薬局(医療機関)コード")
        self.label_todohukenmei = ttk.Label(self.frame_syohouiryoukikan, text="都道府県名")
        self.label_yakkyokuname = ttk.Label(self.frame_syohouiryoukikan, text="薬局(医療機関)の名称")
        self.input_todohukenmei = ttk.Entry(self.frame_syohouiryoukikan)
        self.input_yakkyokucode = ttk.Entry(self.frame_syohouiryoukikan)
        self.input_yakkyokuname = ttk.Entry(self.frame_syohouiryoukikan)
        self.frame_syohouiryoukikan.place(relx=0.34, rely=0.17)
        self.label_syohouiryoukikan.place(relx=0.01, rely=0.17)
        self.label_yakkyokucode.place(relx=0.15, rely=0.12)
        self.label_todohukenmei.place(relx=0.325, rely=0.12)
        self.label_yakkyokuname.place(relx=0.15, rely=0.55)
        self.input_todohukenmei.place(relx=0.421, rely=0.075, width=100)
        self.input_yakkyokucode.place(relx=0.575, rely=0.075, width=280)
        self.input_yakkyokuname.place(relx=0.325, rely=0.525, width=450)

        # 保険者番号
        self.frame_hokensyabango = ttk.Frame(self.tab_form, relief="sunken", width=220, height=29)
        self.label_hokensyabango = ttk.Label(self.frame_hokensyabango, text="保険者番号")
        self.input_hokensyabango = ttk.Entry(self.frame_hokensyabango)
        self.frame_hokensyabango.place(relx=0.01, rely=0.255)
        self.label_hokensyabango.place(relx=0.01, rely=0.17)
        self.input_hokensyabango.place(relx=0.415, rely=0.12)

        # 記号・番号
        self.frame_kigoubango = ttk.Frame(self.tab_form, relief="sunken", width=210, height=29)
        self.label_kigoubango = ttk.Label(self.frame_kigoubango, text="記号・番号")
        self.input_kigoubango = ttk.Entry(self.frame_kigoubango)
        self.frame_kigoubango.place(relx=0.224, rely=0.255)
        self.label_kigoubango.place(relx=0.01, rely=0.17)
        self.input_kigoubango.place(relx=0.38, rely=0.12)

        # 公費負担番号①
        self.frame_kouhihutanbango = ttk.Frame(self.tab_form, relief="sunken", width=220, height=29)
        self.label_kouhihutanbango = ttk.Label(self.frame_kouhihutanbango, text="公費負担番号①")
        self.input_kouhihutanbango = ttk.Entry(self.frame_kouhihutanbango)
        self.frame_kouhihutanbango.place(relx=0.01, rely=0.29)
        self.label_kouhihutanbango.place(relx=0.01, rely=0.17)
        self.input_kouhihutanbango.place(relx=0.415, rely=0.12)

        # 受給者番号①
        self.frame_jukyusyabango = ttk.Frame(self.tab_form, relief="sunken", width=210, height=29)
        self.label_jukyusyabango = ttk.Label(self.frame_jukyusyabango, text="受給者番号①")
        self.input_jukyusyabango = ttk.Entry(self.frame_jukyusyabango)
        self.frame_jukyusyabango.place(relx=0.224, rely=0.29)
        self.label_jukyusyabango.place(relx=0.01, rely=0.17)
        self.input_jukyusyabango.place(relx=0.38, rely=0.12)

        # 公費負担番号②
        self.frame_kouhihutanbango2 = ttk.Frame(self.tab_form, relief="sunken", width=220, height=29)
        self.label_kouhihutanbango2 = ttk.Label(self.frame_kouhihutanbango2, text="公費負担番号②")
        self.input_kouhihutanbango2 = ttk.Entry(self.frame_kouhihutanbango2)
        self.frame_kouhihutanbango2.place(relx=0.01, rely=0.325)
        self.label_kouhihutanbango2.place(relx=0.01, rely=0.17)
        self.input_kouhihutanbango2.place(relx=0.415, rely=0.12)

        # 受給者番号②
        self.frame_jukyusyabango2 = ttk.Frame(self.tab_form, relief="sunken", width=210, height=29)
        self.label_jukyusyabango2 = ttk.Label(self.frame_jukyusyabango2, text="受給者番号②")
        self.input_jukyusyabango2 = ttk.Entry(self.frame_jukyusyabango2)
        self.frame_jukyusyabango2.place(relx=0.224, rely=0.325)
        self.label_jukyusyabango2.place(relx=0.01, rely=0.17)
        self.input_jukyusyabango2.place(relx=0.38, rely=0.12)

        # 氏名
        self.frame_patient_name = ttk.Frame(self.tab_form, relief="sunken", width=260, height=87)
        self.label_patient_cname = ttk.Label(self.frame_patient_name, text="ﾌﾘｶﾞﾅ")
        self.label_patient_name = ttk.Label(self.frame_patient_name, text="氏名")
        self.label_seibetsu = ttk.Label(self.frame_patient_name, text="性別")
        self.v_seibetsu = tkinter.StringVar()
        self.t_seibetsu = ('1.男', '2.女')
        self.cb_seibetsu = ttk.Combobox(self.frame_patient_name, textvariable=self.v_seibetsu, values=self.t_seibetsu,
                                        width=4)
        self.input_patient_name = ttk.Entry(self.frame_patient_name)
        self.input_patient_cname = ttk.Entry(self.frame_patient_name)
        self.label_seinengappi = ttk.Label(self.frame_patient_name, text="生年月日")
        self.input_seinengappi = ttk.Entry(self.frame_patient_name)
        self.frame_patient_name.place(relx=0.43, rely=0.255)
        self.label_patient_cname.place(relx=0.01, rely=0.1)
        self.label_patient_name.place(relx=0.01, rely=0.385)
        self.input_patient_cname.place(relx=0.145, rely=0.07, width=218)
        self.input_patient_name.place(relx=0.145, rely=0.38, width=218)
        self.label_seibetsu.place(relx=0.01, rely=0.7)
        self.cb_seibetsu.place(relx=0.12, rely=0.675)
        self.label_seinengappi.place(relx=0.35, rely=0.7)
        self.input_seinengappi.place(relx=0.56, rely=0.675, width=110)

        # 写しの有無
        self.frame_utsushinoumu = ttk.Frame(self.tab_form, relief="sunken", width=257, height=29)
        self.label_utsushinoumu = ttk.Label(self.frame_utsushinoumu, text="写しの有無")
        self.v_utsushinoumu = tkinter.StringVar()
        self.t_utsushinoumu = ('有', '無')
        self.cb_utsushinoumu = ttk.Combobox(self.frame_utsushinoumu, textvariable=self.v_utsushinoumu,
                                            values=self.t_utsushinoumu, style='war.TCombobox')
        self.frame_utsushinoumu.place(relx=0.43, rely=0.37)
        self.label_utsushinoumu.place(relx=0.01, rely=0.17)
        self.cb_utsushinoumu.place(relx=0.33, rely=0.1)

        # 当初請求点数
        self.frame_tousyoseikyutensu = ttk.Frame(self.tab_form, relief="sunken", width=183, height=29)
        self.label_tousyoseikyutensu = ttk.Label(self.frame_tousyoseikyutensu, text="当初請求点数")
        self.input_tousyoseikyutensu = ttk.Entry(self.frame_tousyoseikyutensu, justify='right')
        self.ten = ttk.Label(self.frame_tousyoseikyutensu, text="点")
        self.frame_tousyoseikyutensu.place(relx=0.685, rely=0.255)
        self.label_tousyoseikyutensu.place(relx=0.01, rely=0.17)
        self.input_tousyoseikyutensu.place(relx=0.62, rely=0.12, width=50)
        self.ten.place(relx=0.89, rely=0.17)

        # 一部負担金
        self.frame_itibuhutankin = ttk.Frame(self.tab_form, relief="sunken", width=135, height=29)
        self.label_itibuhutankin = ttk.Label(self.frame_itibuhutankin, text="一部負担金")
        self.input_itibuhutankin = ttk.Entry(self.frame_itibuhutankin, justify='right')
        self.en_i = ttk.Label(self.frame_itibuhutankin, text="点")
        self.frame_itibuhutankin.place(relx=0.863, rely=0.255)
        self.label_itibuhutankin.place(relx=0.015, rely=0.17)
        self.input_itibuhutankin.place(relx=0.48, rely=0.12, width=50)
        self.en_i.place(relx=0.855, rely=0.17)

        # 当初請求食事療養費
        self.frame_tousyoseikyuhi = ttk.Frame(self.tab_form, relief="sunken", width=183, height=29)
        self.label_tousyoseikyuhi = ttk.Label(self.frame_tousyoseikyuhi, text="当初請求食事療養費")
        self.input_tousyoseikyuhi = ttk.Entry(self.frame_tousyoseikyuhi, justify='right')
        self.en = ttk.Label(self.frame_tousyoseikyuhi, text="円")
        self.frame_tousyoseikyuhi.place(relx=0.685, rely=0.29)
        self.label_tousyoseikyuhi.place(relx=0.01, rely=0.17)
        self.input_tousyoseikyuhi.place(relx=0.62, rely=0.12, width=50)
        self.en.place(relx=0.89, rely=0.17)

        # 標準負担額
        self.frame_hyojunhutangaku = ttk.Frame(self.tab_form, relief="sunken", width=135, height=29)
        self.label_hyojunhutangaku = ttk.Label(self.frame_hyojunhutangaku, text="標準負担額")
        self.input_hyojunhutangaku = ttk.Entry(self.frame_hyojunhutangaku, justify='right')
        self.en_h = ttk.Label(self.frame_hyojunhutangaku, text="点")
        self.frame_hyojunhutangaku.place(relx=0.863, rely=0.29)
        self.label_hyojunhutangaku.place(relx=0.02, rely=0.17)
        self.input_hyojunhutangaku.place(relx=0.48, rely=0.12, width=50)
        self.en_h.place(relx=0.855, rely=0.17)

        # 取下げ理由
        self.frame_torisageriyu = ttk.Frame(self.tab_form, relief="sunken", width=320, height=29)
        self.label_torisageriyu = ttk.Label(self.frame_torisageriyu, text="取下げ理由")
        self.input_torisageriyu = ttk.Entry(self.frame_torisageriyu)
        self.frame_torisageriyu.place(relx=0.685, rely=0.325)
        self.label_torisageriyu.place(relx=0.01, rely=0.17)
        self.input_torisageriyu.place(relx=0.2, rely=0.12, width=250)

        # 減点内容
        self.frame_genten = ttk.Frame(self.tab_form, width=1035, height=260)
        self.frame_genten.place(relx=0.01, rely=0.42)

        def fetch(entries):
            num = 0
            for entrie in entries:
                entrie[0].insert(tkinter.END, str(num))
                entrie[1].insert(tkinter.END, str(num))
                entrie[2].insert(tkinter.END, str(num))
                num = num + 1

        header_dict = {'減点点数': [9, 0.02], '減点事由及び箇所': [55, 0.082], '減点内容': [95, 0.41]}
        for k, v in header_dict.items():
            header_label = ttk.Label(self.frame_genten, relief='raised', text=k, padding=3, anchor='center', width=v[0])
            header_label.place(relx=v[1], rely=0.01)

        def makegentenform(self):
            entries = []
            index_y = 0.107
            for i in range(1, 11):
                label_gentenindex = ttk.Label(self.frame_genten, relief="raised", text=str(i) if i > 9 else f'  {i}',
                                              padding=3)
                input_gententensu = ttk.Entry(self.frame_genten)
                input_gentenziyu = ttk.Entry(self.frame_genten)
                input_gentennaiyo = ttk.Entry(self.frame_genten)
                label_gentenindex.place(relx=0.002, rely=index_y)
                input_gentenziyu.place(relx=0.022, rely=index_y + 0.002, width=61)
                input_gententensu.place(relx=0.08, rely=index_y + 0.002, width=345)
                input_gentennaiyo.place(relx=0.409, rely=index_y + 0.002, width=580)
                index_y = index_y + 0.084
                entries.append((input_gententensu, input_gentenziyu, input_gentennaiyo))
            return entries

        self.ents = makegentenform(self)

        # 請求理由
        self.frame_seikyuriyu = ttk.Frame(self.tab_form, relief="sunken", width=1020, height=150)
        self.label_seikyuriyu = ttk.Label(self.frame_seikyuriyu, text='請求理由')
        self.input_seikyuriyu = ScrolledText(self.frame_seikyuriyu, font=("", 10), height=9, width=140)
        self.frame_seikyuriyu.place(relx=0.01, rely=0.75)
        self.label_seikyuriyu.place(relx=0.005, rely=0.01)
        self.input_seikyuriyu.place(relx=0.005, rely=0.13)

        # 再審査結果
        self.input_receipt_num = ttk.Entry(self.tab_form)
        self.input_patient_id = ttk.Entry(self.tab_form)
        self.input_receipt_num.pack_forget()
        self.input_patient_id.pack_forget()

        # 出力ボタン
        self.btn_output = tkinter.Button(self.tab_form, bd=3, bg='moccasin', text='出力', command=self.export)
        self.btn_output.place(relx=0.81, rely=0.945, width=188, height=39)

        # 査定CSV
        self.frame_detail_sateicsv = tkinter.Frame(self.tab_satei)
        self.frame_detail_sateicsv.pack()
        self.tree_detail_sateicsv = ttk.Treeview(self.frame_detail_sateicsv, height=34)
        d = json.load(open(Path.cwd()/'システム設定'/'sateicsv_colum.json', 'r', encoding='utf8'))
        self.tree_detail_sateicsv['columns'] = tuple(list(d.keys()))
        self.tree_detail_sateicsv["show"] = "headings"
        show_colum_list = []
        for index, value in d.items():
            self.tree_detail_sateicsv.column(index, width=value['width'])
            self.tree_detail_sateicsv.heading(index, text=index)
            if value['visible']:
                show_colum_list.append(index)
        self.tree_detail_sateicsv['displaycolumns'] = tuple(show_colum_list)
        self.tree_detail_sateicsv.pack()
        hscrollbar = ttk.Scrollbar(self.frame_detail_sateicsv, orient=tkinter.HORIZONTAL,
                                   command=self.tree_detail_sateicsv.xview)
        self.tree_detail_sateicsv.configure(xscrollcommand=lambda f, l: hscrollbar.set(f, l))
        hscrollbar.pack(fill='x')
        self.btn_show_all = tkinter.Button(self.tab_satei, bd=3, bg='moccasin', text='全表示',
                                           command=self.all_show_satei)
        self.btn_show_partial = tkinter.Button(self.tab_satei, bd=3, bg='moccasin', text='部分表示',
                                               command=self.partial_show_satei)
        self.btn_show_all.place(relx=0.8, rely=0.95, width=100)
        self.btn_show_partial.place(relx=0.9, rely=0.95, width=100)

        # レセ電コード
        self.input_receipt = ScrolledText(self.tab_uke, font=("", 10), height=59, width=142)
        self.input_receipt.place(relx=0.01, rely=0.01)

    # サブ画面
    def subWindow_default(self):
        if self.sub_window is None or not self.sub_window.winfo_exists():
            # 「initial.json」から初期値読み込み
            master_dict = json.load(open(Path.cwd()/'システム設定'/'initial.json', 'r', encoding='utf8'))
            master_list = [v for k, v in master_dict.items()]
            self.sub_window = tkinter.Toplevel()
            self.sub_window.title("再審査請求書　初期値設定")
            self.sub_window.configure(bg='#f8f8ff')
            self.sub_window.geometry("500x200")
            def fetch(entries):
                result_dict = {}
                for entry in entries:
                    result_dict[entry[0]] = entry[1].get()
                writefile = open(Path.cwd() / 'システム設定' / 'initial.json', 'w', encoding='utf8')
                json.dump(result_dict, writefile, indent=2, ensure_ascii=False)
                writefile.close()
                self.insert_initial()
                self.sub_window.destroy()

            def make_initialform(self):
                entries = []
                labels = []
                for i, v in master_dict.items():
                    row = tkinter.Frame(self.sub_window, bg='GhostWhite')
                    label = tkinter.Label(row, text=i, bg='GhostWhite')
                    entry = tkinter.Entry(row, width=62)
                    entry.insert(tkinter.END, v)
                    row.pack(side=tkinter.TOP, fill=tkinter.X, padx=5, pady=1)
                    label.pack(side=tkinter.LEFT)
                    entry.pack(side=tkinter.RIGHT, fill=tkinter.X)
                    labels.append(i)
                    entries.append((i, entry))
                return entries
            ents = make_initialform(self)

            btn_save_initial = tkinter.Button(self.sub_window, text='保存', command=(lambda e=ents: fetch(e)))
            btn_save_initial.pack(pady=1)

    def subWindow_dialog(self):
        if self.sub_window is None or not self.sub_window.winfo_exists():
            # 「dialog.json」から初期値読み込み
            master_dict = json.load(open(Path.cwd() / 'システム設定' / 'dialog.json', 'r', encoding='utf8'))
            self.sub_window = tkinter.Toplevel()
            self.sub_window.title("ファイルダイアログ　初期値設定")
            self.sub_window.geometry("450x130")
            self.sub_window.configure(bg='#f8f8ff')
            frame_sateicsv_dialog = ttk.Frame(self.sub_window)
            label_sateicsv_dialog = ttk.Label(frame_sateicsv_dialog, text='査定CSV', width=8)
            self.input_sateicsv_dialog = ttk.Entry(frame_sateicsv_dialog, width=55)
            self.input_sateicsv_dialog.insert(tkinter.END, master_dict['査定CSV'])
            btn_sateicsv_dialog = tkinter.Button(frame_sateicsv_dialog, text='参照', command=lambda: self.select_folder_dialog(self.input_sateicsv_dialog))
            frame_sateicsv_dialog.pack(pady=[5, 0])
            label_sateicsv_dialog.pack(side=tkinter.LEFT)
            self.input_sateicsv_dialog.pack(side=tkinter.LEFT)
            btn_sateicsv_dialog.pack(side=tkinter.LEFT)
            frame_uke_dialog = ttk.Frame(self.sub_window)
            label_uke_dialog = ttk.Label(frame_uke_dialog, text='UKE', width=8)
            self.input_uke_dialog = ttk.Entry(frame_uke_dialog, width=55)
            self.input_uke_dialog.insert(tkinter.END, master_dict['UKE'])
            btn_uke_dialog = tkinter.Button(frame_uke_dialog, text='参照', command=lambda: self.select_folder_dialog(self.input_uke_dialog))
            frame_uke_dialog.pack()
            label_uke_dialog.pack(side=tkinter.LEFT)
            self.input_uke_dialog.pack(side=tkinter.LEFT)
            btn_uke_dialog.pack(side=tkinter.LEFT)
            frame_output_dialog = ttk.Frame(self.sub_window)
            label_output_dialog = ttk.Label(frame_output_dialog, text='保存先', width=8)
            self.input_output_dialog = ttk.Entry(frame_output_dialog, width=55)
            self.input_output_dialog.insert(tkinter.END, master_dict['保存先'])
            btn_output_dialog = tkinter.Button(frame_output_dialog, text='参照', command=lambda: self.select_folder_dialog(self.input_output_dialog))
            frame_output_dialog.pack()
            label_output_dialog.pack(side=tkinter.LEFT)
            self.input_output_dialog.pack(side=tkinter.LEFT)
            btn_output_dialog.pack(side=tkinter.LEFT)
            btn_save_dialog = tkinter.Button(self.sub_window, text='保存', command=self.update_dialog)
            btn_save_dialog.pack(pady=5)

    def select_folder_dialog(self, target_control):
        fld = filedialog.askdirectory()
        target_control.delete(0, tkinter.END)
        target_control.insert(tkinter.END, fld)
        self.sub_window.lift()

    def update_dialog(self):
        d = {}
        d['査定CSV'] = self.input_sateicsv_dialog.get()
        d['UKE'] = self.input_uke_dialog.get()
        d['保存先'] = self.input_output_dialog.get()
        writefile = open(Path.cwd()/'システム設定'/'dialog.json', 'w', encoding='utf8')
        json.dump(d, writefile, indent=2, ensure_ascii=False)
        writefile.close()
        self.insert_initial()
        self.sub_window.destroy()

    def uke_file_select(self):
        idir = json.load(open(Path.cwd()/'システム設定'/'dialog.json', 'r', encoding='utf8'))['UKE']
        filetype = [("UKE", "*.UKE")]
        file_path = tkinter.filedialog.askopenfilename(filetypes=filetype,
                                                       initialdir=idir)
        UKE_file_name = os.path.basename(file_path)
        import_path = os.path.join(import_file.uke_path(), UKE_file_name)
        import_file.delete_file(import_file.uke_path())  # 対象フォルダを空にする
        self.input_import_uke.delete(0, tkinter.END)
        shutil.copyfile(file_path, import_path)
        self.input_import_uke.insert(tkinter.END, file_path)
        if import_file.chk():
            self.btn_import.pack(pady=[100, 10], ipady=40)
        else:
            self.btn_import.pack_forget()

    def satei_file_select(self):
        idir = json.load(open(Path.cwd() / 'システム設定' / 'dialog.json', 'r', encoding='utf8'))['査定CSV']
        filetype = [("csv", "*.csv")]
        file_path = tkinter.filedialog.askopenfilename(filetypes=filetype,
                                                       initialdir=idir)
        satei_file_name = os.path.basename(file_path)
        import_path = os.path.join(import_file.satei_path(), satei_file_name)
        import_file.delete_file(import_file.satei_path())  # 対象フォルダを空にする
        self.input_import_sateicsv.delete(0, tkinter.END)
        shutil.copyfile(file_path, import_path)
        self.input_import_sateicsv.insert(tkinter.END, file_path)
        if import_file.chk():
            self.btn_import.pack(pady=[100, 10], ipady=40)
        else:
            self.btn_import.pack_forget()

    # 再審査請求書ボタンをクリックしたとき
    def read_files(self):
        self.insert_patient_list()
        self.notebook.select(self.tab_saishinsaseikyusyo)
        today_list = convert_to_wareki.convert_to_wareki().split()
        self.v_year.set(today_list[0])
        self.input_year.delete(0, tkinter.END)
        self.input_year.insert(tkinter.END, today_list[1])
        self.input_month.delete(0, tkinter.END)
        self.input_month.insert(tkinter.END, today_list[3])
        self.input_day.delete(0, tkinter.END)
        self.input_day.insert(tkinter.END, today_list[5])
        tensuhyokubun_t, meisaisyokubun1_t, meisaisyokubun2_t, saishinsasyubetsu_t, iryoukikan, seikyunengetsu = import_file.get_ir_info()
        self.t_tensuhyokubun = tensuhyokubun_t  # 点数表区分
        self.cb_tensuhyokubun['values'] = self.t_tensuhyokubun
        self.v_tensuhyokubun.set(self.t_tensuhyokubun[0])  # 点数表区分
        self.t_meisaisyokubun1 = meisaisyokubun1_t  # 明細書区分
        self.cb_meisaisyokubun1['value'] = self.t_meisaisyokubun1
        self.t_meisaisyokubun2 = meisaisyokubun2_t  # 明細書区分２
        self.cb_meisaisyokubun2['value'] = self.t_meisaisyokubun2
        self.t_saishinsasyubetsu = saishinsasyubetsu_t  # 再審査種別
        self.cb_saishinsasyubetsu['value'] = self.t_saishinsasyubetsu
        self.input_iryokikanbangou.delete(0, tkinter.END)  # 医療機関
        self.input_iryokikanbangou.insert(tkinter.END, iryoukikan)
        self.input_seikyuunengetsu.delete(0, tkinter.END)  # 請求年月
        self.input_seikyuunengetsu.insert(tkinter.END, seikyunengetsu)
        self.insert_initial()

    # 初期値を代入
    def insert_initial(self):
        d = json.load(open(Path.cwd()/'システム設定'/'initial.json', 'r', encoding='utf8'))
        self.input_shibu.delete(0, tkinter.END)
        self.input_address.delete(0, tkinter.END)
        self.input_hospitalname.delete(0, tkinter.END)
        self.input_founder.delete(0, tkinter.END)
        self.input_tel.delete(0, tkinter.END)
        if d['支払基金']:
            self.input_shibu.insert(tkinter.END, d['支払基金'])
        if d['保険医療機関の住所']:
            self.input_address.insert(tkinter.END, d['保険医療機関の住所'])
        if d['保険医療機関の名称']:
            self.input_hospitalname.insert(tkinter.END, d['保険医療機関の名称'])
        if d['開設者氏名']:
            self.input_founder.insert(tkinter.END, d['開設者氏名'])
        if d['電話番号']:
            self.input_tel.insert(tkinter.END, d['電話番号'])

    def insert_patient_list(self):
        # treeviewを削除する
        for item in self.tree_patient.get_children():
            self.tree_patient.delete(item)
        self.tree_patient.tag_configure('oddrow', background="white")
        self.tree_patient.tag_configure('evenrow', background="lavender")
        global count
        count = 0
        regex = r'[0-9]{6,10}'
        for row in import_file.get_satei_list():
            text_list = row.split(',')
            if len(text_list) == 30:
                if text_list[18] and re.match(regex, text_list[18]) and (text_list[0] == 3 or text_list[0] == '3'):
                    receipt_id = str(int(text_list[2]))
                    patient_id = text_list[18]
                    patient_name = text_list[17]
                    self.add_row(receipt_id, patient_id, patient_name, count)
                    count += 1

    def search(self):
        target_id = self.patientid_input.get()
        # treeviewを削除する
        for item in self.tree_patient.get_children():
            self.tree_patient.delete(item)

        self.tree_patient.tag_configure('oddrow', background="white")
        self.tree_patient.tag_configure('evenrow', background="lavender")
        global count
        count = 0
        regex = r'[0-9]{6,10}'
        for row in import_file.get_satei_list():
            row_list = row.split(',')
            if len(row_list) == 30:
                if row_list[18] and re.match(regex, row_list[18]) and (row_list[0] == 3 or row_list[0] == '3'):
                    receipt_id = str(int(row_list[2]))
                    patient_id = row_list[18]
                    patient_name = row_list[17]
                    if target_id:
                        if target_id in patient_id:
                            self.add_row(receipt_id, patient_id, patient_name, count)
                    elif not target_id:
                        self.add_row(receipt_id, patient_id, patient_name, count)
                    count += 1

    def add_row(self, ri, pid, pname, c):
        if c % 2 == 0:
            self.tree_patient.insert(parent="", index="end", values=(ri, pid, pname), tags=('oddrow',))
        else:
            self.tree_patient.insert(parent="", index="end", values=(ri, pid, pname), tags=('evenrow',))

    def all_show_satei(self):
        d = json.load(open(Path.cwd()/'システム設定'/'sateicsv_colum.json', 'r', encoding='utf8'))
        self.tree_detail_sateicsv['displaycolumns'] = tuple(list(d.keys()))

    def partial_show_satei(self):
        d = json.load(open(Path.cwd()/'システム設定'/'sateicsv_colum.json', 'r', encoding='utf8'))
        l = [i for i, v in d.items() if v['visible']]
        self.tree_detail_sateicsv['displaycolumns'] = tuple(l)

    # 患者ごとに変わる入力欄
    def patient_form_list(self):
        input_list = [self.input_shinryounengetsu, self.input_yakkyokucode,
                      self.input_todohukenmei, self.input_yakkyokuname,
                      self.input_hokensyabango, self.input_kigoubango, self.input_kouhihutanbango,
                      self.input_jukyusyabango,
                      self.input_kouhihutanbango2, self.input_jukyusyabango2, self.input_patient_cname,
                      self.input_patient_name,
                      self.input_seinengappi, self.input_tousyoseikyutensu, self.input_itibuhutankin,
                      self.input_tousyoseikyuhi,
                      self.input_hyojunhutangaku, self.input_torisageriyu, self.input_receipt_num, self.input_patient_id]
        cb_list = [self.cb_shinryouka, self.cb_meisaisyokubun1, self.cb_meisaisyokubun2, self.cb_saishinsasyubetsu, self.cb_seibetsu,
                   self.cb_utsushinoumu]
        text_area_list = [self.input_seikyuriyu]

        return (input_list, cb_list, text_area_list)

    # 患者を選択したとき
    def select_record(self, event):
        input_list, cb_list, text_area_list = self.patient_form_list()
        for i in input_list:
            i.delete(0, 'end')
        for i in cb_list:
            i.set('')
        for i in text_area_list:
            i.delete('1.0', 'end')
        d = json.load(open(Path.cwd() / 'システム設定' / 'initial.json', 'r', encoding='utf8'))
        try:
            self.v_utsushinoumu.set(self.t_utsushinoumu[int(d["写しの有無"]) - 1])
        except:
            pass
        try:
            self.cb_saishinsasyubetsu.set(self.t_saishinsasyubetsu[int(d["再審査等対象種別"]) - 1])
        except:
            pass
        record_id = self.tree_patient.focus()
        # 選択行のレコードを取得
        record_value = self.tree_patient.item(record_id, 'values')
        self.input_receipt_num.insert(tkinter.END, record_value[0])
        self.input_patient_id.insert(tkinter.END, record_value[1])
        self.input_receipt.delete("1.0", "end")
        # 選択した患者のレコードをレセ電コードタブに入れる
        receipt_value, index_list = import_file.patient_detail(record_value[0], record_value[1])
        self.input_receipt.insert(tkinter.END, receipt_value)

        # レセ電コードタブの内容を再審査請求書に反映させる
        master_dept = import_file.dept()
        detail = self.input_receipt.get("1.0", "end-1c")
        detail_list = detail.split('\n')
        config = import_file.get_uke_config()
        write_row = False
        receipt_type = ''
        dekidaka_hyozyunhutangaku = 0
        for r in detail_list:
            r = r.replace('\n', '')
            r_list = r.split(',')
            if r_list[0] == 'RE':
                if len(r_list) == 30:
                    if int(r_list[18]) < 2:
                        write_row = True
                    else:
                        write_row = False
                else:
                    write_row = True
            if write_row:
                if r_list[0] == 'RE':
                    # DPCか出来高か
                    if not receipt_type:
                        receipt_type = '出来高' if len(r_list) == 38 else 'DPC'
                    # 診療科
                    self.cb_shinryouka.set(master_dept[r_list[config['RE']['診療科名']]])
                    # 診療年月
                    self.input_shinryounengetsu.insert(tkinter.END,
                                                       convert_to_wareki.year_conversion(r_list[config['RE']['診療年月']]))
                    # 患者情報
                    self.input_patient_cname.insert(tkinter.END, r_list[config['RE']['患者カナ氏名']])
                    self.input_patient_name.insert(tkinter.END, r_list[config['RE']['患者氏名']])
                    if r_list[config['RE']['性別']] == '1':
                        self.v_seibetsu.set('1.男')  # 性別
                    elif r_list[config['RE']['性別']] == '2':
                        self.v_seibetsu.set('2.女')
                    birthday = convert_to_wareki.convert_to_wareki(r_list[config['RE']['生年月日']])
                    self.input_seinengappi.insert(tkinter.END, birthday)
                    # レセプト種別
                    self.v_meisaisyokubun2.set(
                        self.t_meisaisyokubun2[int(str(int(r_list[config['RE']['レセプト種別']][-2:]) - 1)[-1:])])
                elif r_list[0] == 'HO':
                    # 保険者番号
                    hokensyabango = r_list[config['HO']['保険者番号']].replace(' ', '0')
                    self.input_hokensyabango.insert(tkinter.END, hokensyabango)
                    # 明細書区分1
                    if not self.v_meisaisyokubun1.get():
                        if hokensyabango[:2] == '39':
                            self.v_meisaisyokubun1.set(self.t_meisaisyokubun1[1])
                        elif hokensyabango[:2] == '67':
                            self.v_meisaisyokubun1.set(self.t_meisaisyokubun1[2])
                        else:
                            self.v_meisaisyokubun1.set(self.t_meisaisyokubun1[0])
                    # 記号・番号
                    self.input_kigoubango.insert(tkinter.END,
                                                 r_list[config['HO']['記号番号'][0]] + '・' + r_list[
                                                     config['HO']['記号番号'][1]])
                    self.input_tousyoseikyutensu.insert(tkinter.END, r_list[config['HO']['合計点数']])
                    # 一部負担金
                    if not self.input_itibuhutankin.get():
                        self.input_itibuhutankin.insert(tkinter.END, r_list[config['HO']['一部負担金']])
                    # 標準負担額
                    if not self.input_hyojunhutangaku.get() and config['HO']['標準負担額']:
                        self.input_hyojunhutangaku.insert(tkinter.END, r_list[config['HO']['標準負担額']])
                    # 食事・生活請求金額
                    if not self.input_tousyoseikyuhi.get():
                        self.input_tousyoseikyuhi.insert(tkinter.END, r_list[config['HO']['食事生活療養']])
                elif r_list[0] == 'KO':
                    # 公費負担番号
                    if self.input_kouhihutanbango.get():
                        self.input_kouhihutanbango2.insert(tkinter.END, r_list[config['KO']['負担者番号']])
                    else:
                        self.input_kouhihutanbango.insert(tkinter.END, r_list[config['KO']['負担者番号']])
                        # 明細書区分1
                        if r_list[config['KO']['負担者番号']][:2] == '12':
                            self.v_meisaisyokubun1.set(self.t_meisaisyokubun1[0])
                    # 受給者番号
                    if self.input_jukyusyabango.get():
                        self.input_jukyusyabango2.insert(tkinter.END, r_list[config['KO']['受給者番号']])
                    else:
                        self.input_jukyusyabango.insert(tkinter.END, r_list[config['KO']['受給者番号']])
                    # 一部負担金、標準負担額、食事・生活請求金額
                    if not 'HO' in index_list:
                        # 一部負担金
                        if not self.input_itibuhutankin.get() and config['KO']['一部負担金']:
                            self.input_itibuhutankin.insert(tkinter.END, r_list[config['KO']['一部負担金']])
                        # 標準負担額
                        if not self.input_hyojunhutangaku.get() and config['KO']['標準負担額']:
                            self.input_hyojunhutangaku.insert(tkinter.END, r_list[config['KO']['標準負担額']])
                        # 食事・生活請求金額
                        if not self.input_tousyoseikyuhi.get():
                            self.input_tousyoseikyuhi.insert(tkinter.END, r_list[config['KO']['食事生活療養']])
                elif r_list[0] == 'SI':
                    if receipt_type == '出来高':
                        if r_list[3] in import_file.get_hyozyunhutangaku_codelist():
                            try:
                                dekidaka_hyozyunhutangaku += int(r_list[config['SI']['点数']]) * int(r_list[config['SI']['回数']])
                            except:
                                print('出来高の標準負担額の計算が出来ませんでした。')
        if dekidaka_hyozyunhutangaku > 0:
            self.input_hyojunhutangaku.insert(tkinter.END, str(dekidaka_hyozyunhutangaku))
        # 査定CSVを読み込む
        class SateiDetail:
            def __init__(self):
                self.TENSU = ''
                self.ZIYUKASYO = ''
                self.NAIYO = ''

        target_row = False
        self.tree_detail_sateicsv.tag_configure('white_row', background="white")
        # 査定CSVタブのtreeviewを削除する
        for item in self.tree_detail_sateicsv.get_children():
            self.tree_detail_sateicsv.delete(item)
        wn = 0
        for entrie in self.ents:
            entrie[0].delete(0, 'end')
            entrie[1].delete(0, 'end')
            entrie[2].delete(0, 'end')
        zougentensu = ''
        kasyo1 = ''
        kasyo2 = ''
        naiyo = ''
        d = json.load(open(Path.cwd()/'システム設定'/'sateicsv_colum.json', 'r', encoding='utf8'))
        pid_col = int(d['カルテ番号等']['i'])-1
        zougentensu_col = int(d['増減点数（金額）']['i'])-1
        kasyo1_col = int(d['箇所１']['i'])-1
        kasyo2_col = int(d['箇所２']['i'])-1
        sateigonaiyo_col = int(d['補正・査定後内容']['i'])-1
        seikyunaiyo_col = int(d['請求内容']['i'])-1
        ziyu_col = int(d['事由']['i'])-1
        results = []
        for i in import_file.read_sateicsv():
            row_list = i.split(',')
            if len(row_list) == 30:
                if row_list[pid_col]:
                    if self.input_patient_id.get() == row_list[pid_col] and self.input_hokensyabango.get() == row_list[
                        7].replace(' ', '0'):
                        target_row = True
                    else:
                        target_row = False
                if target_row:
                    self.tree_detail_sateicsv.insert(parent="", index="end",
                                                     values=tuple(row_list), tags=('white_row',))
                    if wn < 10:
                        if row_list[zougentensu_col]:
                            zougentensu = row_list[zougentensu_col]
                        if row_list[kasyo1_col]:
                            kasyo1 = row_list[kasyo1_col]
                        if row_list[kasyo2_col] and (row_list[0] == "3" or row_list[0] == 3 or row_list[0] == "4" or row_list[0] == 4):
                            kasyo2 = row_list[kasyo2_col]
                        if row_list[ziyu_col]:
                            if row_list[sateigonaiyo_col]:
                                naiyo = f'{row_list[seikyunaiyo_col]} → {row_list[sateigonaiyo_col]}'
                            else:
                                naiyo = row_list[seikyunaiyo_col]
                            result = SateiDetail()
                            result.TENSU = zougentensu
                            result.ZIYUKASYO = f'{row_list[ziyu_col]} {kasyo1} {kasyo2}'
                            result.NAIYO = naiyo
                            results.append(result)
                            wn += 1
                else:
                    wn = 0
        wn = 0
        for row in results:
            self.ents[wn][1].insert(tkinter.END, row.TENSU)
            self.ents[wn][0].insert(tkinter.END, row.ZIYUKASYO)
            self.ents[wn][2].insert(tkinter.END, row.NAIYO)
            wn += 1

    # 出力ボタンをクリックしたとき
    def export(self):
        if self.input_patient_name.get():
            template_dir = os.path.join(os.getcwd(), '雛型')
            out_path = json.load(open(Path.cwd() / 'システム設定' / 'dialog.json', 'r', encoding='utf8'))['保存先']
            genten_values1 = []
            genten_values2 = []
            genten_num = 0
            for value in self.ents:
                if value[0].get() or value[1].get() or value[2].get():
                    genten_num += 1
                if genten_num > 5:
                    genten_values2.append((value[1].get(), value[0].get(), value[2].get()))
                else:
                    genten_values1.append((value[1].get(), value[0].get(), value[2].get()))
            if not os.path.isdir(out_path):
                os.makedirs(out_path)
            syaho_or_kokuho = import_file.syaho_or_kokuho()
            if syaho_or_kokuho == 2:
                template = os.path.join(template_dir, '再審査請求書.xlsx')
                json_open = open(Path.cwd() / '雛型' / 'shape_map_kokuho.json', 'r', encoding='utf8')
                shape_map_dict = json.load(json_open)
                # Excelを開く
                xl = win32com.client.Dispatch('Excel.Application')
                xl.Visible = True
                wb = xl.Workbooks.Open(template)
                ws = wb.Worksheets(1)
                ws.Activate()
                system_date = f"{self.v_year.get()} {self.input_year.get()} 年  {self.input_month.get()} 月  {self.input_day.get()} 日"
                ws.Range("AM2").Value = system_date  # システム日付
                ws.Range("A3").Value = f'　　愛知県国民健康保険診療報酬審査委員会{self.input_shibu.get()}御中　　　　　　　　　'  # 支部
                ws.Range("AF4").Value = self.input_address.get() if self.input_address.get() else ''  # 保険医療機関の情報
                ws.Range(
                    "AF5").Value = self.input_hospitalname.get() if self.input_hospitalname.get() else ''  # 保険医療機関の情報
                ws.Range("AF6").Value = self.input_founder.get() if self.input_founder.get() else ''  # 保険医療機関の情報
                ws.Range("AF7").Value = self.input_tel.get() if self.input_tel.get() else ''  # 保険医療機関の情報
                if self.v_tensuhyokubun.get():
                    ws.Shapes(shape_map_dict['点数表'][0]).Visible = True
                    ws.Shapes(shape_map_dict['点数表'][0]).Left = \
                        shape_map_dict['点数表'][self.cb_tensuhyokubun['values'].index(self.v_tensuhyokubun.get()) + 1][0]
                    ws.Shapes(shape_map_dict['点数表'][0]).Top = \
                        shape_map_dict['点数表'][self.cb_tensuhyokubun['values'].index(self.v_tensuhyokubun.get()) + 1][1]
                else:
                    ws.Shapes(shape_map_dict['点数表'][0]).Visible = False
                ws.Range("U10").Value = self.input_iryokikanbangou.get()  # 医療機関
                ws.Range('AS10').value = self.v_shinryouka.get()  # 診療科
                ws.Range('G11').value = self.input_shinryounengetsu.get()  # 診療年月
                ws.Range('P11').value = self.input_seikyuunengetsu.get()  # 診療年月
                if self.v_meisaisyokubun1.get():
                    ws.Shapes(shape_map_dict['明細書区分1'][0]).Visible = True
                    ws.Shapes(shape_map_dict['明細書区分1'][0]).Left = \
                        shape_map_dict['明細書区分1'][
                            self.cb_meisaisyokubun1['values'].index(self.v_meisaisyokubun1.get()) + 1][0]
                    ws.Shapes(shape_map_dict['明細書区分1'][0]).Top = \
                        shape_map_dict['明細書区分1'][
                            self.cb_meisaisyokubun1['values'].index(self.v_meisaisyokubun1.get()) + 1][1]
                else:
                    ws.Shapes(shape_map_dict['明細書区分1'][0]).Visible = False
                if self.v_meisaisyokubun2.get():
                    ws.Shapes(shape_map_dict['明細書区分2'][0]).Visible = True
                    ws.Shapes(shape_map_dict['明細書区分2'][0]).Left = \
                        shape_map_dict['明細書区分2'][
                            self.cb_meisaisyokubun2['values'].index(self.v_meisaisyokubun2.get()) + 1][0]
                    ws.Shapes(shape_map_dict['明細書区分2'][0]).Top = \
                        shape_map_dict['明細書区分2'][
                            self.cb_meisaisyokubun2['values'].index(self.v_meisaisyokubun2.get()) + 1][1]
                else:
                    ws.Shapes(shape_map_dict['明細書区分2'][0]).Visible = False
                if self.v_saishinsasyubetsu.get():
                    ws.Shapes(shape_map_dict['再審査等対象種別'][0]).Visible = True
                    ws.Shapes(shape_map_dict['再審査等対象種別'][0]).Left = shape_map_dict['再審査等対象種別'][
                        self.cb_saishinsasyubetsu['values'].index(self.v_saishinsasyubetsu.get()) + 1][0]
                    ws.Shapes(shape_map_dict['再審査等対象種別'][0]).Top = shape_map_dict['再審査等対象種別'][
                        self.cb_saishinsasyubetsu['values'].index(self.v_saishinsasyubetsu.get()) + 1][1]
                else:
                    ws.Shapes(shape_map_dict['再審査等対象種別'][0]).Visible = False

                ws.Range('W12').Value = f'（都道府県名　{self.input_todohukenmei.get()}　） {self.input_yakkyokucode.get()}'
                ws.Range('W13').Value = self.input_yakkyokuname.get()

                def write_num(row, n, txt):
                    for i in txt:
                        ws.Cells(row, n).value = i
                        n += 2

                write_num(14, 12, self.input_hokensyabango.get())  # 保険者番号
                ws.Range('AL14').value = self.input_kigoubango.get()  # 記号番号
                write_num(15, 12, self.input_kouhihutanbango.get())  # 公費負担番号
                write_num(15, 38, self.input_jukyusyabango.get())  # 受給者番号
                write_num(16, 12, self.input_kouhihutanbango2.get())  # 公費負担番号
                write_num(16, 38, self.input_jukyusyabango2.get())
                ws.Range('L17').value = self.input_patient_cname.get()
                ws.Range('L18').value = self.input_patient_name.get()
                ws.Range('AL17').value = self.input_tousyoseikyutensu.get() + '点'
                ws.Range('AL19').value = self.input_tousyoseikyuhi.get() + '円'
                if self.v_seibetsu.get():
                    ws.Shapes(shape_map_dict['性別'][0]).Visible = True
                    ws.Shapes(shape_map_dict['性別'][0]).Left = \
                    shape_map_dict['性別'][self.cb_seibetsu['values'].index(self.v_seibetsu.get()) + 1][0]
                    ws.Shapes(shape_map_dict['性別'][0]).Top = \
                    shape_map_dict['性別'][self.cb_seibetsu['values'].index(self.v_seibetsu.get()) + 1][1]
                else:
                    ws.Shapes(shape_map_dict['性別'][0]).Visible = False
                birthday = self.input_seinengappi.get().split()
                map_list = shape_map_dict[birthday[0]]
                ws.Shapes(shape_map_dict['年号']).Visible = True
                ws.Shapes(shape_map_dict['年号']).Left = shape_map_dict[birthday[0]][0]
                ws.Shapes(shape_map_dict['年号']).Top = shape_map_dict[birthday[0]][1]
                ws.Range('AL21').value = '1明　2大　3昭　4平　5令\n  ' + birthday[1] + ' ．' + birthday[3] + ' . ' + birthday[
                    5] + ' 生'
                for row, value in zip(range(23, 28), self.ents):
                    ws.Cells(row, 5).value = value[1].get()
                    ws.Cells(row, 14).value = value[0].get()
                    ws.Cells(row, 24).value = value[2].get()
                ws.Range("J29").value = self.input_seikyuriyu.get("1.0", "end-1c")
                row = 33
                for value in self.ents:
                    if row < 53:
                        ws.Cells(row, 5).value = value[1].get()
                        ws.Cells(row, 14).value = value[0].get()
                        ws.Cells(row, 24).value = value[2].get()
                        row += 1
                if genten_num > 5:
                    wb.Sheets.Copy(None, After=ws)
                    ws2 = wb.Worksheets(2)
                    for row, value in zip(range(23, 28), genten_values2):
                        ws2.Cells(row, 5).value = value[0]
                        ws2.Cells(row, 14).value = value[1]
                        ws2.Cells(row, 24).value = value[2]
                ws.Activate()

                out_file = self.input_patient_id.get() + '_再審査請求書.xlsx'
                out_path = os.path.abspath(out_path)
                output_file = os.path.join(out_path, out_file)
                wb.SaveAs(output_file)
            elif syaho_or_kokuho == 1:
                template = os.path.join(template_dir, '再審査請求書.docx')
                Application = win32com.client.Dispatch("Word.Application")
                Application.Visible = True
                doc = Application.Documents.Open(FileName=template, ConfirmConversions=None)
                json_open = open(Path.cwd() / '雛型' / 'shape_map_syaho.json', 'r', encoding='utf8')
                shape_map_dict = json.load(json_open)

                # システム日付
                system_date = [self.input_year.get(), self.input_month.get(), self.input_day.get()]
                doc.Range(3, 5).Text = system_date[0] if len(system_date[0]) > 1 else f' {system_date[0]}'
                doc.Range(6, 8).Text = system_date[1] if len(system_date[1]) > 1 else f' {system_date[1]}'
                doc.Range(9, 11).Text = system_date[2] if len(system_date[2]) > 1 else f' {system_date[2]}'

                # 支部
                shibu = self.input_shibu.get()
                doc.Range(33, 33 + len(shibu)).Text = shibu

                # 保険医療機関の情報
                hos_info = ''
                hos_info += self.input_address.get() + '\n' if self.input_address.get() else '\n'
                hos_info += self.input_hospitalname.get() + '\n' if self.input_hospitalname.get() else '\n'
                hos_info += self.input_founder.get() + '\n' if self.input_founder.get() else '\n'
                hos_info += self.input_tel.get() if self.input_tel.get() else ''
                doc.Shapes(shape_map_dict['保険医療機関の情報']).TextFrame.TextRange.Text = hos_info

                # 再審査or取り下げ
                if self.v_saishinsa_torisage.get():
                    doc.Shapes(shape_map_dict['再審査or取下'][0]).Left = shape_map_dict['再審査or取下'][
                        self.cb_saishinsa_torisage['values'].index(self.v_saishinsa_torisage.get()) + 1][0]
                    doc.Shapes(shape_map_dict['再審査or取下'][0]).Top = shape_map_dict['再審査or取下'][
                        self.cb_saishinsa_torisage['values'].index(self.v_saishinsa_torisage.get()) + 1][1]
                else:
                    doc.Shapes(shape_map_dict['再審査or取下'][0]).Visible = False

                # 点数表
                if self.v_tensuhyokubun.get():
                    doc.Shapes(shape_map_dict['点数表'][0]).Left = \
                        shape_map_dict['点数表'][self.cb_tensuhyokubun['values'].index(self.v_tensuhyokubun.get()) + 1][0]
                    doc.Shapes(shape_map_dict['点数表'][0]).Top = \
                        shape_map_dict['点数表'][self.cb_tensuhyokubun['values'].index(self.v_tensuhyokubun.get()) + 1][1]
                else:
                    doc.Shapes(shape_map_dict['点数表'][0]).Visible = False
                # 明細書区分
                if self.v_meisaisyokubun1.get():
                    doc.Shapes(shape_map_dict['明細書区分1'][0]).Left = \
                        shape_map_dict['明細書区分1'][
                            self.cb_meisaisyokubun1['values'].index(self.v_meisaisyokubun1.get()) + 1][0]
                    doc.Shapes(shape_map_dict['明細書区分1'][0]).Top = \
                        shape_map_dict['明細書区分1'][
                            self.cb_meisaisyokubun1['values'].index(self.v_meisaisyokubun1.get()) + 1][1]
                else:
                    doc.Shapes(shape_map_dict['明細書区分1'][0]).Visible = False
                if self.v_meisaisyokubun2.get():
                    doc.Shapes(shape_map_dict['明細書区分2'][0]).Visible = True
                    doc.Shapes(shape_map_dict['明細書区分2'][0]).Left = \
                        shape_map_dict['明細書区分2'][
                            self.cb_meisaisyokubun2['values'].index(self.v_meisaisyokubun2.get()) + 1][0]
                    doc.Shapes(shape_map_dict['明細書区分2'][0]).Top = \
                        shape_map_dict['明細書区分2'][
                            self.cb_meisaisyokubun2['values'].index(self.v_meisaisyokubun2.get()) + 1][1]
                    doc.Shapes(shape_map_dict['明細書区分2'][0]).Width = \
                        shape_map_dict['明細書区分2'][
                            self.cb_meisaisyokubun2['values'].index(self.v_meisaisyokubun2.get()) + 1][2]
                else:
                    doc.Shapes(shape_map_dict['明細書区分2'][0]).Visible = False
                # 再審査等対象種別
                if self.v_saishinsasyubetsu.get():
                    doc.Shapes(shape_map_dict['再審査等対象種別'][0]).Left = shape_map_dict['再審査等対象種別'][
                        self.cb_saishinsasyubetsu['values'].index(self.v_saishinsasyubetsu.get()) + 1][0]
                    doc.Shapes(shape_map_dict['再審査等対象種別'][0]).Top = shape_map_dict['再審査等対象種別'][
                        self.cb_saishinsasyubetsu['values'].index(self.v_saishinsasyubetsu.get()) + 1][1]
                else:
                    doc.Shapes(shape_map_dict['再審査等対象種別'][0]).Visible = False
                doc.Tables(1).Cell(Row=1, Column=5).Range.text = self.input_iryokikanbangou.get()
                doc.Tables(1).Cell(Row=1, Column=7).Range.text = self.v_shinryouka.get()
                doc.Tables(1).Cell(Row=2, Column=3).Range.text = self.input_shinryounengetsu.get()
                doc.Tables(1).Cell(Row=2, Column=5).Range.text = self.input_seikyuunengetsu.get()

                # 薬局コード
                if self.input_todohukenmei.get() or self.input_yakkyokucode.get():
                    doc.Tables(1).Cell(Row=3,
                                       Column=4).Range.text = f'(       {self.input_todohukenmei.get()} )  {self.input_yakkyokucode.get()}'

                def write_num_docx(row, start, end, bango_list):
                    for col, value in zip(range(start, end), bango_list):
                        doc.Tables(1).Cell(Row=row, Column=col).Range.text = value

                write_num_docx(5, 3, 11, list(self.input_hokensyabango.get()))
                doc.Tables(1).Cell(Row=5, Column=12).Range.text = self.input_kigoubango.get()
                write_num_docx(6, 3, 11, list(self.input_kouhihutanbango.get()))
                write_num_docx(6, 12, 19, list(self.input_jukyusyabango.get()))
                doc.Tables(1).Cell(Row=7, Column=3).Range.text = self.input_patient_cname.get()
                doc.Tables(1).Cell(Row=8, Column=3).Range.text = self.input_patient_name.get()

                def write_birthday(year, month, day):
                    b = doc.Tables(1).Cell(Row=8, Column=4).Range.Text
                    bl = list(b)

                    def make_newlist(cell_num, value, l):
                        # 年
                        if int(value) < 10:
                            l[cell_num] = value
                        else:
                            l[cell_num - 1] = value[0]
                            l[cell_num] = value[1]

                    make_newlist(20, year, bl)
                    make_newlist(23, month, bl)
                    make_newlist(26, day, bl)

                    b = ''.join(bl)
                    doc.Tables(1).Cell(Row=8, Column=4).Range.Text = b[:-2]

                patient_birthday_list = self.input_seinengappi.get().split()
                write_birthday(patient_birthday_list[1], patient_birthday_list[3], patient_birthday_list[5])

                # 写しの有無
                if self.v_utsushinoumu.get():
                    doc.Shapes(shape_map_dict['写しの有無'][0]).Left = shape_map_dict['写しの有無'][
                        self.cb_utsushinoumu['values'].index(self.v_utsushinoumu.get()) + 1][0]
                    doc.Shapes(shape_map_dict['写しの有無'][0]).Top = shape_map_dict['写しの有無'][
                        self.cb_utsushinoumu['values'].index(self.v_utsushinoumu.get()) + 1][1]
                else:
                    doc.Shapes(shape_map_dict['写しの有無'][0]).Visible = False

                doc.Tables(1).Cell(Row=9, Column=3).Range.text = self.input_tousyoseikyutensu.get() + '点'
                doc.Tables(1).Cell(Row=9, Column=6).Range.text = self.input_itibuhutankin.get() + '円'
                doc.Tables(1).Cell(Row=10, Column=3).Range.text = self.input_tousyoseikyuhi.get() + '円'
                doc.Tables(1).Cell(Row=10, Column=5).Range.text = self.input_hyojunhutangaku.get() + '円'
                doc.Tables(1).Cell(Row=10, Column=7).Range.text = self.input_torisageriyu.get()
                for row, value in zip(range(12, 17), self.ents):
                    doc.Tables(1).Cell(Row=row, Column=3).Range.text = value[1].get()
                    doc.Tables(1).Cell(Row=row, Column=4).Range.text = value[0].get()
                    doc.Tables(1).Cell(Row=row, Column=5).Range.text = value[2].get()

                # 請求理由
                doc.Tables(1).Cell(Row=17, Column=1).Range.text = '\n請求理由\n\n' + self.input_seikyuriyu.get("1.0",
                                                                                                           "end-1c")
                doc.Shapes(shape_map_dict['再審査の結果'][0]).Left = shape_map_dict['再審査の結果'][1][0]
                doc.Shapes(shape_map_dict['再審査の結果'][0]).Top = shape_map_dict['再審査の結果'][1][1]

                row = 1
                for value in self.ents:
                    if row < 29:
                        doc.Tables(2).Cell(Row=row, Column=1).Range.text = value[1].get()
                        doc.Tables(2).Cell(Row=row, Column=2).Range.text = value[0].get()
                        doc.Tables(2).Cell(Row=row, Column=3).Range.text = value[2].get()
                        row += 1

                out_file = self.input_patient_id.get() + '_再審査請求書.docx'
                output_file = os.path.join(out_path, out_file)
                if os.path.isfile(output_file):
                    os.remove(output_file)
                doc.SaveAs2(output_file)


root = tkinter.Tk()
root.title('再審査請求書作成')
root.geometry("1300x900+0+0")
root.resizable(0, 0)
import_file.delete()
app = Application(root=root)
### メニューバー作成
menubar = tkinter.Menu(master=root)

### 編集メニュー作成
setmenu = tkinter.Menu(master=menubar, tearoff=0)
setmenu.add_command(label="再審査請求書　初期値設定", command=app.subWindow_default)
setmenu.add_command(label="ファイルダイアログ　初期値設定", command=app.subWindow_dialog)

### 各メニューを設定
menubar.add_cascade(label="設定", menu=setmenu)

### メニューバー配置
root.config(menu=menubar)

app.mainloop()
