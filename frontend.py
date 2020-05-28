import tkinter
from tkinter import ttk
from tkinter import *
from tkinter import messagebox as mb
from pyautogui import press
import backend

class Main_window(tkinter.Frame):

    def __init__(self, root):
        super().__init__(root)
        self.menu_bar()
        self.init_main()
            

    def menu_bar(self):
        #   ВЕРХНЄ МЕНЮ(ЗВІТНІСТЬ, СПРАВКА і.т.д)
        bar_menu = Menu(root)
        root.configure(menu=bar_menu)
        root.option_add("*Font", ('helvetica', 12))
        report_menu = tkinter.Menu(bar_menu, tearoff=0)
        report_menu.add_command(label='Формування районного звіту')
        report_menu.add_command(label='Формування порівняльного звіту')
        report_menu.add_command(label='Формування форми №14')
        bar_menu.add_cascade(label='Вибірка даних', menu=report_menu)
        helpmenu = tkinter.Menu(bar_menu, tearoff=0)
        helpmenu.add_command(label='Допомога')
        helpmenu.add_command(label='Про програму')
        bar_menu.add_cascade(label='Справка', menu=helpmenu)


    def init_main(self):
        """МЕНЮ З ЯРЛИКАМИ"""
        self.add_img = tkinter.PhotoImage(file='add.png')
        btn_open_dialog = tkinter.Button(text='Додати', bd=2, command=self.add_dialog, image=self.add_img)  #bg='white', 
        btn_open_dialog.place(x=200, y=300)
        self.search_img = tkinter.PhotoImage(file='search.png')
        btn_search_dialog = tkinter.Button(text='Пошук', bd=2, command=self.search_dialog, image=self.search_img)
        btn_search_dialog.place(x=500, y=300)
        self.edit_img = tkinter.PhotoImage(file='edit.png')
        btn_edit_dialog = tkinter.Button(text='Редагувати', bd=2, command=self.edit_dialog, image=self.edit_img)
        btn_edit_dialog.place(x=800, y=300)
        self.delete_img = tkinter.PhotoImage(file='delete.png')
        btn_delete_dialog = tkinter.Button(text='Видалити', bd=2, command=self.delete_dialog, image=self.delete_img)
        btn_delete_dialog.place(x=1100, y=300)
        self.quit_img = tkinter.PhotoImage(file='quit.png')
        btn_quit_dialog = tkinter.Button(text='Вийти', bd=2, command=self.quit_dialog, image=self.quit_img)
        btn_quit_dialog.place(x=1400, y=300)

    def add_dialog(self):
        Append_window()

    def search_dialog(self):
        pass

    def edit_dialog(self):
        pass

    def delete_dialog(self):
        pass
    
    def quit_dialog(self):
        answer = mb.askyesno(title='Вийти', message='Дійсно бажаєте вийти?')
        if answer == True:
        	root.destroy()

class Append_window(tkinter.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.font_style = ('helvetica', 16)
        self.init_append_window()
        #self.limiter()


    def init_append_window(self):
        """ВІКНО ВВОДУ ТАЛОНУ"""
        self.title('Додати інформацію')
        w = main_window.winfo_screenwidth() - 620
        h = main_window.winfo_screenheight() - 260
        x = (main_window.winfo_screenwidth() // 2) - (w // 2)
        y = (main_window.winfo_screenheight() // 2 - 30) - (h // 2)
        self.geometry('{}x{}+{}+{}'.format(w, h, x, y))
        self.resizable(False, False)
        """ВИБІР КОМІСІЇ"""
        self.DICT_KOMIS = {0:'--', 1:'Обласна', 2:'Міськміжрайонна', 3:'Калуська', 4:'Коломийська', 5:'Кардіологічна', 6:'Міська', 
                           7:'Травматологічна', 8:'Психіатрична', 9:'Фтизіо-пульмонологічна', 10:'Міжрайонна', 11:'Коломия №2'}
        self.label_komis = ttk.Label(self, text='Вибір комісії:', font = self.font_style).place(x=10, y=30)
        self.komis = ttk.Combobox(self, values=[self.DICT_KOMIS[i] for i in self.DICT_KOMIS], font = self.font_style)
        self.komis.current(0)
        self.komis.place(x=160, y=30)
        """ОГЛЯД"""
        self.label_oglad = ttk.Label(self, text='Огляд', font = self.font_style).place(x=10, y=65)
        self.oglad_var = IntVar()
        self.oglad_var.set(2)
        self.first = ttk.Radiobutton(self, text="Первинний", variable=self.oglad_var, value=1)
        self.first.place(x=160, y=70)
        self.second = ttk.Radiobutton(self, text="Повторний", variable=self.oglad_var, value=2)
        self.second.place(x=310, y=70)
<<<<<<< HEAD
=======
        
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
        """КОД ТАЛОНУ"""
        self.label_kod = ttk.Label(self, text='Код талону', font = self.font_style).place(x=10, y=100)
        self.kod_var = StringVar()
        self.kod = ttk.Entry(self, font = self.font_style, width=7, textvariable = self.kod_var)
        self.kod.place(x=160, y=100)
        self.kod_var.trace("w", lambda *args: self.limiter_symbols(self.kod_var, self.kod, 6))
<<<<<<< HEAD
=======

>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
        """СТАРИЙ КОД ТАЛОНУ"""
        self.label_old_kod = ttk.Label(self, text='Попередній код', font = self.font_style).place(x=10, y=135)
        self.old_kod_var = StringVar()
        self.old_kod = ttk.Entry(self, font = self.font_style, width=7, textvariable = self.old_kod_var)
        self.old_kod.place(x=190, y=135)
        self.old_kod_var.trace("w", lambda *args: self.limiter_symbols(self.old_kod_var, self.old_kod, 6))
<<<<<<< HEAD
=======

>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
        """РІК ПОПЕРЕДНЬОГО ТАЛОНУ"""
        self.label_year_pop_tal = ttk.Label(self, text='Рік', font = self.font_style).place(x=290, y=135)
        self.year_pop_tal_var = StringVar()
        self.year_pop_tal = ttk.Entry(self, font = self.font_style, width=6, textvariable = self.year_pop_tal_var)
        self.year_pop_tal.place(x=335, y=135)
        self.year_pop_tal_var.trace("w", lambda *args: self.limiter_symbols(self.year_pop_tal_var, self.year_pop_tal, 4))
<<<<<<< HEAD
=======
       
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
        """КНОПКА ПОШУК ПОПЕРЕДНЬОГО ТАЛОНУ"""
        self.btn_find_pop_tal = ttk.Button(self, text='ПОШУК ПОПЕРЕДНЬГО')
        self.btn_find_pop_tal.place(x=430, y=137)
        self.btn_find_pop_tal.bind('<Button-1>', lambda event: db.find_old_data([k for k, v in self.DICT_KOMIS.items() if v == self.komis.get()][0],
<<<<<<< HEAD
                                                                                self.kod_var.get(), self.old_kod_var.get(), self.year_pop_tal_var.get(), self.oglad_var,
                                                                                #___ПЕРША_ПОЛОВИНА
=======
                                                                                self.kod_var.get(),
                                                                                self.old_kod_var.get(), self.year_pop_tal_var.get(), self.oglad_var.get(),
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
                                                                                self.sex_var, self.first_name_var, self.name_var, self.surname_var, 
                                                                                self.birth_date_day_var, self.birth_date_mount_var, self.birth_date_year_var, 
                                                                                self.sity, self.raion, self.selo, self.street_var, self.work_var,
                                                                                self.social_var, self.special_var, self.osvita_var, self.diagnoz_1_var,
                                                                                self.uskladn_var, self.zaklad_var,# args[17] - ЦЕ "zaklad_var"
                                                                                #___ДРУГА_ПОЛОВИНА
                                                                                self.data_ogl_day_var, self.data_ogl_mount_var, self.data_ogl_year_var, 
                                                                                self.misce_ogl_var, self.pop_grupa_A_var, self.pop_grupa_B_var, 
                                                                                self.meta_var, self.vstan_grupa_A_var, self.vstan_grupa_B_var, self.strock_var,
                                                                                self.poteri_var, self.prichin_var, self.diagnoz_2_var, self.prac_var,
                                                                                self.expert_var))
        """СТАТЬ"""
        self.sex = ttk.Label(self, text='Стать', font = self.font_style).place(x=10, y=170)
        self.sex_var = IntVar()
        self.sex_var.set(1)
        self.men = ttk.Radiobutton(self, text="чол.", variable=self.sex_var, value=1)
        self.men.place(x=160, y=175)
        self.women = ttk.Radiobutton(self, text="жін.", variable=self.sex_var, value=2)
        self.women.place(x=260, y=175)
        """ПРІЗВИЩЕ"""
        self.label_first_name = ttk.Label(self, text='Прізвище', font = self.font_style).place(x=10, y=205)
        self.first_name_var = StringVar()
        self.first_name = ttk.Entry(self, font = self.font_style, textvariable = self.first_name_var)
        self.first_name.place(x=160, y=205)
        self.first_name_var.trace("w", lambda *args: self.limiter_symbols(self.first_name_var, self.first_name, 20))
        """ІМЯ"""
        self.label_name = ttk.Label(self, text='Ім\'я', font = self.font_style).place(x=10, y=240)
        self.name_var = StringVar()
        self.name = ttk.Entry(self, font = self.font_style, textvariable = self.name_var)
        self.name.place(x=160, y=240)
        self.name_var.trace("w", lambda *args: self.limiter_symbols(self.name_var, self.name, 20))
        """ПО БАТЬКОВІ"""
        self.label_surname = ttk.Label(self, text='Побатькові', font = self.font_style).place(x=10, y=275)
        self.surname_var = StringVar()
        self.surname = ttk.Entry(self, font = self.font_style, textvariable = self.surname_var)
        self.surname.place(x=160, y=275)
        self.surname_var.trace("w", lambda *args: self.limiter_symbols(self.surname_var, self.surname, 20))
        """ДАТА НАРОДЖЕННЯ"""
        self.label_birth_date = ttk.Label(self, text='Дата народження', font = self.font_style).place(x=10, y=310)
        self.birth_date_day_var = StringVar()
        self.birth_date_day = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.birth_date_day_var)
        self.birth_date_day.place(x=200, y=310)
        self.birth_date_day_var.trace("w", lambda *args: self.limiter_symbols(self.birth_date_day_var, self.birth_date_day, 2))
        self.birth_date_mount_var = StringVar()
        self.birth_date_mount = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.birth_date_mount_var)
        self.birth_date_mount.place(x=250, y=310)
        self.birth_date_mount_var.trace("w", lambda *args: self.limiter_symbols(self.birth_date_mount_var, self.birth_date_mount, 2))
        self.birth_date_year_var = StringVar()
        self.birth_date_year = ttk.Entry(self, font = self.font_style, width=5, textvariable = self.birth_date_year_var)
        self.birth_date_year.place(x=300, y=310)
        self.birth_date_year_var.trace("w", lambda *args: self.limiter_symbols(self.birth_date_year_var, self.birth_date_year, 4))
        """МІСТО"""
        self.label_city = ttk.Label(self, text='Місто', font = self.font_style).place(x=10, y=345)
        self.sity = ttk.Combobox(self, values=['--', 'Івано-Франківськ', 'Інше'], font = self.font_style)
        self.sity.current(0)
        self.sity.place(x=160, y=345)
        """РАЙОН"""
        self.label_raion = ttk.Label(self, text='Район', font = self.font_style).place(x=10, y=380)
        self.raion = ttk.Combobox(self, values=['--', 'Богородчанський'], font = self.font_style)
        self.raion.current(0)
        self.raion.place(x=160, y=380)
        """СЕЛО"""
        self.label_selo = ttk.Label(self, text='Село', font = self.font_style).place(x=10, y=415)
        self.selo = ttk.Combobox(self, values=['--', 'Старуня'], font = self.font_style)
        self.selo.current(0)
        self.selo.place(x=160, y=415)
        """ВУЛИЦЯ"""
        self.label_sreet = ttk.Label(self, text='Вулиця', font = self.font_style).place(x=10, y=450)
        self.street_var = StringVar()
        self.street = ttk.Entry(self, font = self.font_style, textvariable = self.street_var)
        self.street.place(x=160, y=450)
        self.street_var.trace("w", lambda *args: self.limiter_symbols(self.street_var, self.street, 20))
        """ПРАЦЮЄ.НЕПРАЦЮЄ"""
        self.label_work = ttk.Label(self, text='Працює/Не працює', font = self.font_style).place(x=10, y=485)
        self.work_var = IntVar()
        self.work_var.set(2)
        self.work = ttk.Radiobutton(self, text="Працює", variable=self.work_var, value=1)
        self.work.place(x=220, y=490)
        self.not_work = ttk.Radiobutton(self, text="Не працює", variable=self.work_var, value=2)
        self.not_work.place(x=320, y=490)
        """СОЦІАЛЬНА КАТЕГОРІЯ"""
        self.label_social = ttk.Label(self, text='Соціальна категорія', font = self.font_style).place(x=10, y=520)
        self.social_var = StringVar()
        self.social = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.social_var)
        self.social.place(x=230, y=520)
        self.social_var.trace("w", lambda *args: self.limiter_symbols(self.social_var, self.social, 2))
        """ПРОФЕСІЯ ЗА ФАХОМ"""
        self.label_special = ttk.Label(self, text='Професія за фахом', font = self.font_style).place(x=10, y=555)
        self.special_var = StringVar()
        self.special = ttk.Entry(self, font = self.font_style, width=11, textvariable = self.special_var)
        self.special.place(x=230, y=555)
        self.special_var.trace("w", lambda *args: self.limiter_symbols(self.special_var, self.special, 10))
        """ОСВІТА"""
        self.label_osvita = ttk.Label(self, text='Освіта', font = self.font_style).place(x=10, y=590)
        self.osvita_var = StringVar()
        self.osvita = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.osvita_var)
        self.osvita.place(x=230, y=590)
        self.osvita_var.trace("w", lambda *args: self.limiter_symbols(self.osvita_var, self.osvita, 1))
        """ДІАГНОЗ"""
        self.label_diagnoz_1 = ttk.Label(self, text='Попередній діагноз', font = self.font_style).place(x=10, y=625)
        self.diagnoz_1_var = StringVar()
        self.diagnoz_1 = ttk.Entry(self, font = self.font_style, width=6, textvariable = self.diagnoz_1_var)
        self.diagnoz_1.place(x=230, y=625)
        self.diagnoz_1_var.trace("w", lambda *args: self.limiter_symbols(self.diagnoz_1_var, self.diagnoz_1, 5))
        """УСКЛАДНЕННЯ"""
        self.label_uskladn =  ttk.Label(self, text='Ускладнення', font = self.font_style).place(x=320, y=625)
        self.uskladn_var = StringVar()
        self.uskladn = ttk.Entry(self, font = self.font_style, width=6, textvariable = self.uskladn_var)
        self.uskladn.place(x=470, y=625)
        self.uskladn_var.trace("w", lambda *args: self.limiter_symbols(self.uskladn_var, self.uskladn, 5))
        """НАПРАВЛЕННЯ"""
        self.label_zaklad =  ttk.Label(self, text='Направлений \n(ким)', font = self.font_style).place(x=10, y=660)
        self.zaklad_var = StringVar()
        self.zaklad = ttk.Entry(self, font = self.font_style, textvariable = self.zaklad_var)
        self.zaklad.place(x=160, y=660)
        self.zaklad_var.trace("w", lambda *args: self.limiter_symbols(self.zaklad_var, self.zaklad, 20))
        """ДАТА ОГЛЯДУ"""
        self.label_data_ogl =  ttk.Label(self, text='Дата огляду', font = self.font_style).place(x=700, y=170)
        self.data_ogl_day_var = StringVar()
        self.data_ogl_day = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.data_ogl_day_var)
        self.data_ogl_day.place(x=840, y=170)
        self.data_ogl_day_var.trace("w", lambda *args: self.limiter_symbols(self.data_ogl_day_var, self.data_ogl_day, 2))
        self.data_ogl_mount_var = StringVar()
        self.data_ogl_mount = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.data_ogl_mount_var)
        self.data_ogl_mount.place(x=890, y=170)
        self.data_ogl_mount_var.trace("w", lambda *args: self.limiter_symbols(self.data_ogl_mount_var, self.data_ogl_mount, 2))
        self.data_ogl_year_var = StringVar()
        self.data_ogl_year = ttk.Entry(self, font = self.font_style, width=5, textvariable = self.data_ogl_year_var)
        self.data_ogl_year.place(x=940, y=170)
        self.data_ogl_year_var.trace("w", lambda *args: self.limiter_symbols(self.data_ogl_year_var, self.data_ogl_year, 4))
        """МІСЦЕ ОГЛЯДУ"""
        self.label_misce_ogl =  ttk.Label(self, text='Місце огляду', font = self.font_style).place(x=700, y=205)
        self.misce_ogl_var = StringVar()
        self.misce_ogl = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.misce_ogl_var)
        self.misce_ogl.place(x=850, y=205)
        self.misce_ogl_var.trace("w", lambda *args: self.limiter_symbols(self.misce_ogl_var, self.misce_ogl, 1))
        """ПОПЕРЕДНЯ ГРУПА a"""
        self.label_pop_grupa_A =  ttk.Label(self, text='Попередня група', font = self.font_style).place(x=700, y=240)
        self.pop_grupa_A_var = StringVar()
        self.pop_grupa_A = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.pop_grupa_A_var)
        self.pop_grupa_A.place(x=900, y=240)
        self.pop_grupa_A_var.trace("w", lambda *args: self.limiter_symbols(self.pop_grupa_A_var, self.pop_grupa_A, 1))
        """ПОПЕРЕДНЯ ГРУПА б"""
        self.pop_grupa_B_var = StringVar()
        self.pop_grupa_B = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.pop_grupa_B_var)
        self.pop_grupa_B.place(x=950, y=240)
        self.pop_grupa_B_var.trace("w", lambda *args: self.limiter_symbols(self.pop_grupa_B_var, self.pop_grupa_B, 1))        
        """МЕТА"""
        self.label_meta =  ttk.Label(self, text='Мета огляду', font = self.font_style).place(x=700, y=275)
        self.meta_var = StringVar()
        self.meta = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.meta_var)
        self.meta.place(x=840, y=275)
        self.meta_var.trace("w", lambda *args: self.limiter_symbols(self.meta_var, self.meta, 2))
        """ВСТАНОВЛ.ГРУПА a"""
        self.label_vstan_grupa_A = ttk.Label(self, text='Встановлена група', font = self.font_style).place(x=700, y=310)
        self.vstan_grupa_A_var = StringVar()
        self.vstan_grupa_A = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.vstan_grupa_A_var)
        self.vstan_grupa_A.place(x=920, y=310)
        self.vstan_grupa_A_var.trace("w", lambda *args: self.limiter_symbols(self.vstan_grupa_A_var, self.vstan_grupa_A, 1))
        """ВСТАНОВЛ.ГРУПА б"""
        self.vstan_grupa_B_var = StringVar()
        self.vstan_grupa_B = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.vstan_grupa_B_var)
        self.vstan_grupa_B.place(x=970, y=310)
        self.vstan_grupa_B_var.trace("w", lambda *args: self.limiter_symbols(self.vstan_grupa_B_var,   self.vstan_grupa_B, 1))
        """СТРОК"""
        self.label_strock =  ttk.Label(self, text='Строк', font = self.font_style).place(x=700, y=345)
        self.strock_var = StringVar()
        self.strock = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.strock_var)
        self.strock.place(x=840, y=345)
        self.strock_var.trace("w", lambda *args: self.limiter_symbols(self.strock_var, self.strock, 1))
        """ВІДСОТКИ ВТРАТИ"""
        self.label_poteri =  ttk.Label(self, text='Відсотки втрати.', font = self.font_style).place(x=700, y=380)
        self.poteri_var = StringVar()
        self.poteri = ttk.Entry(self, font = self.font_style, width=4, textvariable = self.poteri_var)
        self.poteri.place(x=890, y=380)
        self.poteri_var.trace("w", lambda *args: self.limiter_symbols(self.poteri_var, self.poteri, 3))
        """ПРИЧИНА"""
        self.label_prichin =  ttk.Label(self, text='Причина', font = self.font_style).place(x=700, y=415)
        self.prichin_var = StringVar()
        self.prichin = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.prichin_var)
        self.prichin.place(x=840, y=415)
        self.prichin_var.trace("w", lambda *args: self.limiter_symbols(self.prichin_var, self.prichin, 2))
        """ВСТАНОВЛЕНИЙ ДІАГНОЗ"""
        self.label_diagnoz_2 = ttk.Label(self, text='Встановл.діагноз', font = self.font_style).place(x=700, y=450)
        self.diagnoz_2_var = StringVar()
        self.diagnoz_2 = ttk.Entry(self, font = self.font_style, width=6, textvariable = self.diagnoz_2_var)
        self.diagnoz_2.place(x=920, y=450)
        self.diagnoz_2_var.trace("w", lambda *args: self.limiter_symbols(self.diagnoz_2_var, self.diagnoz_2, 5))
        """ПРАЦЕВЛАШТУВАННЯ"""
        self.label_prac = ttk.Label(self, text='Працевлаштування', font = self.font_style).place(x=700, y=485)
        self.prac_var = StringVar()
        self.prac = ttk.Entry(self, font = self.font_style, width=2, textvariable = self.prac_var)
        self.prac.place(x=920, y=485)
        self.prac_var.trace("w", lambda *args: self.limiter_symbols(self.prac_var, self.prac, 1))
        """ПОТРЕБУЄ ЛІКУВАННЯ"""
        self.label_potr_lik = ttk.Label(self, text='Потребує лікування', font = self.font_style).place(x=700, y=520)
        self.potr_lik_1_var = IntVar()
        self.potr_lik_1_var.set(0)
        self.potr_lik_1 = Checkbutton(self, text='1', variable=self.potr_lik_1_var, onvalue=1, offvalue=0)
        self.potr_lik_1.place(x=920, y=520)
        self.potr_lik_2_var = IntVar()
        self.potr_lik_2_var.set(0)
        self.potr_lik_2 = Checkbutton(self, text='2', variable=self.potr_lik_2_var, onvalue=2, offvalue=0)
        self.potr_lik_2.place(x=970, y=520)     
        self.potr_lik_3_var = IntVar()
        self.potr_lik_3_var.set(0)
        self.potr_lik_3 = Checkbutton(self, text='3', variable=self.potr_lik_3_var, onvalue=3, offvalue=0)
        self.potr_lik_3.place(x=1020, y=520)
        self.potr_lik_4_var = IntVar()
        self.potr_lik_4_var.set(0)
        self.potr_lik_4 = Checkbutton(self, text='4', variable=self.potr_lik_4_var, onvalue=4, offvalue=0)
        self.potr_lik_4.place(x=1070, y=520)
        """ЕКСПЕРТНЕ РІШЕННЯ"""
        self.label_expert =  ttk.Label(self, text='Експертне рішення', font = self.font_style).place(x=700, y=555)
        self.expert_var = StringVar()
        self.expert = ttk.Entry(self, font = self.font_style, width=3, textvariable = self.expert_var)
        self.expert.place(x=920, y=555)
        self.expert_var.trace("w", lambda *args: self.limiter_symbols(self.expert_var, self.expert, 2))
        """Рек. з працевлаштування:"""
        self.label_rek_prac =  ttk.Label(self, text='Рек. з працевлаштування:', font = self.font_style).place(x=700, y=585)
        self.rek_prac_1_var = IntVar()
        self.rek_prac_1 = Checkbutton(self, text='1', variable=self.rek_prac_1_var, onvalue=1, offvalue=0)
        self.rek_prac_1.place(x=980, y=585)
        self.rek_prac_2_var = IntVar()
        self.rek_prac_2 = Checkbutton(self, text='2', variable=self.rek_prac_2_var, onvalue=2, offvalue=0) 
        self.rek_prac_2.place(x=1030, y=585)
        self.rek_prac_3_var = IntVar()
        self.rek_prac_3 = Checkbutton(self, text='3', variable=self.rek_prac_3_var, onvalue=3, offvalue=0) 
        self.rek_prac_3.place(x=1080, y=585)
        """Рек. з професійного навчання інвалідів"""
        self.label_prof_nav =  ttk.Label(self, text='Рекомендації з проф. навчання', font = self.font_style).place(x=700, y=620)
        self.prof_nav_1_var = IntVar()
        self.prof_nav_1 = Checkbutton(self, text='4', variable=self.prof_nav_1_var, onvalue=4, offvalue=0)
        self.prof_nav_1.place(x=1040, y=620)
        self.prof_nav_2_var = IntVar()
        self.prof_nav_2 = Checkbutton(self, text='5', variable=self.prof_nav_2_var, onvalue=5, offvalue=0)
        self.prof_nav_2.place(x=1090, y=620)
        self.prof_nav_3_var = IntVar()
        self.prof_nav_3 = Checkbutton(self, text='6', variable=self.prof_nav_3_var, onvalue=6, offvalue=0)
        self.prof_nav_3.place(x=1140, y=620)
        self.prof_nav_4_var = IntVar()
        self.prof_nav_4 = Checkbutton(self, text='7', variable=self.prof_nav_4_var, onvalue=7, offvalue=0)
        self.prof_nav_4.place(x=1190, y=620)
        """Рек. щодо потреби у технічних засобах реабілітації:"""
        self.label_teh_zas =  ttk.Label(self, text='Рек. щодо потреби у тех. зас. реаб.', font = self.font_style).place(x=700, y=655)
        self.teh_zas_1_var = IntVar()
        self.teh_zas_1 = Checkbutton(self, text='8', variable=self.teh_zas_1_var, onvalue=8, offvalue=0)
        self.teh_zas_1.place(x=1080, y=655)
        self.teh_zas_2_var = IntVar()
        self.teh_zas_2 = Checkbutton(self, text='9', variable=self.teh_zas_2_var, onvalue=9, offvalue=0)
        self.teh_zas_2.place(x=1120, y=655)
        self.teh_zas_3_var = IntVar()
        self.teh_zas_3 = Checkbutton(self, text='10', variable=self.teh_zas_3_var, onvalue=10, offvalue=0)
        self.teh_zas_3.place(x=1160, y=655)
        self.teh_zas_4_var = IntVar()
        self.teh_zas_4 = Checkbutton(self, text='11', variable=self.teh_zas_4_var, onvalue=11, offvalue=0)
        self.teh_zas_4.place(x=1210, y=655)
        self.teh_zas_5_var = IntVar()
        self.teh_zas_5 = Checkbutton(self, text='12', variable=self.teh_zas_5_var, onvalue=12, offvalue=0)
        self.teh_zas_5.place(x=1080, y=690)
        self.teh_zas_6_var = IntVar()
        self.teh_zas_6 = Checkbutton(self, text='13', variable=self.teh_zas_6_var, onvalue=13, offvalue=0)
        self.teh_zas_6.place(x=1130, y=690)
        self.teh_zas_7_var = IntVar()
        self.teh_zas_7 = Checkbutton(self, text='14', variable=self.teh_zas_7_var, onvalue=14, offvalue=0)
        self.teh_zas_7.place(x=1180, y=690)
        self.teh_zas_8_var = IntVar()
        self.teh_zas_8 = Checkbutton(self, text='15', variable=self.teh_zas_8_var, onvalue=15, offvalue=0)
        self.teh_zas_8.place(x=1230, y=690)
        """ЕФЕКТИВНІСТЬ ПРОГРАМИ РЕАБІЛІТАЦІЇ"""
        self.label_prog_reab =  ttk.Label(self, text='Ефективність прогр. реаб.:', font = self.font_style).place(x=700, y=725)
        self.prog_reab_var = IntVar()
        self.prog_reab_var.set(2)
        self.prog_reab_1 = ttk.Radiobutton(self, text="повністю", variable=self.prog_reab_var, value=1)
        self.prog_reab_1.place(x=1000, y=730)
        self.prog_reab_2 = ttk.Radiobutton(self, text="часково", variable=self.prog_reab_var, value=2)
        self.prog_reab_2.place(x=1080, y=730)
        self.prog_reab_3 = ttk.Radiobutton(self, text="не виконана", variable=self.prog_reab_var, value=3)
        self.prog_reab_3.place(x=1160, y=730)
        """КНОПКА ДОДАТИ"""
        self.btn_append = ttk.Button(self, text='ДОДАТИ', width=15)
        self.btn_append.place(x=1160, y=765)
        self.btn_append.bind('<Button-1>', lambda event: db.insert_data([k for k, v in self.DICT_KOMIS.items() if v == self.komis.get()][0],
                                                                        [self.kod.get() + ' ' * (8 - len(self.kod.get())) if len(self.kod.get()) <= 8 else self.kod.get()][0], 
                                                                        self.oglad_var.get(), 
                                                                        "-".join((self.data_ogl_day_var.get(), self.data_ogl_mount_var.get(), self.data_ogl_year_var.get())[::-1]),   
                                                                        self.first_name.get(), self.name.get(), self.surname.get(), self.sex_var.get(), "-".join((self.birth_date_day_var.get(), 
                                                                        self.birth_date_mount_var.get(), self.birth_date_year_var.get())[::-1]), self.sity.get(), self.raion.get(), self.selo.get(),
                                                                        self.street.get(), self.work_var.get(), self.social.get(), self.special.get(), self.osvita.get(),
                                                                        self.diagnoz_1.get(), self.zaklad.get(), self.uskladn.get(), self.misce_ogl.get(), self.pop_grupa_A.get(), self.pop_grupa_B.get(),
                                                                        self.meta.get(), self.vstan_grupa_A.get(), self.vstan_grupa_B.get(), self.strock.get(), self.poteri.get(), self.prichin.get(), 
                                                                        '-'.join((str(self.potr_lik_1_var.get()), str(self.potr_lik_2_var.get()), str(self.potr_lik_3_var.get()), str(self.potr_lik_4_var.get()))).strip(), 
                                                                        self.diagnoz_2.get(), self.expert.get(),
                                                                        '-'.join((str(self.rek_prac_1_var.get()), str(self.rek_prac_2_var.get()), str(self.rek_prac_3_var.get()))).strip(),
                                                                        '-'.join((str(self.teh_zas_1_var.get()), str(self.teh_zas_2_var.get()), str(self.teh_zas_3_var.get()), str(self.teh_zas_4_var.get()), 
                                                                                  str(self.teh_zas_5_var.get()), str(self.teh_zas_6_var.get()), str(self.teh_zas_7_var.get()), str(self.teh_zas_8_var.get()))).strip(),
                                                                        self.prog_reab_var.get(), self.prac.get()))
<<<<<<< HEAD
        "КНОПКА РЕДАГУВАТИ"
        self.btn_edit = ttk.Button(self, text='РЕДАГУВАТИ', width=15)
        self.btn_edit.place(x=1000, y=765)
        self.btn_edit.bind('<Button-1>', lambda event: db.edit_data([k for k, v in self.DICT_KOMIS.items() if v == self.komis.get()][0],
                                                                    [self.kod.get() + ' ' * (8 - len(self.kod.get())) if len(self.kod.get()) <= 8 else self.kod.get()][0], 
=======
#___________________________________________________________________
        self.btn_edit = ttk.Button(self, text='РЕДАГУВАТИ', width=15)
        self.btn_edit.place(x=1000, y=765)
        self.btn_edit.bind('<Button-1>', lambda event: db.edit_data(self.kod.get(), 
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
                                                                    self.oglad_var.get(), 
                                                                    "-".join((self.data_ogl_day_var.get(), self.data_ogl_mount_var.get(), self.data_ogl_year_var.get())[::-1]),   
                                                                    self.first_name.get(), 
                                                                    self.name.get(), 
                                                                    self.surname.get(), 
                                                                    self.sex_var.get(), 
                                                                    "-".join((self.birth_date_day_var.get(), self.birth_date_mount_var.get(), self.birth_date_year_var.get())[::-1]), 
                                                                    self.sity.get(), 
                                                                    self.raion.get(), 
                                                                    self.selo.get(),
                                                                    self.street.get(), 
                                                                    self.work_var.get(), 
                                                                    self.social.get(), 
                                                                    self.special.get(), 
                                                                    self.osvita.get(),
                                                                    self.diagnoz_1.get(), 
                                                                    self.zaklad.get(), 
                                                                    self.uskladn.get(), 
                                                                    self.misce_ogl.get(), 
                                                                    self.pop_grupa_A.get(), 
                                                                    self.pop_grupa_B.get(),
                                                                    self.meta.get(), 
                                                                    self.vstan_grupa_A.get(), 
                                                                    self.vstan_grupa_B.get(), 
                                                                    self.strock.get(), 
                                                                    self.poteri.get(),
<<<<<<< HEAD
                                                                    self.prichin.get(),
=======
                                                                    self.prichin.get(), 
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
                                                                    '-'.join((str(self.potr_lik_1_var.get()), str(self.potr_lik_2_var.get()), str(self.potr_lik_3_var.get()), str(self.potr_lik_4_var.get()))).strip(), 
                                                                    self.diagnoz_2.get(), 
                                                                    self.expert.get(),
                                                                    '-'.join((str(self.rek_prac_1_var.get()), str(self.rek_prac_2_var.get()), str(self.rek_prac_3_var.get()))).strip(),
                                                                    '-'.join((str(self.teh_zas_1_var.get()), str(self.teh_zas_2_var.get()), str(self.teh_zas_3_var.get()), str(self.teh_zas_4_var.get()), str(self.teh_zas_5_var.get()), str(self.teh_zas_6_var.get()), str(self.teh_zas_7_var.get()), str(self.teh_zas_8_var.get()))).strip(),
                                                                    self.prog_reab_var.get(), 
                                                                    self.prac.get()))
<<<<<<< HEAD
=======
#___________________________________________________________________
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
        self.grab_set()
        self.focus_set()





    def limiter_symbols(self, Entry_var, Entry, count_sumbol):
        if len(Entry.get()) > 0:
            Entry_var.set(Entry_var.get()[:count_sumbol])
        if len(Entry.get()) == count_sumbol:
            press('tab')


if __name__ == "__main__":
    root = tkinter.Tk()
    main_window = Main_window(root)
    main_window.grid()
    db = backend.DB()
    root.title('Івано-Франківське ОБ МСЕ')
    #root.config(bg="white")
    root.wm_state('zoomed')
    root.mainloop()