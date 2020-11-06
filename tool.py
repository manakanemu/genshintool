import os
import ctypes
import psutil
import winreg
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import tkinter as tk
import tkinter.ttk as ttk


class Genshin(ttk.Frame):
    class Audio:
        def __init__(self):
            self.__genshin_autio_session = None
            self.is_prepare = False
            self.__history_volume = 1.0
            self.init()

        def init(self):
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                try:
                    if session.Process.name() == 'YuanShen.exe':
                        self.__genshin_autio_session = session
                        self.__genshin_autio_interface = session._ctl.QueryInterface(ISimpleAudioVolume)
                        self.__genshin_volume = self.__genshin_autio_interface.GetMasterVolume()
                        self.__history_volume = self.__genshin_volume
                        self.is_prepare = True
                        break
                except:
                    pass


        def get_volume(self):
            if not self.__genshin_autio_session:
                self.init()
            try:
                return self.__genshin_volume
            except:
                print('未找到原神音频线程')
                return None

        def set_volume(self, volume):
            if not self.__genshin_autio_session:
                self.init()
            volume = min(1.0, max(0.0, volume))
            try:
                self.__history_volume = self.__genshin_volume
                self.__genshin_volume = volume
                self.__genshin_autio_interface.SetMasterVolume(volume, None)
            except:
                print('未找到原神音频线程')

        def peace(self):
            self.set_volume(0.0)

        def impeace(self):
            self.set_volume(self.__history_volume)

    class Environment:
        def __init__(self):
            work_path = None

            try:
                reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\原神')
                value,type = winreg.QueryValueEx(reg_key,'UninstallString')
                work_path = os.path.split(value)[0]
            except:
                work_path = None

            if not work_path:
                work_path = os.path.abspath('..')

            file_list = os.listdir(work_path)
            if 'launcher.exe' not in file_list:
                print('\033[31m未在注册表搜索到原神安装目录，请把启动器文件夹放到原神根目录')
                input('按回车键退出')
                exit(0)

            self.work_path = work_path
            self.gensin_launcher_path = os.path.join(work_path, 'launcher.exe')


    def __init__(self,app=None):
        super().__init__(app)
        self.pack()
        self.creat_widget()

        self.audio = self.Audio()
        self.path = self.Environment()
        self.peace = False
        self.is_prepare = self.search_process() or self.audio.is_prepare

        if not self.is_prepare:
            self.launcher_game()

    def search_process(self):
        pids = psutil.pids()
        for pid in pids:
            process = psutil.Process(pid)
            if process.name() == 'launcher.exe':
                return True
        return False


    def creat_widget(self):
        self.message_text = tk.StringVar()
        self.op1_text = tk.StringVar()
        self.op2_text = tk.StringVar()
        self.op3_text = tk.StringVar()
        self.op4_text = tk.StringVar()

        self.message_text.set('')
        self.op1_text.set('1.禁用爬墙透明')
        self.op2_text.set('2.游戏静音')
        self.op3_text.set('3.半音量')
        self.op4_text.set('4.关闭游戏')

        self.message_lable = ttk.Label(self,textvariable=self.message_text,width=20)
        self.op1_button = ttk.Button(self, textvariable=self.op1_text,width=20,command=self.clean_model)
        self.op2_button = ttk.Button(self, textvariable=self.op2_text,width=20,command=self.swich_peace)
        self.op3_button = ttk.Button(self, textvariable=self.op3_text,width=20,command=self.swich_half_volume)
        self.op4_button = ttk.Button(self, textvariable=self.op4_text,width=20,command=self.close_game)

        self.message_lable.pack()
        self.op1_button.pack()
        self.op2_button.pack()
        self.op3_button.pack()
        self.op4_button.pack()

        self.bind_all('<KeyPress>',self.keypress)

    def keypress(self,key):
        if key.keycode == 49:
            self.clean_model()
        if key.keycode == 50:
            self.swich_peace()
        if key.keycode == 51:
            self.swich_half_volume()
        if key.keycode == 52:
            self.close_game()


    def swich_peace(self):
        if self.peace:
            self.audio.impeace()
            self.peace = False
            self.message_text.set('取消静音')
            self.op2_text.set('2：游戏静音')
        else:
            self.audio.peace()
            self.peace = True
            self.message_text.set('已静音')
            self.op2_text.set('2：取消游戏静音')

    def swich_half_volume(self):
        volume = self.audio.get_volume()
        if volume <=0.5:
            self.audio.set_volume(1.0)
            self.message_text.set('游戏全音量')
            self.op3_text.set('3：半音量')

        else:
            self.audio.set_volume(0.5)
            self.message_text.set('游戏半音量')
            self.op3_text.set('3：全音量')

        self.peace = False
        self.info_op2 = '2：游戏静音\n'


    def launcher_game(self):
        os.startfile('"{}"'.format(self.path.gensin_launcher_path))

    def close_game(self):
        os.system('{}{}'.format("taskkill /F /IM ", 'YuanShen.exe'))
        os.system('{}{}'.format("taskkill /F /IM ", 'launcher.exe'))
        self.quit()

    def clean_model(self):
        try:
            model_path = os.path.join(self.path.work_path,
                                      'Genshin Impact Game\\YuanShen_Data\\Persistent\\AssetBundles\\blocks\\00\\29342328.blk')
            open(model_path, 'w').close()
            self.message_text.set('已禁用爬墙透明模型')
        except:
            self.message_text.set('禁用失败')




ctypes.windll.shcore.SetProcessDpiAwareness(1)
ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)

window = tk.Tk()
window.title('Genshin Tool')
window.maxsize(400,300)
window.minsize(400,300)
window.tk.call('tk', 'scaling', ScaleFactor/75)
genshin = Genshin(window)


window.mainloop()

