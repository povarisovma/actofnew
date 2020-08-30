import wx
import settings


ID_MD_CHOSDIRLOC = 105
ID_MD_CHOSDIR = 106
ID_MD_CHOSTMPL = 107
ID_MD_CHOSFILETMPLDOCX = 108
ID_MD_PATHDIRACTLOC = 111
ID_MD_PATHDIRACT = 112
ID_MD_PATHFILETEMPLDOCX = 113



class SettingsDlg(wx.Dialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        #Обьявление главного сайзера:
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)

        #Блок 1, виджеты для указания пути к папке локальных актов--------------------------------------------------
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с локальными актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.folderactssizer = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tc_actloc_path = wx.StaticText(self, id=ID_MD_PATHDIRACTLOC, style=wx.ST_ELLIPSIZE_START,
                                            label=settings.get_local_acts_path_folder())
        self.folderactssizer.Add(self.tc_actloc_path, proportion=1)
        self.folderactssizer.Add(wx.Button(self, id=ID_MD_CHOSDIRLOC, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)

        # Блок 2, виджеты для указания пути к папке общих актов------------------------------------------------------
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с общими актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.folderactssizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer2, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tc_act_path = wx.StaticText(self, id=ID_MD_PATHDIRACT, style=wx.ST_ELLIPSIZE_START,
                                         label=settings.get_general_acts_path_folder())
        self.folderactssizer2.Add(self.tc_act_path, proportion=1)
        self.folderactssizer2.Add(wx.Button(self, id=ID_MD_CHOSDIR, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)

        #Блок 3, виджеты дляуказания пути к шаблону docx
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к файлу шаблона docx:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tmpldocx_path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.tmpldocx_path_sizer, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tc_tmpl_docxfile_path = wx.StaticText(self, id=ID_MD_PATHFILETEMPLDOCX, style=wx.ST_ELLIPSIZE_START,
                                         label=settings.get_docx_templ_path())
        self.tmpldocx_path_sizer.Add(self.tc_tmpl_docxfile_path, proportion=1)
        self.tmpldocx_path_sizer.Add(wx.Button(self, id=ID_MD_CHOSFILETMPLDOCX, label='...'),
                                     flag=wx.EXPAND | wx.LEFT, border=10)

        #Добавление главного сайзера с виджетами в окно
        self.SetSizer(self.mainsizer)

        #Назначение кнопок и функций
        self.Bind(wx.EVT_BUTTON, self.choosediractsloc, id=ID_MD_CHOSDIRLOC)
        self.Bind(wx.EVT_BUTTON, self.choosediracts, id=ID_MD_CHOSDIR)
        self.Bind(wx.EVT_BUTTON, self.choosefiletmpldocx, id=ID_MD_CHOSFILETMPLDOCX)

    def choosediractsloc(self, event):
        dlg = wx.DirDialog(self, message="Выберите папку расположения локальных актов", style=wx.RESIZE_BORDER,
                           defaultPath=self.tc_actloc_path.Label)
        res = dlg.ShowModal()
        if res == wx.ID_OK:
            self.tc_actloc_path.SetLabel(path_to_string(dlg.GetPath()))
            settings.set_local_acts_path_folder_in_settings(path_to_string(dlg.GetPath()))

    def choosediracts(self, event):
        dlg = wx.DirDialog(self, message="Выберите папку расположения общих актов", defaultPath=self.tc_act_path.Label)
        res = dlg.ShowModal()
        if res == wx.ID_OK:
            self.tc_act_path.SetLabel(path_to_string(dlg.GetPath()))
            settings.set_general_acts_path_folder_in_settings(path_to_string(dlg.GetPath()))

    def choosefiletmpldocx(self, event):
        with wx.FileDialog(self, "Выбрать файл шаблона docx", wildcard="Файл MS Word docx(*.docx)|*.docx", style=wx.FD_OPEN) as file:
            if file.ShowModal() == wx.ID_CANCEL:
                return
            pathname = file.GetPath()
            self.tc_tmpl_docxfile_path.SetLabel(pathname)
            settings.set_path_to_docx_templ_in_settings(pathname)


def path_to_string(path):
    if path[-1:-2:-1] == "\\":
        return path
    else:
        return path + "\\"
