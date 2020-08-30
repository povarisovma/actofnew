import wx
import templatesdb as db

ID_BTN_OK = 1
ID_BTN_CANCEL = 2


class ChangeTemplDlg(wx.Dialog):
    numtempl = 0
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #Получение выбраного номера шаблона:
        self.numtempl = self.Parent.OLVtempl_acts.GetSelectedObjects()[0]['templnum']

        #Объявление главного сайзера и создание панели:
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)

        #Задание шрифта форм
        self.font = wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Montserrat')

        #Создание строки изменения описания шаблона:
        self.desc_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.ST_desc = wx.StaticText(self, wx.ID_ANY, label="Описание шаблона:")
        self.ST_desc.SetFont(self.font)

        self.TC_desc = wx.TextCtrl(self, value='', size=(300, -1))
        self.TC_desc.SetValue(db.gettemplatedescfromdb(self.numtempl))
        
        #Создание поля для текста акта
        self.ST_body = wx.StaticText(self, wx.ID_ANY, label="Текст шаблона:")
        self.ST_body.SetFont(self.font)

        self.TC_body = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE)
        self.TC_body.SetValue(db.gettemplatetextfromdb(self.numtempl))

        #Создание кнопок ок и отмены
        self.BTN_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.BTN_ok = wx.Button(self, id=ID_BTN_OK, size=(100, -1), label='Применить')
        self.BTN_cancel = wx.Button(self, id=ID_BTN_CANCEL, size=(100, -1), label='Отмена')
        self.BTN_ok.SetFont(self.font)
        self.BTN_cancel.SetFont(self.font)

        #Добавление в сайзеры всех
        self.desc_sizer.Add(self.ST_desc, flag= wx.LEFT | wx.TOP, border=5)
        self.desc_sizer.Add(self.TC_desc, flag= wx.LEFT | wx.TOP | wx.RIGHT, border=5)
        self.mainsizer.Add(self.desc_sizer)
        self.mainsizer.Add(self.ST_body, flag= wx.LEFT | wx.TOP, border=5)
        self.mainsizer.Add(self.TC_body, proportion=1, flag= wx.LEFT | wx.TOP | wx.RIGHT | wx.EXPAND, border=5)
        self.BTN_sizer.Add(self.BTN_ok, flag=wx.LEFT | wx.TOP | wx.BOTTOM, border=5)
        self.BTN_sizer.Add(self.BTN_cancel, flag=wx.LEFT | wx.TOP | wx.RIGHT | wx.BOTTOM, border=5)
        self.mainsizer.Add(self.BTN_sizer, flag=wx.ALIGN_RIGHT)

        #Подключение главного сайзера:
        self.SetSizer(self.mainsizer)
        
        #Назначение клавиш
        self.Bind(wx.EVT_BUTTON, self.btn_on_ok, id=ID_BTN_OK)
        self.Bind(wx.EVT_BUTTON, self.btn_on_cancel, id=ID_BTN_CANCEL)
        

    def btn_on_ok(self, event):
        db.setupdatetemplateindb(self.numtempl, self.TC_desc.GetValue(), self.TC_body.GetValue())
        self.EndModal(wx.ID_CANCEL)

        

    def btn_on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)