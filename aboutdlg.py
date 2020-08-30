import wx

text = """Описание программы:
Программа ActOf предназначена для облегчения процесса
создания актов.
Автор: 
Поварисов Максим Александрович
Версия:
0.9"""

class AboutDlg(wx.Dialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)
        self.text = wx.StaticText(self, wx.ID_ANY, label=text, style=wx.ALIGN_CENTER)
        # self.text.SetLabelMarkup()
        self.mainsizer.Add(self.text, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)

        self.SetSizer(self.mainsizer)