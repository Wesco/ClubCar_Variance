__author__ = 'TReische'

import wx


class MainWindow(wx.Frame):
    def __init__(self, title):
        """

        :rtype : None
        """
        wx.Frame.__init__(self, None, -1, title)

        # Controls
        self.dtList1 = None
        self.dtList2 = None
        self.okButton = None

        # Setup
        self.init()

        # Display the window
        self.Show()

    def init(self):
        """

        :rtype : None
        """
        panel1 = wx.Panel(self, -1)
        panel1.SetSizer(wx.BoxSizer(wx.HORIZONTAL))

        #panel2 = wx.Panel(self, -1)
        #panel2.SetSizer(wx.BoxSizer(wx.HORIZONTAL))

        self.dtList1 = wx.ListCtrl(parent=panel1,
                                   id=101,
                                   size=(120, -1),
                                   style=wx.LC_REPORT | wx.LC_SINGLE_SEL)

        self.dtList1.InsertColumn(0, 'Forecast')
        self.dtList1.SetColumnWidth(0, 120)

        panel1.Sizer.Add(item=self.dtList1,
                         proportion=1,
                         flag=wx.ALL | wx.EXPAND,
                         border=0)

        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        hsizer.Add(panel1, 0, wx.EXPAND)
        #hsizer.Add(panel2, 1, wx.EXPAND)

        vsizer = wx.BoxSizer(wx.VERTICAL)
        vsizer.Add(hsizer, 1, wx.EXPAND)

        self.SetAutoLayout(True)
        self.SetSizer(vsizer)
        self.Fit()
        self.SetSize((540, 375))


class FreightApp(wx.App):
    def __init__(self):
        wx.App.__init__(self)
        MainWindow('Variance Report')


if __name__ == '__main__':
    app = FreightApp()
    app.MainLoop()