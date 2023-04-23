import wx
import wx.lib.inspection
import win32com.client as win32


class MainPanel(wx.Panel):     
    def __init__(self, parent):
        super().__init__(parent, size=(800,400))
        
        #main wrapper
        MainSizer = wx.BoxSizer(wx.VERTICAL)

        # first wrapper
        Row_1 = wx.BoxSizer(wx.HORIZONTAL)        
        
        #three buttons added to the wrapper
        my_btn_1 = wx.Button(self, label='', size=(200,20))
        my_btn_2 = wx.Button(self, label='Guide', size=(200,20))
        my_btn_3 = wx.Button(self, label='Settings', size=(200,20))
        Row_1.Add(my_btn_1, 0, wx.ALL | wx.EXPAND, 5)    
        Row_1.Add(my_btn_2, 0, wx.ALL | wx.EXPAND, 5)         
        Row_1.Add(my_btn_3, 0, wx.ALL | wx.EXPAND, 5)        


        #second wrapper
        Row_2 = wx.BoxSizer(wx.HORIZONTAL)        
        
        #three buttons added to the wrapper
        my_btn_1 = wx.Button(self, label='Contact Management', size=(200,40))
        my_btn_1.Bind(wx.EVT_BUTTON, self.Show_CM_Panel)
        
        my_btn_2 = wx.Button(self, label='Create a hello world \n email', size=(200,40))
        my_btn_2.Bind(wx.EVT_BUTTON, self.creat_email)
        
        my_btn_3 = wx.Button(self, label='Campaign Management', size=(200,40))
        Row_2.Add(my_btn_1, 0, wx.ALL | wx.EXPAND, 5)    
        Row_2.Add(my_btn_2, 0, wx.ALL | wx.EXPAND, 5)         
        Row_2.Add(my_btn_3, 0, wx.ALL | wx.EXPAND, 5)        

        MainSizer.Add(Row_1, 0, wx.ALL|wx.CENTER, 5)
        MainSizer.Add(wx.StaticLine(self), 0, wx.ALL|wx.EXPAND, 5)

        MainSizer.Add(Row_2, 0, wx.ALL|wx.CENTER, 5)

        self.SetSizer(MainSizer)
        MainSizer.Fit(self)
        self.Layout()


    def Show_CM_Panel(self, event):        
        self.Parent.ShowCMPanel()



    def creat_email(self, event):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ""
        mail.Subject = "Hello World!!"
        mail.HtmlBody = "Hello World!!"
        mail.Save()
        mail.Close(True)
        
        print("\n\nEmail created and saved as draft in Outlook.\n\n")

        #mail.Display(True)



class ContactManagementPanel(wx.Panel):     
    def __init__(self, parent):
        super().__init__(parent, size=(800,400))
        
        #main wrapper
        MainSizer = wx.BoxSizer(wx.VERTICAL)

        # first wrapper with subtitel
        
        Row_1 = wx.BoxSizer(wx.HORIZONTAL)       
        SubTitle = wx.StaticText(self, label ="Contact Management Panel", style=wx.ALIGN_CENTRE_HORIZONTAL)
        Row_1.Add(SubTitle, 0, wx.ALL | wx.EXPAND, 5)
        MainSizer.Add(Row_1, 0, wx.ALL|wx.CENTER, 5)
        
        # second wrapper with context buttons
        Row_2 = wx.BoxSizer(wx.HORIZONTAL)        
        

        #three buttons added to the wrapper
        my_btn_0 = wx.Button(self, label='Back', size=(200,40))
        my_btn_0.Bind(wx.EVT_BUTTON, self.ShowMainPanel)


        my_btn_1 = wx.Button(self, label='Add new contact', size=(200,40))
        #my_btn_1.Bind(wx.EVT_BUTTON, self.ShowMainPanel)

        my_btn_2 = wx.Button(self, label='Create distribution list', size=(200,40))
        #my_btn_2.Bind(wx.EVT_BUTTON, self.ShowMainPanel)

        my_btn_3 = wx.Button(self, label='Settings', size=(200,40))
        #my_btn_3.Bind(wx.EVT_BUTTON, self.ShowMainPanel)


        Row_2.Add(my_btn_0, 0, wx.ALL | wx.EXPAND, 5)    
        Row_2.Add(my_btn_1, 0, wx.ALL | wx.EXPAND, 5)    
        Row_2.Add(my_btn_2, 0, wx.ALL | wx.EXPAND, 5)         
        Row_2.Add(my_btn_3, 0, wx.ALL | wx.EXPAND, 5)        
        
        MainSizer.Add(Row_2, 0, wx.ALL|wx.EXPAND, 5)

#----------------COLUMNS------------------------------
        Row_3_for_columns = wx.BoxSizer(wx.HORIZONTAL)

        
#----COLUMN_1        
        column_1 = wx.BoxSizer(wx.VERTICAL)

        column_title_txt = wx.StaticText(self, label ="Contacts", style=wx.ALIGN_CENTRE_HORIZONTAL)
        
        column_1.Add(column_title_txt, 0, wx.ALL|wx.EXPAND, 5)
        column_1.Add(wx.StaticLine(self), 0, wx.ALL|wx.EXPAND, 5)

#----END_COLUMN_1        


#----COLUMN_2        
        column_2 = wx.BoxSizer(wx.VERTICAL)

        column_title_txt = wx.StaticText(self, label ="Distribution list", style=wx.ALIGN_CENTRE_HORIZONTAL)
        
        column_2.Add(column_title_txt, 0, wx.ALL|wx.EXPAND, 5)
        column_2.Add(wx.StaticLine(self), 0, wx.ALL|wx.EXPAND, 5)

#----END_COLUMN_2   

        Row_3_for_columns.Add(column_1, 0, wx.ALL|wx.EXPAND, 5)
        Row_3_for_columns.Add(column_2, wx.SizerFlags().Expand().Border(wx.ALL, 5))
        
        MainSizer.Add(Row_3_for_columns, 0, wx.ALL|wx.CENTER, 5)

#----------------END_COLUMNS------------------------------










        self.SetSizer(MainSizer)
        MainSizer.Fit(self)
        self.Layout()
        
    def ShowMainPanel(self, event):
        self.Parent.ShowMainPanel()

class HomeFrame(wx.Frame):  


    def __init__(self):
        #Title bar
        super().__init__(parent=None, title='Marketing Campaign', size=(800,400), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        
        sizer = wx.BoxSizer()
        self.SetSizer(sizer)

        #Create the main panel, and add it to the frame.
        self.MainPanel = MainPanel(self)
        self.Sizer.Add(self.MainPanel,1,wx.EXPAND)

        #Create the other panel, and add it to the frame.
        
        self.CMPanel = ContactManagementPanel(self)
        self.Sizer.Add(self.CMPanel,1,wx.EXPAND)
        self.CMPanel.Hide()
        
        self.Centre()

        #self.ShowMainPanel()

    def ShowMainPanel(parent):
        parent.CMPanel.Hide()
        parent.MainPanel.Show()
        parent.MainPanel.SetSize(parent.Sizer.GetSize())
        parent.MainPanel.Centre()

    def ShowCMPanel(parent):
        parent.MainPanel.Hide()
        parent.CMPanel.Show()
        parent.CMPanel.SetSize(parent.Sizer.GetSize())
        parent.CMPanel.Centre()


if __name__ == '__main__':

    app = wx.App(False)
    frame = HomeFrame()
    frame.Show()
    wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()
