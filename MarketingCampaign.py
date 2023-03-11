import wx
#import wx.lib.inspection

import win32com.client as win32



class MainPanel(wx.Panel):     
    def __init__(self, parent):
        super().__init__(parent)
        
        #main wrapper
        MainSizer = wx.BoxSizer(wx.VERTICAL)

        # first wrapper
        Row_1 = wx.BoxSizer(wx.HORIZONTAL)        
        
        #three buttons added to the wrapper
        my_btn_1 = wx.Button(self, label='', size=(200,20))
        my_btn_2 = wx.Button(self, label='', size=(200,20))
        my_btn_3 = wx.Button(self, label='Settings', size=(200,20))
        Row_1.Add(my_btn_1, 0, wx.ALL | wx.EXPAND, 5)    
        Row_1.Add(my_btn_2, 0, wx.ALL | wx.EXPAND, 5)         
        Row_1.Add(my_btn_3, 0, wx.ALL | wx.EXPAND, 5)        


        #second wrapper
        Row_2 = wx.BoxSizer(wx.HORIZONTAL)        
        
        #three buttons added to the wrapper
        my_btn_1 = wx.Button(self, label='Contact Management', size=(200,40))
        my_btn_1.Bind(wx.EVT_BUTTON, self.on_press_1)
        
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


    def on_press_1(self, event):
        print("Inside MainPanel(). This is the parent: ", self.Parent)
        
        self.Parent.ShowCMPanel()

        #self.Parent.panel =  CMPanel(self.Parent)
        #self.Parent.Show()



    def creat_email(self, event):
        print("Hello")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ""
        mail.Subject = "Hello World!"
        mail.HtmlBody = "Hello World!"
        mail.Save()
        mail.Close(True)
        
        #mail.Display(True)




class CMPanel(wx.Panel):     
    def __init__(self, parent):
        super().__init__(parent)
        
        #main wrapper
        MainSizer = wx.BoxSizer(wx.VERTICAL)

        # first wrapper
        Row_1 = wx.BoxSizer(wx.HORIZONTAL)        
        
        #three buttons added to the wrapper
        my_btn_1 = wx.Button(self, label='', size=(200,20))
        my_btn_2 = wx.Button(self, label='', size=(200,20))
        my_btn_3 = wx.Button(self, label='Settings', size=(200,20))
        Row_1.Add(my_btn_1, 0, wx.ALL | wx.EXPAND, 5)    
        Row_1.Add(my_btn_2, 0, wx.ALL | wx.EXPAND, 5)         
        Row_1.Add(my_btn_3, 0, wx.ALL | wx.EXPAND, 5)        

        MainSizer.Add(Row_1, 0, wx.ALL|wx.CENTER, 5)
        MainSizer.Add(wx.StaticLine(self), 0, wx.ALL|wx.EXPAND, 5)

        self.SetSizer(MainSizer)
        MainSizer.Fit(self)
        self.Layout()

    '''
    def on_press(self, event):
        print("Hide?")
        self.Hide()
    ''' 

class MyFrame(wx.Frame):  


    def __init__(self):
        #Title bar
        super().__init__(parent=None, title='Marketing Campaign', size=(800,400))
        
        sizer = wx.BoxSizer()
        self.SetSizer(sizer)

        #Create the main panel, and add it to the frame.
        self.MainPanel = MainPanel(self)
        self.Sizer.Add(self.MainPanel,1,wx.EXPAND)

        #Create the other panel, and add it to the frame.
        self.CMPanel = CMPanel(self)
        self.Sizer.Add(self.CMPanel,1,wx.EXPAND)
        self.CMPanel.Hide()

        self.Centre()

        self.ShowMainPanel()

    def ShowMainPanel(parent):
        '''
        #hide the previous panel, in case there was any
        if(parent.panel):
            parent.panel.Hide()
        '''
        #parent.Hide()

        print("Inside MyFrame.ShowMainPanel(). This is the parent: ", parent)

        #Show the panel in the frame       
        parent.MainPanel.Show()
        parent.MainPanel.Layout()

    def ShowCMPanel(parent):
        
        parent.MainPanel.Hide()
        parent.Sizer.Clear()
        parent.Sizer.Add(parent.CMPanel)
        parent.CMPanel.Show()
        parent.CMPanel.Layout()

        print("Inside MyFrame.ShowCMPanel(). This is the parent: ",parent)
    


if __name__ == '__main__':

    app = wx.App(False)
    frame = MyFrame()
    frame.Show()
    #wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()
