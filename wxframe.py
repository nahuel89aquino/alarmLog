import wx
import os.path
import alarmlog


class MainWindow(wx.Frame):
    def __init__(self, filename='Generador de log.xlsx'):
        super(MainWindow, self).__init__(None, size=(600, 300))
        self.filename = filename
        self.dirname = '.'
        self.CreateInteriorWindowComponents()
        self.CreateExteriorWindowComponents()

    def CreateInteriorWindowComponents(self):
        ''' Create "interior" window components. In this case it is just a
            simple multiline text control. '''
        self.browse = wx.DirPickerCtrl(self, style=wx.DIRP_USE_TEXTCTRL, pos= (280,10), path=self.dirname)
        self.button = wx.Button(self,label= "save", pos=(490,10))
        self.Bind(wx.EVT_BUTTON, self.OnSave, self.button)
        self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY,pos=(0,50), size=(600,300))

    def CreateExteriorWindowComponents(self):
        ''' Create "exterior" window components, such as menu and status
            bar. '''
        self.CreateMenu()
        self.CreateStatusBar()
        self.SetTitle()

    def CreateMenu(self):
        fileMenu = wx.Menu()
        for id, label, helpText, handler in \
            [(wx.ID_ABOUT, '&About', 'Information about this program',
                self.OnAbout),
             (wx.ID_OPEN, '&Open', 'Open a new file', self.OnOpen),
             (None, None, None, None),
             (wx.ID_EXIT, 'E&xit', 'Terminate the program', self.OnExit)]:
            if id == None:
                fileMenu.AppendSeparator()
            else:
                item = fileMenu.Append(id, label, helpText)
                self.Bind(wx.EVT_MENU, handler, item)

        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, '&File') # Add the fileMenu to the MenuBar
        self.SetMenuBar(menuBar)  # Add the menuBar to the Frame

    def SetTitle(self):
        # MainWindow.SetTitle overrides wx.Frame.SetTitle, so we have to
        # call it using super:
        super(MainWindow, self).SetTitle('%s'%self.filename)


    # Helper methods:

    def defaultFileDialogOptions(self):
        ''' Return a dictionary with file dialog options that can be
            used in both the save file dialog as well as in the open
            file dialog. '''
        return dict(message='Choose a file', defaultDir=self.dirname,
                    wildcard='*.*')

    def askUserForFilename(self, **dialogOptions):
        dialog = wx.FileDialog(self, **dialogOptions)
        if dialog.ShowModal() == wx.ID_OK:
            userProvidedFilename = True
            self.filename = dialog.GetFilename()
            self.dirname = dialog.GetDirectory()
            self.SetTitle() # Update the window title with the new filename
        else:
            userProvidedFilename = False
        dialog.Destroy()
        return userProvidedFilename

    # Event handlers:

    def OnAbout(self, event):
        dialog = wx.MessageDialog(self, 'A sample editor\n'
            'in wxPython', 'About Sample Editor', wx.OK)
        dialog.ShowModal()
        dialog.Destroy()

    def OnExit(self, event):
        self.Close()  # Close the main window.

    def OnSave(self, event):
        newdirname = self.browse.GetPath()
        textfile = open(os.path.join(self.dirname, self.filename), 'r')
        alarmlog.open_xlsx(textfile, newdirname)
        self.control.SetValue("\n\nProcesses successfully completed...\n")
        textfile.close()

    def OnOpen(self, event):
        if self.askUserForFilename(style=wx.ID_OPEN,
                                   **self.defaultFileDialogOptions()):
            textfile = open(os.path.join(self.dirname, self.filename), 'r')
            while True:
                line = textfile.readline()
                if line == '':
                    break
                else:
                    self.control.SetValue(textfile.read())
            textfile.close()


app = wx.App()
frame = MainWindow()
frame.Show()
app.MainLoop()
