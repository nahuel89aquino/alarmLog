import wx
import os.path
import codecs
from moduls import alarmlog


class MainWindow(wx.Frame):
    def __init__(self, filename='Generador de log.xlsx'):
        super(MainWindow, self).__init__(None, size=(600, 300))
        self.filename = filename
        self.dirname = '.'
        self.SetBackgroundColour(wx.Colour(240,240,240))
        self.SetIcon(wx.Icon('img\icon.ico'))
        self.CreateInteriorWindowComponents()
        self.CreateExteriorWindowComponents()

    def CreateInteriorWindowComponents(self):
        ''' Create "interior" window components. In this case it is just a
            simple multiline text control. '''
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        vSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.browse = wx.DirPickerCtrl(self, style=wx.DIRP_USE_TEXTCTRL,size=(250, 25), path=self.dirname)
        self.button = wx.Button(self, label="Convertir", pos=(490,10), size=(80,25))
        self.Bind(wx.EVT_BUTTON, self.OnSave, self.button)
        self.texto = wx.StaticText(self, label="Numero de serie")
        self.editext = wx.TextCtrl(self,value='*', style= wx.TE_MULTILINE | wx.TE_NO_VSCROLL, size=(150, 25))
        vSizer.Add(self.texto, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL,2)
        vSizer.Add(self.editext,0,wx.ALIGN_LEFT | wx.ALL,2)
        vSizer.Add(self.browse, 0, wx.ALIGN_RIGHT | wx.ALL,2)
        hSizer.Add(vSizer, 0, wx.EXPAND)
        hSizer.Add(self.button, 0, wx.ALIGN_RIGHT | wx.ALL,2)
        self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY, size=(600,300))
        mainSizer.Add(hSizer, 0,wx.ALIGN_RIGHT)
        mainSizer.Add(self.control, 1, wx.EXPAND)
        self.SetSizerAndFit(mainSizer)

    def CreateExteriorWindowComponents(self):
        ''' Create "exterior" window components, such as menu and status
            bar. '''
        self.CreateMenu()
        self.CreateStatusBar()
        self.SetTitle()

    def CreateMenu(self):
        fileMenu = wx.Menu()
        for id, label, helpText, handler in \
            [(wx.ID_ABOUT, '&Acerca de', 'Información sobre este programa',
                self.OnAbout),
             (wx.ID_OPEN, '&Abrir', 'Abrir un nuevo archivo de log de eventos', self.OnOpen),
             (None, None, None, None),
             (wx.ID_EXIT, '&Salir', 'Salir del programa', self.OnExit)]:
            if id == None:
                fileMenu.AppendSeparator()
            else:
                item = fileMenu.Append(id, label, helpText)
                self.Bind(wx.EVT_MENU, handler, item)

        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, '&Archivo') # Add the fileMenu to the MenuBar
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
        return dict(message='Seleccionar archivo', defaultDir=self.dirname,
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
        dialog = wx.MessageDialog(self, 'Este es un conversor de log.txt a archivos exel', 'Conversor', wx.OK)
        dialog.ShowModal()
        dialog.Destroy()

    def OnExit(self, event):
        self.Close()  # Close the main window.

    def OnSave(self, event):

        if self.control.GetValue() is '':
            dialog = wx.MessageDialog(self, 'No hay ningun documento abierto','Error', wx.OK)
            dialog.ShowModal()
            dialog.Destroy()
        elif self.editext.GetValue() is '*':
            dialog = wx.MessageDialog(self, 'Debe ingresar el numero de serie del equipo al que pertenece'
                                            ' el log de alarmas','Error', wx.OK)
            dialog.ShowModal()
            dialog.Destroy()
        else:
            try:
                newdirname = self.browse.GetPath()
                name = self.editext.GetValue()
                textfile = codecs.open(os.path.join(self.dirname, self.filename), 'r','utf-8')
                alarmlog.open_xlsx(textfile, newdirname, name)
                self.control.SetValue("\nProcesses successfully completed...\n")
                textfile.close()
            except ValueError:
                dialog = wx.MessageDialog(self, 'El campo no puede estar vacio'
                                                '\n Ingrese el numero de serie','Error', wx.OK)
                dialog.ShowModal()
                dialog.Destroy()


    def OnOpen(self, event):
        if self.askUserForFilename(style=wx.ID_OPEN,
                                   **self.defaultFileDialogOptions()):
            textfile = codecs.open(os.path.join(self.dirname, self.filename), 'r', 'utf-8')
            while True:
                line = textfile.readline()
                if line == '':
                    break
                else:
                    self.control.SetValue(textfile.read())
            textfile.close()

def main():
    app = wx.App()
    frame = MainWindow()
    frame.Show()
    app.MainLoop()

if __name__ == '__main__':
    main()
