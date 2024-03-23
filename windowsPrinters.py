#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
Handy convenience class for managing (installing, removing, etc)
windows printers
"""
import typing
import subprocess


class WindowsPrinters(object):
    """
    Handy convenience class for managing (installing, removing, etc)
    windows printers
    """

    def __init__(self):
        """
        NOTE: there may be a better default driver to use
        """
        self.defaultPostscriptPrinterDriver='HP Color LaserJet 2800 Series PS'

    def removePort(self,printerPortName):
        """
        Remove a printer port
        """
        cmd=['cscript',
             r'c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs',
             '-d','-r',printerPortName]
        print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def removePrinter(self,name):
        """
        Remove an installed printer
        """
        cmd=['rundll32','printui.dll,PrintUIEntry','/dl','/n',name]
        #print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def listPorts(self):
        """
        List all installed printer ports
        """
        cmd=['cscript',
             r'c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs',
             '-l']
        #print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def makePrinterDefault(self,name):
        """
        Assign a printer to be the system default
        """
        cmd=['rundll32','printui.dll,PrintUIEntry','/y','/n',name]
        #print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def setPrinterComment(self,name:str,comment:str)->None:
        """
        Add a comment to the given printer device
        """
        comment=comment.replace('"','\\"').replace('\n','\\n')
        cmd=['rundll32','printui.dll,PrintUIEntry','/Xs',
            '/n',name,
            'comment',comment]
        #print(cmd)
        po=subprocess.Popen(
            cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def addPrinter(self,
        name:str,
        host:str='127.0.0.1',port:int=9101,
        printerPortName:typing.Optional[str]=None,
        makeDefault:bool=False,
        comment:typing.Optional[str]=None
        )->None:
        """
        Add a new printer to the system
        """
        port=str(port)
        # -- create the printer port
        if printerPortName is None:
            printerPortName=host+':'+port
        cmd=['cscript',
            r'c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs',
            '-md','-a','-o','raw',
            '-r',printerPortName,
            '-h',host,'-n',port]
        #print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)
        # -- create the printer
        cmd=['rundll32','printui.dll,PrintUIEntry','/if',
            '/b',name,
            '/r',printerPortName,
            '/m',self.defaultPostscriptPrinterDriver,
            '/Z']
        print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)
        # -- set the default printer flag
        if makeDefault:
            self.makePrinterDefault(name)
        # -- set the printer comment
        if comment is not None:
            self.setPrinterComment(name,comment)

    def printTestPage(self,name:str)->None:
        """
        Send a test page to the given printer
        """
        cmd=['rundll32',
            'printui.dll,PrintUIEntry',
            '/k',
            '/n',name]
        print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def showSettingsDialog(self,name:str)->None:
        """
        Pop up the settings dialog for a given printer
        """
        cmd=['rundll32',
            'printui.dll,PrintUIEntry',
            '/e',
            '/n',name]
        print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def saveSettings(self,name:str,filename:str)->None:
        """
        Save printer settings out to a binary file

        (I have no idea what the format is,
        so it cannot yet be edited externally.)
        """
        cmd=['rundll32',
            'printui.dll,PrintUIEntry',
            '/Ss',
            '/n',name,
            '/a',filename]
        print(cmd)
        po=subprocess.Popen(cmd,
            stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def loadSettings(self,name,filename):
        """
        Load printer settings in from a file

        (I have no idea what the format is,
        so it cannot yet be edited externally.)
        """
        cmd=['rundll32',
            'printui.dll,PrintUIEntry',
            '/Sr',
            '/n',name,
            '/a',filename]
        print(cmd)
        po=subprocess.Popen(
            cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)

    def showPrintUIdllOptions(self):
        """
        Show the ui options
        """
        cmd=r'rundll32 PrintUI.dll,PrintUIEntry /?'
        po=subprocess.Popen(
            cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,
            shell=True)
        stdout,_=po.communicate()
        print(stdout)
        # TODO: it might be cool to do this someday
        #ctypes.windll.PrintUI.PrintUIEntry('/?')


if __name__=='__main__':
    print('Presently the command line is diagnostic only,')
    print('registering, and then unregistering a printer')
    p=WindowsPrinters()
    p.addPrinter("my new printer")
    p.removePrinter("my new printer")
    #p.removePort("127.0.0.1:9101")
