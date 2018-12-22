#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
Handy convenience class for managing (installing, removing, etc)
windows printers
"""
import subprocess


class WindowsPrinters(object):
	"""
	Handy convenience class for managing (installing, removing, etc)
	windows printers
	"""

	def __init__(self,computer='.'):
		"""
		computer can be used to attach to a remote computer if that thrills you

		NOTE: there may be a better default driver to use
		"""
		self.defaultPostscriptPrinterDriver='HP Color LaserJet 2800 Series PS'

	def removePort(self,printerPortName):
		"""
		Remove a printer port
		"""
		cmd=r'cscript c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -d -r "'+printerPortName+'"'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def removePrinter(self,name):
		"""
		Remove an installed printer
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /dl /n "'+name+'"'
		#print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def listPorts(self):
		"""
		List all installed printer ports
		"""
		cmd=r'cscript c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -l'
		#print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def makePrinterDefault(self,name):
		"""
		Assign a printer to be the system default
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /y /n "'+name+'"'
		#print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def setPrinterComment(self,name,comment):
		"""
		Add a comment to the given printer device
		"""
		comment=comment.replace('"','\\"').replace('\n','\\n')
		cmd=r'rundll32 printui.dll,PrintUIEntry /Xs /n "'+name+'" comment "'+comment+'"'
		#print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def addPrinter(self,name,host='127.0.0.1',port=9101,printerPortName=None,makeDefault=False,comment=None):
		"""
		Add a new printer to the system
		"""
		port=str(port)
		# -- create the printer port
		if printerPortName is None:
			printerPortName=host+':'+port
		cmd=r'cscript c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -md -a -o raw -r "'+printerPortName+'" -h '+host+' -n '+port
		#print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		# -- create the printer
		cmd=r'rundll32 printui.dll,PrintUIEntry /if /b "'+name+'" /r "'+printerPortName+'" /m "'+self.defaultPostscriptPrinterDriver+'" /Z'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		# -- set the default printer flag
		if makeDefault:
			self.makePrinterDefault(name)
		# -- set the printer comment
		if comment!=None:
			self.setPrinterComment(name,comment)

	def printTestPage(self,name):
		"""
		Send a test page to the given printer
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /k /n "'+name+'"'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def showSettingsDialog(self,name):
		"""
		Pop up the settings dialog for a given printer
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /e /n "'+name+'"'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def saveSettings(self,name,filename):
		"""
		Save printer settings out to a binary file

		(I have no idea what the format is, so it cannot yet be edited externally.)
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /Ss /n "'+name+'" /a "'+filename+'"'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def loadSettings(self,name,filename):
		"""
		Load printer settings in from a file

		(I have no idea what the format is, so it cannot yet be edited externally.)
		"""
		cmd=r'rundll32 printui.dll,PrintUIEntry /Sr /n "'+name+'" /a "'+filename+'"'
		print cmd
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout

	def showPrintUIdllOptions(self):
		"""
		Show the ui options
		"""
		cmd=r'rundll32 PrintUI.dll,PrintUIEntry /?'
		stdout,_=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		#ctypes.windll.PrintUI.PrintUIEntry('/?') # it might be cool to do this someday


if __name__=='__main__':
	print "Presently the command line is diagnostic only, registering, and then unregistering a printer"
	p=WindowsPrinters()
	p.addPrinter("my new printer")
	p.removePrinter("my new printer")
	#p.removePort("127.0.0.1:9101")
