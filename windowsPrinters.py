import subprocess

class WindowsPrinters:
	def __init__(self,computer='.'):
		"""
		computer can be used to attach to a remote computer if that thrills you
		"""
		self.defaultPostscriptPrinterDriver='HP Color LaserJet 2800 Series PS'
		
	def removePort(self,printerPortName):
		cmd='cscript c:\\Windows\\System32\\Printing_Admin_Scripts\\en-US\\prnport.vbs -d -r "'+printerPortName+'"'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def removePrinter(self,name):
		cmd='rundll32 printui.dll,PrintUIEntry /dl /n "'+name+'"'
		#print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def listPorts(self):
		cmd='cscript c:\\Windows\\System32\\Printing_Admin_Scripts\\en-US\\prnport.vbs -l'
		#print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def makePrinterDefault(self,name):
		cmd='rundll32 printui.dll,PrintUIEntry /y /n "'+name+'"'
		#print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def setPrinterComment(self,name,comment):
		comment=comment.replace('"','\\"').replace('\n','\\n')
		cmd='rundll32 printui.dll,PrintUIEntry /Xs /n "'+name+'" comment "'+comment+'"'
		#print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def addPrinter(self,name,host='127.0.0.1',port=9101,printerPortName=None,makeDefault=False,comment=None):
		port=str(port)
		# -- create the printer port
		if printerPortName==None:
			printerPortName=host+':'+port
		cmd='cscript c:\\Windows\\System32\\Printing_Admin_Scripts\\en-US\\prnport.vbs -md -a -o raw -r "'+printerPortName+'" -h '+host+' -n '+port
		#print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		# -- create the printer
		cmd='rundll32 printui.dll,PrintUIEntry /if /b "'+name+'" /r "'+printerPortName+'" /m "'+self.defaultPostscriptPrinterDriver+'" /Z'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		# -- set the default printer flag
		if makeDefault:
			self.makePrinterDefault(name)
		# -- set the printer comment
		if comment!=None:
			self.setPrinterComment(name,comment)
		
	def printTestPage(self,name):
		cmd='rundll32 printui.dll,PrintUIEntry /k /n "'+name+'"'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def showSettingsDialog(self,name):
		cmd='rundll32 printui.dll,PrintUIEntry /e /n "'+name+'"'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def saveSettings(self,name,filename):
		cmd='rundll32 printui.dll,PrintUIEntry /Ss /n "'+name+'" /a "'+filename+'"'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def loadSettings(self,name,filename):
		cmd='rundll32 printui.dll,PrintUIEntry /Sr /n "'+name+'" /a "'+filename+'"'
		print cmd
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		
	def showPrintUIdllOptions(self):
		cmd='rundll32 PrintUI.dll,PrintUIEntry /?'
		stdout,blah=subprocess.Popen(cmd,stdin=None,stderr=subprocess.STDOUT,stdout=subprocess.PIPE,shell=True).communicate()
		print stdout
		#ctypes.windll.PrintUI.PrintUIEntry('/?') # it might be cool to do this someday


if __name__=='__main__':			
	p=WindowsPrinters()
	p.addPrinter("my new printer")
	p.removePrinter("my new printer")
	#p.removePort("127.0.0.1:9101")


