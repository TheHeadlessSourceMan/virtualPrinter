#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
We could use RedMon to redirect a port to a program, but the idea of this file is to bypass that.

Instead, we simply set up a loopback ip and act like a network printer.
"""
import os
import time
import socket
import atexit
import select


class PrintServer(object):
	"""
	We could use RedMon to redirect a port to a program, but the idea of this file is to bypass that.

	Instead, we simply set up a loopback ip and act like a network printer.
	"""

	def __init__(self,printerName='My Virtual Printer',ip='127.0.0.1',port=9001,
		autoInstallPrinter=True,printCallbackFn=None):
		"""
		You can do an ip other than 127.0.0.1 (localhost), but really
		a better way is to install the printer and use windows sharing.

		If you choose another port, you need to right click on your printer
		and go into properties->Ports->Configure Port
		and then change the port number.

		autoInstallPrinter is used to install the printer in the operating system
		(currently only supports Windows)

		printCallbackFn is a function to be called with received print data
			if it is None, then will save it out to a file.
		"""
		self.ip=ip
		if port is None:
			port=0 # meaning, "any unused port"
		self.port=port
		self.buffersize=20  # Normally 1024, but we want fast response
		self.autoInstallPrinter=autoInstallPrinter
		self.printerName=printerName
		self.running=False
		self.keepGoing=False
		self.osPrinterManager=None
		self.printerPortName=None
		self.printCallbackFn=printCallbackFn

	def __del__(self):
		"""
		Do some clean up when object is deleted
		"""
		if self: # this will always be called on program exit,
					#so may come in again if the object is already deleted
			if self.autoInstallPrinter:
				self._uninstallPrinter()

	def _installPrinter(self,ip,port):
		"""
		Install the printer to the ip address
		"""
		atexit.register(self.__del__) # ensure that __del__ always gets called when the program exits
		if os.name=='nt':
			import windowsPrinters
			self.osPrinterManager=windowsPrinters.WindowsPrinters()
			self.printerPortName=self.printerName+' Port'
			makeDefault=False
			comment='Virtual printer created in Python'
			self.osPrinterManager.addPrinter(self.printerName,ip,port,
				self.printerPortName,makeDefault,comment)
		else:
			print 'WARN: Auto printer installation not implemented for '+os.name

	def _uninstallPrinter(self):
		"""
		remove the printer
		"""
		if self.osPrinterManager:
			self.osPrinterManager.removePrinter(self.printerName)
			self.osPrinterManager.removePort(self.printerPortName)

	def run(self):
		"""
		server mainloop
		"""
		if self.running:
			return
		self.running=True
		self.keepGoing=True
		sock=socket.socket(socket.AF_INET, socket.SOCK_STREAM)
		sock.bind((self.ip,self.port))
		ip,port=sock.getsockname()
		print 'Opening',ip+':'+str(port)
		if self.autoInstallPrinter:
			self._installPrinter(ip,port)
		#sock.setblocking(0)
		sock.listen(1)
		while self.keepGoing:
			print '\nListening for incoming print job...'
			while self.keepGoing: # let select() yield some time to this thread
									#so we can detect ctrl+c and keepGoing change
				inputready,outputready,exceptready=select.select([sock],[],[],1.0)
				if sock in inputready:
					break
			if not self.keepGoing:
				continue
			print 'Incoming job... spooling...'
			conn,addr=sock.accept()
			if self.printCallbackFn is None:
				f=open('I_printed_this.ps','wb')
				while True:
					data=conn.recv(self.buffersize)
					if not data:
						break
					f.write(data)
					f.flush()
			elif True:
				buf=[]
				while True:
					data=str(conn.recv(self.buffersize))
					if not data:
						break
					buf.append(data)
				buf=''.join(buf)
				# get whatever meta info we can
				author=None
				title=None
				filename=None
				header='@'+buf.split('%!PS-',1)[0].split('@',1)[1]
				#print header
				for line in header.split('\n'):
					line=line.strip()
					if line.startswith('@PJL JOB NAME='):
						n=line.split('"',1)[1].rsplit('"',1)[0]
						if os.path.isfile(n):
							filename=n
						else:
							title=n
					elif line.startswith('@PJL COMMENT'):
						line=line.split('"',1)[1].rsplit('"',1)[0].split(';')
						for param in line:
							param=param.split(':',1)
							if len(param)>1:
								param[0]=param[0].strip().lower()
								param[1]=param[1].strip()
								if param[0]=='username':
									author=param[1]
								elif param[0]=='app filename':
									if title is None:
										if os.path.isfile(param[1]):
											filename=param[1]
										else:
											title=param[1]
				if title is None and filename!=None:
					title=filename.rsplit(os.sep,1)[-1].split('.',1)[0]
				self.printCallbackFn(buf,title=title,author=author,filename=filename)
			else:
				buf=[]
				printjobHeader=[]
				fillingBuf=False
				while True:
					data=str(conn.recv(self.buffersize))
					if not data:
						break
					if not fillingBuf:
						i=data.find('%!PS-')
						if i<0:
							printjobHeader.append(data)
						elif i==0:
							buf.append(data)
							fillingBuf=True
						else:
							printjobHeader.append(data[0:i])
							buf.append(data[i:])
							fillingBuf=True
					else:
						buf.append(data)
				if buf:
					self.printCallbackFn(''.join(buf))
			conn.close()
			time.sleep(0.1)


if __name__=='__main__':
	import sys
	port=9001
	ip='127.0.0.1'
	runit=True
	for arg in sys.argv[1:]:
		pass # TODO: do args
	ps=PrintServer(ip, port)
	ps.run()
