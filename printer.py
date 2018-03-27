#!/usr/bin/env python
import subprocess
import os
import StringIO

# get the location of ghostscript
if os.name=='nt':
	GHOSTSCRIPT_APP=None
	for gsDir in [os.environ['ProgramFiles']+os.sep+'gs',os.environ['ProgramFiles(x86)']+os.sep+'gs']: 
		if os.path.isdir(gsDir):
			# find the newest version
			bestVersion=0
			for f in os.listdir(gsDir):
				path=gsDir+os.sep+f
				if os.path.isdir(path) and f.startswith('gs'):
					try:
						val=float(f[2:])
					except:
						val=0
					if bestVersion<val:
						for appName in ['gswin64c.exe','gswin32c.exe','gswin.exe','gs.exe']:
							appName=gsDir+os.sep+f+os.sep+'bin'+os.sep+appName
							if os.path.isfile(appName):
								bestVersion=val
								GHOSTSCRIPT_APP='"'+appName+'"'
								break
	if GHOSTSCRIPT_APP==None:
		raise Exception('ERR: Ghostscript not found!\n\nYou can get it from:\n\thttp://www.ghostscript.com')
else: # assume we can find it
	GHOSTSCRIPT_APP='gs'
print 'GHOSTSCRIPT_APP='+GHOSTSCRIPT_APP

def shell_escape(param):
	try:
		import shlex
		return shlex.quote(param)
	except ImportError:
		import pipes
		return pipes.quote(param)

class Printer:
	"""
	You can derive from this class to create your own printer!
	
	Simply send in the options you want in Printer.__init__
	and then override printThis() to do what you want.
	DONE!
	Ready to run it with run()
	"""
	
	def __init__(self,name='My Virtual Printer',acceptsFormat='png',acceptsColors='rgba'):
		"""
		name - the name of the printer to be installed
		
		acceptsFormat - the format that the printThis() method accepts
		Available formats are "pdf", or "png" (default=png)
		
		acceptsColors - the color format that the printThis() method accepts (if relevent to acceptsFormat)
		Available colors are "grey", "rgb", or "rgba" (default=rgba)
		"""
		self._server=None
		self.name=name
		self.acceptsFormat=acceptsFormat
		self.acceptsColors=acceptsColors
		self.bgColor='#ffffff' # not sure how necessary this is

	def printThis(self,doc,title=None,author=None,filename=None):
		"""
		you probably want to override this
		
		called when something is being printed
		
		defaults to saving a file
		"""
		if title==None:
			title='printed'
		if author!=None:
			title=title+' - '+author
		f=open(shell_escape(title+'.'+self.acceptsFormat),'wb')
		f.write(doc)
		f.close()
		
	def run(self,host='127.0.0.1',port=None,autoInstallPrinter=True):
		"""
		normally all the default values are exactly what you need!
		
		autoInstallPrinter is used to install the printer in the operating system (currently only supports Windows)
			startServer is required for this
		"""
		import printServer
		self._server=printServer.PrintServer(self.name,host,port,autoInstallPrinter,self.printPostscript)
		self._server.run()
		del self._server # delete it so it gets un-registered
		self._server=None
				
	def _postscriptToFormat(self,data,gsDev='pdfwrite',gsDevOptions=[],outputDebug=True):
		"""
		Converts postscript data (in a string) to pdf data (in a string)
		
		gsDev is a ghostscript format device
		
		For ghostscript command line, see also:
			http://www.ghostscript.com/doc/current/Devices.htm
			http://www.ghostscript.com/doc/current/Use.htm#Options
		"""
		cmd=GHOSTSCRIPT_APP+' -q -sDEVICE='+gsDev+' '+(' '.join(gsDevOptions))+' -sstdout=%stderr -sOutputFile=- -dBATCH -'
		if outputDebug:
			print cmd
		data,gsStdoutStderr=subprocess.Popen(cmd,stdin=subprocess.PIPE,stderr=subprocess.PIPE,stdout=subprocess.PIPE,shell=True).communicate(input=data)
		if outputDebug:
			print gsStdoutStderr # note: stdout also goes to stderr because of the -sstdout=%stderr flag
		return data
		
	def printPostscript(self,datasource,datasourceIsFilename=False,title=None,author=None,filename=None):
		"""
		datasource is either:
			a filename
			None to get data from stdin
			a file-like object
			something else to convert using str() and then print
		Keep in mind that it MUST contain postscript data
		"""
		# -- open data source
		data=None
		if datasource==None:
			data=sys.stdin.read()
		elif type(datasource)==str or type(datasource)==unicode:
			if datasourceIsFilename:
				f=open(datasource,'rb')
				data=f.read()
				f.close()
				if title==None:
					title=datasource.rsplit(os.sep,1)[-1].rsplit('.',1)[0]
			else:
				data=datasource
		elif hasattr(datasource,'read'):
			data=datasource.read()
		else:
			data=str(datasource)
		# -- convert the data to the required format
		print 'Converting data...'
		gsDevOptions=[]
		if self.acceptsFormat=='pdf':
			gsDev='pdfwrite'
		elif self.acceptsFormat=='png':
			gsDevOptions.append('-r600')
			gsDevOptions.append('-dDownScaleFactor=3')
			if self.acceptsColors=='grey':
				gsDev='pnggray'
			elif self.acceptsColors=='rgba':
				if self.bgColor==None: # I'm not sure how necessary background color is
					self.bgColor='#ffffff'
				gsDev='pngalpha'
				if self.bgColor!=None:
					if self.bgColor[0]!='#':
						self.bgColor='#'+self.bgColor
					gsDevOptions.append('-dBackgroundColor=16'+self.bgColor)
			#elif self.acceptsColors=='rgb': #TODO
			else:
				raise Exception('Unacceptable color format "'+self.acceptsColors+'"')
		else:
			raise Exception('Unacceptable data type format "'+self.acceptsFormat+'"')
		data=self._postscriptToFormat(data,gsDev,gsDevOptions)
		# -- send the data to the printThis function
		print 'Printing data...'
		self.printThis(data,title=title,author=author,filename=filename)


if __name__=='__main__':
	import os,sys
	usage="""
	USAGE:
		virtualPrinter filename.ps ..... to print a file
		virtualPrinter - ............... to print postscript piped in from stdin
		virtualPrinter ip[:port]........ to start a print server
	NOTE:
		you can do multiple commands with the same virtualPrinter
	"""
	if len(sys.argv)<2:
		print usage
	else:
		p=Printer()
		for arg in sys.argv[1:]:
			if arg=='-':
				p.printPostscript()
			elif arg.find(os.sep)<0 and len(sys.argv[1].split('.'))>2:
				# looks like an ip to me!
				ip=arg.rsplit(':',1)
				if len(ip)>1:
					port=ip[-1]
				else:
					port=None
				ip=ip[0]
				p.run(ip,port)
			else:
				p.printPostscript(arg,True)
	