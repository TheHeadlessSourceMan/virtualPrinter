#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
A virtual printer device that creates a pdf as output
"""
from virtualPrinter.printer import Printer
import os


class MyPrinter(Printer):
	"""
	A virtual printer device that creates a pdf as output
	"""
	
	def __init__(self):
		Printer.__init__(self,'Print to PDF',acceptsFormat='pdf')
	
	def doSaveAsDialog(self,originFilename):
		"""
		Quick ans simple way to toss up a "save as" dialog.
		
		originFilename - for convenience, we'll guess at an output filename based on this (foo.html -> foo.pdf)
		"""
		import tkFileDialog
		import Tkinter
		Tkinter.Tk().withdraw() # this trick prevents a blank application window from showing up
		filetypes=(("PDF Files",'*.pdf'),("all files","*.*"))
		originFilename=originFilename.rsplit(os.sep,1)
		initialfile=originFilename[-1].rsplit('.',1)[0]+'.pdf'
		if len(originFilename)>1:
			initialdir=originFilename[0]
		else:
			initialdir=None
		val=tkFileDialog.asksaveasfilename(confirmoverwrite=True,title='Save As...',defaultextension='.pdf',filetypes=filetypes,initialfile=initialfile,initialdir=initialdir)
		if val!=None and val.strip()!='':
			if os.sep!='/':
				val=val.replace('/',os.sep)
			return val
		return None
	
	def printThis(self,doc,title=None,author=None,filename=None):
		"""
		Called whenever something is being printed.
		
		We'll save the input doc to a book or book format
		"""
		print "Printing:"
		print "\tTitle:",title
		print "\tAuthor:",author
		print "\tFilename:",filename
		val=self.doSaveAsDialog(filename)
		if val!=None:
			print "Saving to:",val
			f=open(val,'wb')
			f.write(doc)
			f.close()
		else:
			print "No output filename.  So never mind then, I guess."
		
		
if __name__=='__main__':
	"""
	Simply run the printer
	"""
	p=MyPrinter()
	print 'Starting printer... [CTRL+C to stop]'
	p.run()	