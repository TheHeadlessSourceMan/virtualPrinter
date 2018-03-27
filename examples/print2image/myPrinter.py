#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
A virtual printer device that creates an image file as output
"""
from virtualPrinter.printer import Printer
import os
from PIL import Image

class MyPrinter(Printer):
	"""
	A virtual printer device that creates an image file as output
	"""
	
	def __init__(self):
		Printer.__init__(self,'Print to Image',acceptsFormat='png')
		
	def doc2pil(self,doc):
		from cStringIO import StringIO
		f=StringIO(doc)
		img=Image.open(f)
		return img
	
	def doSaveAsDialog(self,originFilename):
		"""
		Quick ans simple way to toss up a "save as" dialog.
		
		originFilename - for convenience, we'll guess at an output filename based on this (foo.html -> foo.pdf)
		"""
		import tkFileDialog
		import Tkinter
		Tkinter.Tk().withdraw() # this trick prevents a blank application window from showing up
		filetypes=(("PNG Files",'*.png'),("JPEG Files",'*.jpeg;*.jpg;*.jpe'),("BMP Files",'*.bmp'),("all files","*.*"))
		originFilename=originFilename.rsplit(os.sep,1)
		initialfile=originFilename[-1].rsplit('.',1)[0]+'.png'
		if len(originFilename)>1:
			initialdir=originFilename[0]
		else:
			initialdir=None
		val=tkFileDialog.asksaveasfilename(confirmoverwrite=True,title='Save As...',defaultextension='.png',filetypes=filetypes,initialfile=initialfile,initialdir=initialdir)
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
			img=self.doc2pil(doc)
			img.save(val)
		else:
			print "No output filename.  So never mind then, I guess."
		
		
if __name__=='__main__':
	"""
	Simply run the printer
	"""
	p=MyPrinter()
	print 'Starting printer... [CTRL+C to stop]'
	p.run()	