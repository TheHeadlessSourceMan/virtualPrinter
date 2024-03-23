#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
A virtual printer device that creates an image file as output
"""
import typing
import os
from virtualPrinter import Printer,PrintCallbackDocType
try:
    from PIL import Image
except ImportError as e:
    print('Requires PIL (Python Imaging Library)')
    print('Install it with the command:')
    print('  pip install pillow')
    raise e


class MyPrinter(Printer):
    """
    A virtual printer device that creates an image file as output
    """

    def __init__(self):
        Printer.__init__(self,'Print to Image',acceptsFormat='png')

    def doc2pil(self,doc):
        """
        Convert the doc buffer to a PIL image
        """
        from io import StringIO
        f=StringIO(doc)
        img=Image.open(f)
        return img

    def doSaveAsDialog(self,
        originFilename:typing.Optional[str]
        )->typing.Optional[str]:
        """
        Quick ans simple way to toss up a "save as" dialog.

        :param originFilename: for convenience, we'll guess at an output
            filename based on this (foo.html -> foo.pdf)
        """
        import tkinter as tk
        tk.Tk().withdraw() # prevent a blank application window
        filetypes=(
            ("PNG Files",'*.png'),
            ("JPEG Files",'*.jpeg;*.jpg;*.jpe'),
            ("BMP Files",'*.bmp'),
            ("all files","*.*"))
        if originFilename is None:
            originFilename='unknown'
        originPathFilename=originFilename.rsplit(os.sep,1)
        initialfile=originPathFilename[-1].rsplit('.',1)[0]+'.png'
        if len(originPathFilename)>1:
            initialdir=originPathFilename[0]
        else:
            initialdir=None
        val=tk.filedialog.asksaveasfilename(confirmoverwrite=True,
            title='Save As...',defaultextension='.png',filetypes=filetypes,
            initialfile=initialfile,initialdir=initialdir)
        if val is not None and val.strip()!='':
            if os.sep!='/':
                val=val.replace('/',os.sep)
            return val
        return None

    def printThis(self,
        doc:PrintCallbackDocType,
        title:typing.Optional[str]=None,
        author:typing.Optional[str]=None,
        filename:typing.Optional[str]=None
        )->None:
        """
        Called whenever something is being printed.

        We'll save the input doc to a book or book format
        """
        print("Printing:")
        print("\tTitle:",title)
        print("\tAuthor:",author)
        print("\tFilename:",filename)
        val=self.doSaveAsDialog(filename)
        if val is not None:
            print("Saving to:",val)
            img=self.doc2pil(doc)
            img.save(val)
        else:
            print("No output filename.  So never mind then, I guess.")


if __name__=='__main__':
    # Simply run the printer
    p=MyPrinter()
    print('Starting printer... [CTRL+C to stop]')
    p.run()
