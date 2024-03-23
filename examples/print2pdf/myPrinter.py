#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
A virtual printer device that creates a pdf as output
"""
import typing
import os
from virtualPrinter import Printer,PrintCallbackDocType


class MyPrinter(Printer):
    """
    A virtual printer device that creates a pdf as output
    """

    def __init__(self):
        Printer.__init__(self,'Print to PDF',acceptsFormat='pdf')

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
        filetypes=(("PDF Files",'*.pdf'),("all files","*.*"))
        if originFilename is None:
            originFilename='unknown'
        originPathName=originFilename.rsplit(os.sep,1)
        initialfile=originPathName[-1].rsplit('.',1)[0]+'.pdf'
        if len(originPathName)>1:
            initialdir=originPathName[0]
        else:
            initialdir=None
        val=tk.filedialog.asksaveasfilename(confirmoverwrite=True,
            title='Save As...',defaultextension='.pdf',filetypes=filetypes,
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
            with open(val,'wb') as f:
                f.write(doc)
        else:
            print("No output filename.  So never mind then, I guess.")


if __name__=='__main__':
    # Simply run the printer
    p=MyPrinter()
    print('Starting printer... [CTRL+C to stop]')
    p.run()
