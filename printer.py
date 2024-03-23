#!/usr/bin/env
# -*- coding: utf-8 -*-
"""
Simply send in the options you want in Printer.__init__
and then override printThis() to do what you want.
DONE!
Ready to run it with run()
"""
import typing
import os
import sys
import subprocess

from virtualPrinter.printServer import PrintCallbackDocType
from virtualPrinter.printerException import PrinterException


# get the location of ghostscript
if os.name=='nt':
    GHOSTSCRIPT_APP=None
    for gsDir in (
        os.environ['ProgramFiles']+os.sep+'gs',
        os.environ['ProgramFiles(x86)']+os.sep+'gs'
        ):
        if os.path.isdir(gsDir):
            # find the newest version
            bestVersion=0
            for f in os.listdir(gsDir):
                path=gsDir+os.sep+f
                if os.path.isdir(path) and f.startswith('gs'):
                    try:
                        val=float(f[2:])
                    except ValueError:
                        val=0
                    if bestVersion<val:
                        for appName in (
                            'gswin64c.exe',
                            'gswin32c.exe',
                            'gswin.exe',
                            'gs.exe'
                            ):
                            appName=gsDir+os.sep+f+os.sep+'bin'+os.sep+appName
                            if os.path.isfile(appName):
                                bestVersion=val
                                GHOSTSCRIPT_APP='"'+appName+'"'
                                break
    if GHOSTSCRIPT_APP is None:
        errString="""ERR: Ghostscript not found!
            You can get it from:
                http://www.ghostscript.com"""
        raise PrinterException(errString)
else: # assume we can find it
    GHOSTSCRIPT_APP='gs'
print(f'GHOSTSCRIPT_APP={GHOSTSCRIPT_APP}')


# find a good shell_escape routine
try:
    import shlex
    if hasattr(shlex,'quote'):
        shell_escape=shlex.quote
    else:
        import pipes # pylint: disable=deprecated-module
        shell_escape=pipes.quote
except ImportError:
    import pipes # pylint: disable=deprecated-module
    shell_escape=pipes.quote


class Printer:
    """
    You can derive from this class to create your own printer!

    Simply send in the options you want in Printer.__init__
    and then override printThis() to do what you want.
    DONE!
    Ready to run it with run()
    """

    def __init__(self,
        name:str='My Virtual Printer',
        acceptsFormat:str='png',
        acceptsColors:str='rgba'):
        """
        name - the name of the printer to be installed

        acceptsFormat - the format that the printThis() method accepts
        Available formats are "pdf", or "png" (default=png)

        acceptsColors - the color format that the printThis() method accepts
        (if relevent to acceptsFormat)
        Available colors are "grey", "rgb", or "rgba" (default=rgba)
        """
        self._server=None
        self.name=name
        self.acceptsFormat=acceptsFormat
        self.acceptsColors=acceptsColors
        self.bgColor='#ffffff' # not sure how necessary this is

    def printThis(self,
        doc:PrintCallbackDocType,
        title:typing.Optional[str]=None,
        author:typing.Optional[str]=None,
        filename:typing.Optional[str]=None
        )->None:
        """
        you probably want to override this

        called when something is being printed

        defaults to saving a file

        TODO: keep track of filename?
        """
        _=filename # For now, don't care
        if title is None:
            title='printed'
        if author is not None:
            title=title+' - '+author
        with open(shell_escape(title+'.'+self.acceptsFormat),'wb') as f:
            f.write(doc)

    def run(self,
        host:str='127.0.0.1',
        port:typing.Union[None,int,str]=None,
        autoInstallPrinter:bool=True
        )->None:
        """
        normally all the default values are exactly what you need!

        autoInstallPrinter is used to install the printer in the OS
        (currently only supports Windows)
        startServer is required for this
        """
        from virtualPrinter.printServer import PrintServer
        self._server=PrintServer(
            self.name,host,port,autoInstallPrinter,self.printPostscript)
        self._server.run()
        del self._server # delete it so it gets un-registered
        self._server=None

    def _postscriptToFormat(self,
        data,
        gsDev:str='pdfwrite',
        gsDevOptions:typing.Optional[typing.Iterable[str]]=None,
        outputDebug:bool=True
        )->str:
        """
        Converts postscript data (in a string) to pdf data (in a string)

        gsDev is a ghostscript format device

        For ghostscript command line, see also:
            http://www.ghostscript.com/doc/current/Devices.htm
            http://www.ghostscript.com/doc/current/Use.htm#Options
        """
        if GHOSTSCRIPT_APP is None:
            msg="ghostscript is inaccessible. Is it installed?"
            raise PrinterException(msg)
        cmd:typing.List[str]=[GHOSTSCRIPT_APP,'-q','-sDEVICE='+gsDev]
        if gsDevOptions is not None:
            cmd.extend(gsDevOptions)
        cmd.extend((r'-sstdout=%stderr','-sOutputFile=-','-dBATCH','-'))
        if outputDebug:
            print(' '.join(cmd))
        with subprocess.Popen(cmd,
            stdin=subprocess.PIPE,stderr=subprocess.PIPE,
            stdout=subprocess.PIPE,shell=True) as po:
            data,gsStdoutStderr=po.communicate(input=data)
        if outputDebug:
            # note: stdout also goes to stderr because of
            # the -sstdout=%stderr flag above
            print(gsStdoutStderr)
        return data

    def printPostscript(self,
        datasource:PrintCallbackDocType,
        datasourceIsFilename:bool=False,
        title:typing.Optional[str]=None,
        author:typing.Optional[str]=None,
        filename:typing.Optional[str]=None
        )->None:
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
        if datasource is None:
            data=sys.stdin.read()
        elif isinstance(datasource,str):
            if datasourceIsFilename:
                with open(datasource,
                    'rb',
                    encoding='utf-8',
                    errors='ignore') as f: # noqa: E129
                    data=f.read()
                if title is None:
                    title=datasource.rsplit(os.sep,1)[-1].rsplit('.',1)[0]
            else:
                data=datasource
        elif hasattr(datasource,'read'):
            data=datasource.read()
        else:
            data=str(datasource)
        # -- convert the data to the required format
        print('Converting data...')
        gsDevOptions=[]
        if self.acceptsFormat=='pdf':
            gsDev='pdfwrite'
        elif self.acceptsFormat=='png':
            gsDevOptions.append('-r600')
            gsDevOptions.append('-dDownScaleFactor=3')
            if self.acceptsColors=='grey':
                gsDev='pnggray'
            elif self.acceptsColors=='rgba':
                if self.bgColor is None: # how necessary is this?
                    self.bgColor='#ffffff'
                gsDev='pngalpha'
                if self.bgColor is not None:
                    if self.bgColor[0]!='#':
                        self.bgColor='#'+self.bgColor
                    gsDevOptions.append('-dBackgroundColor=16'+self.bgColor)
            #elif self.acceptsColors=='rgb': #TODO
            else:
                msg=f'Unknown color format "{self.acceptsColors}"'
                raise PrinterException(msg)
        else:
            msg=r'Unacceptable data type format "{self.acceptsFormat}"'
            raise PrinterException(msg)
        data=self._postscriptToFormat(data,gsDev,gsDevOptions)
        # -- send the data to the printThis function
        print('Printing data...')
        self.printThis(data,title=title,author=author,filename=filename)


if __name__=='__main__':
    usage="""
    USAGE:
        virtualPrinter filename.ps ..... print a file
        virtualPrinter - ............... print postscript piped in from stdin
        virtualPrinter ip[:port]........ start a print server
    NOTE:
        you can do multiple commands with the same virtualPrinter
    """
    if len(sys.argv)<2:
        print(usage)
    else:
        p=Printer()
        for arg in sys.argv[1:]:
            if arg=='-':
                p.printPostscript(sys.stdin)
            elif arg.find(os.sep)<0 and len(sys.argv[1].split('.'))>2:
                # looks like an ip to me!
                ipPort=arg.rsplit(':',1)
                port:typing.Union[None,str,int]
                if len(ipPort)>1:
                    port=ipPort[-1]
                else:
                    port=None
                ip=ipPort[0]
                p.run(ip,port)
            else:
                p.printPostscript(arg,True)
