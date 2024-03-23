[![Pylint](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/pylint.yml/badge.svg)](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/pylint.yml)[![Flake8](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/flake8.yml/badge.svg)](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/flake8.yml)[![MyPy](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/mypy.yml/badge.svg)](https://github.com/TheHeadlessSourceMan/virtualPrinter/actions/workflows/mypy.yml)
# virtualPrinter

This library allows you to easily create a windows virtual printer.

## How to use
 * Simply create a ```Printer('my printer name',acceptsFormat='png')``` object, and implement its ```printThis(doc,title=None,author=None,filename=None)``` method.
   * It should show up in your list of windows printers.
   * Every print job ultimately calls your ```printThis()``` with a new doc in the ```acceptsFormat``` format
 * see the [examples](./examples) directory for details

## Theory of Operation

1. Open a TCP server on a (loopback) adapter
2. Tell Windows to install a PostScript network printer that lives at that address:port
3. Whenever a print job comes in on that network port, call commandline GhostScript (required) to convert the PostScript into a PDF (works great, since the two formats are closely related)
4. Use PIL to read that PDF into an image
5. User code gets passed a normal, everyday, PIL image, plus some meta info (title,user,etc) gleaned from the original PostScript
