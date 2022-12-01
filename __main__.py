from converter.converter import PyConverter as Converter

__name__ = "pyconverter"
parser = Converter().parseArguments()
parser.readAndConvert()