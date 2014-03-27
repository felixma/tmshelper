#!/usr/bin/env python
from optparse import OptionParser
import sys

def main(argv):
    parser = OptionParser()
    parser.add_option("-f", "--file",  
                          action="store", type="string", dest="filename")  
    #args = ["-f", "foo.txt"]  
    (options, args) = parser.parse_args(argv)  
    print options.filename  
    #print options
    #print args

if __name__ == '__main__':
    main(sys.argv[1:])

