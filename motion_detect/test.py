# -*- coding: utf-8 -*-



import os, sys
import time
import GlobalCongif

def testGlobal():
    print GlobalCongif.g_help

if __name__ == "__main__":
    print(os.path.abspath("__file__"))
    print(os.getcwd())
    print os.path.dirname(sys.argv[0])