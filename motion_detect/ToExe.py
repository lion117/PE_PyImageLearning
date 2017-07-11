#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on 2014/8/21
@author: LEO
"""

from distutils.core import setup
import py2exe



setup(console=["CatchThiefFilter.py"],
      options={"py2exe":{
                  "compressed": True ,
                  "optimize": 2,
                  "bundle_files": 2 ,
              }
         })


