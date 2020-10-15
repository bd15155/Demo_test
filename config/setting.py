#!/usr/bin/env python
# _*_ coding:utf-8 _*_
__author__ = 'Administrator'


import os,sys
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
sys.path.append(BASE_DIR)

CONFIG_DIR = os.path.join(BASE_DIR,"database","user.ini")
TEST_DIR = os.path.join(BASE_DIR,"testcase")
TEST_REPORT = os.path.join(BASE_DIR,"report")
LOG_DIR = os.path.join(BASE_DIR,"logs")
