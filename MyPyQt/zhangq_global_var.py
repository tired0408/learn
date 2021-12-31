#!/usr/bin/python
# -*- coding: utf-8 -*-

def _init():
    global _global_dict
    _global_dict = {}


def Set_value(name, value):
    _global_dict[name] = value


def Get_value(name, defValue=None):
    try:
        return _global_dict[name]
    except KeyError:
        return defValue