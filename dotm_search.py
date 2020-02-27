#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Ybrayym Abamov"


import os
import argparse
from zipfile import ZipFile

files_searched = 0
matched_lines = 0


def create_parser():
    parser = argparse.ArgumentParser(description="search for a particular substring within .dotm files")
    parser.add_argument('--dir', default='.', help=".dotm file search")
    parser.add_argument('text', help="specify the text")
    return parser


def dir_analyze(dir_name, search_text):
    global files_searched
    global matched_lines
    for root, _, files in os.walk(dir_name):
        for name in files:
            if name.endswith('.dotm'):
                # print('Examining file: ' + name)
                files_searched += 1
                dot_m = ZipFile(os.path.join(root, name))
                # print(dot_m.namelist())
                content = dot_m.read('word/document.xml')
                if search_file(content, search_text):
                    print('Match found in file ' + name)
                    matched_lines += 1
    print('FILES SEARCHED: ' + str(files_searched))
    print('MATCHES FOUND: ' + str(matched_lines))


def search_file(text, search_text):
    for line in text.split('\n'):
        index = line.find(search_text)
        if index >= 0:
            print('       ... ' + line[index-40:index+40] + ' ...        ')
            return True
    return False


def main():
    parser = create_parser()
    args = parser.parse_args()
    dir_analyze(args.dir, args.text)
    # print(args)


if __name__ == '__main__':
    main()
