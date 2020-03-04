#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "???"

import argparse
import os
import zipfile


def search_dotm(text, dir):
    print("Searching directory {} for text '{}'".format(dir, text))
    filenames = os.listdir(dir)
    matches = 0
    searches = 0
    for filename in filenames:
        if filename.endswith(".dotm"):
            searches += 1
            true_path = os.path.join(dir, filename)
            with open(true_path) as f:
                zipped = zipfile.ZipFile(f)
                content = zipped.read('word/document.xml')
                text_index = content.find(text)
                if text_index >= 0:
                    matches += 1
                    match_line = content[text_index-40:text_index+41]
                    print('Match found in file ' + true_path)
                    print('    ...{}...'.format(match_line))

    print('Total dotm files matched: {}'.format(matches))
    print('Total dotm files searched: {}'.format(searches))


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("text", help="Text to search for within dotm file")
    parser.add_argument(
        "-d", "--dir", help="The directory to search within", default='.')
    return parser


def main():
    parser = create_parser()
    my_args = parser.parse_args()

    if not my_args:
        parser.print_usage()
        exit(1)

    search_dotm(my_args.text, my_args.dir)


if __name__ == '__main__':
    main()
