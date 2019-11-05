#!/usr/bin/env python
# -*- coding:utf-8 -*-
# @Time  : 2019/10/14 16:33
# @Author: qimeng
# @File  : pdf2word.py

import os
import re
import datetime
from html import unescape
from tika import parser
from configparser import ConfigParser
from docx import Document


starttime = datetime.datetime.now()
flag = input("是否需要处理换行？（输入1表示保留句号与大写开头单词结尾换行，2表示仅保留句号结尾换行，0表示不处理）\n")
flag2 = input("是否需要处理空格问题？（输入1表示处理空格问题，0表示不处理）\n")


def html_to_plain_text ( html ):
    text = re.sub('<head.*?>.*?</head>', '', html, flags=re.M | re.S | re.I)
    text = re.sub(r'<a\s.*?>', ' HYPERLINK ', text, flags=re.M | re.S | re.I)
    text = re.sub('<.*?>', '', text, flags=re.M | re.S)
    text = re.sub(r'-$\n', '', text, flags=re.M | re.S)
    text = re.sub(r'\n(\n+)', '---<<<///qm', text, flags=re.M | re.S)
    text = re.sub(r'(\s*\n)+', '', text, flags=re.M | re.S)
    text = re.sub('---<<<///qm', '\n', text, flags=re.M | re.S)

    if flag == '1':
        text = re.sub(r"([A-Z][a-zA-Z]+)\s*$\n", r"\1!qm! \n", text, flags=re.M | re.S)
        text = re.sub(r"(\.|\?|\!)$\n", r"\1 \n", text, flags=re.M | re.S)
        text = re.sub(r"(?<!((\.|\?|\!)\s))$\n", "", text, flags=re.M | re.S)
        text = re.sub(r"!qm!", r"", text, flags=re.M | re.S)

    if flag == '2':
        text = re.sub(r"(\.|\?|\!)$\n", r"\1 $\n", text, flags=re.M | re.S)
        text = re.sub(r"(?<!((\.|\?|\!)\s))$\n", "", text, flags=re.M | re.S)


    if flag2 == '1':
        text = re.sub(r'\s([a-zA-Z])\s([a-zA-Z])\s', r'\1\2', text, flags=re.M | re.S)

    return unescape(text)


def save_text_to_word(content1, file_path):
    doc = Document()
    for line in content1.split('\n'):
        paragraph = doc.add_paragraph()
        paragraph.add_run(remove_control_characters(line))
    doc.save(file_path)


def remove_control_characters(content2):
        mpa = dict.fromkeys(range(32))
        return content2.translate(mpa)


config_parser = ConfigParser()
config_parser.read('config.cfg')
config = config_parser['default']

tasks = []

for file in os.listdir(config['pdf_folder']):
    extension_name = os.path.splitext(file)[1]
    if extension_name != '.pdf':
        continue
    print("get a pdf ! ")
    file_name = os.path.splitext(file)[0]
    pdf_file = config['pdf_folder'] + '/' + file
    if flag2 == '1':
        file_name = 'S+' + file_name
    if flag == '1':
        file_name = 'L1+' + file_name
    if flag == '2':
        file_name = 'L2+' + file_name

    word_file = config['word_folder'] + '/' + file_name + '.docx'
    print(file_name,"\n",pdf_file,"\n",word_file)
    parse_entire_pdf = parser.from_file(pdf_file, xmlContent=True)
    parse_entire_pdf = parse_entire_pdf['content']
    content = html_to_plain_text(parse_entire_pdf)
    # save_text_to_word(parse_entire_pdf, word_file)
    save_text_to_word(content, word_file)


'''
pdf_file = 'D:/pdf2word/pdf/图论模板.pdf'
word_file = 'D:/pdf2word/word/冷记忆.docx'


parse_entire_pdf = parser.from_file(pdf_file, xmlContent=True)
parse_entire_pdf = parse_entire_pdf['content']
content = html_to_plain_text(parse_entire_pdf)

save_text_to_word(content, word_file)

'''


endtime = datetime.datetime.now()

print ((endtime - starttime).seconds)
os.system("pause")