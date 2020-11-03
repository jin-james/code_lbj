#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys, os

# Set your own model path
MODELDIR = r'D:\Program Files\python-ltp\ltp_data'

from pyltp import SentenceSplitter, Segmentor, Postagger, Parser, NamedEntityRecognizer, SementicRoleLabeller

paragraph = '1912年1月1日，孙中山在南京宣誓就职，中华民国正式成立。'

sentence = SentenceSplitter.split(paragraph)[0]  # 分句

segmentor = Segmentor()
segmentor.load(os.path.join(MODELDIR, "cws.model"))  # 分词
words = segmentor.segment(sentence)
print("\t".join(words))

postagger = Postagger()
postagger.load(os.path.join(MODELDIR, "pos.model"))  # 词性标注
postags = postagger.postag(words)
# list-of-string parameter is support in 0.1.5
# postags = postagger.postag(["中国","进出口","银行","与","中国银行","加强","合作"])
print("\t".join(postags))

parser = Parser()
parser.load(os.path.join(MODELDIR, "parser.model"))  # 依存句法分析
arcs = parser.parse(words, postags)

print("\t".join("%d:%s" % (arc.head, arc.relation) for arc in arcs))

recognizer = NamedEntityRecognizer()
recognizer.load(os.path.join(MODELDIR, "ner.model"))  # 命名实体识别
netags = recognizer.recognize(words, postags)
print("\t".join(netags))

labeller = SementicRoleLabeller()
labeller.load(os.path.join(MODELDIR, "pisrl_win.model"))
# arcs 使用依存句法分析的结果
roles = labeller.label(words, postags, arcs)  # 语义角色标注

for role in roles:
    print(role.index, "".join(
            ["%s:(%d,%d)" % (arg.name, arg.range.start, arg.range.end) for arg in role.arguments]))

segmentor.release()
postagger.release()
parser.release()
recognizer.release()
labeller.release()
