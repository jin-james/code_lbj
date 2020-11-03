# from docx import Document
# import pandoc
import os

doc = 'C:\\Users\\j20687\\Desktop\\demo.docx'
tex = 'demo.html'
os.system('pandoc -s %s --metadata title:"title" -o %s' % (doc, tex))
# blip = doc.inline_shapes[0]._inline.graphic.graphicData.pic.blipFill.blip
# rID = blip.embed
# document_part = doc.part
# image_part = document_part.related_parts[rID]


# fr = open("test.png", "wb")
# fr.write(image_part._blob)
# fr.close()
# from xml.dom.minidom import parse, Document
# dom1 = parse(r"C:\Users\j20687\Desktop\demotest.xml")   # parse an XML file
# print(dom1.toxml())

