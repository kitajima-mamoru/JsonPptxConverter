# coding:utf-8
import json
from pptx import Presentation
from pptx.dml.color import RGBColor
from collections import OrderedDict
from pptx.util import Pt #Inches
from Slide import Slide

class MasterSlides():
  def __init__(self,template,source):
    #jsonをロード
    self.__master_source = json.load(open(source))
    #新規pptxファイルの名前
    self.__presentation_name = self.__master_source['0']['presentation_name']
    #レイアウトをコピーし別名で保存　将来的には設定用ディレクトリを作りそちらに格納する
    Presentation(template).save(self.__presentation_name)
    self.__prs = Presentation(self.__presentation_name)

  def make_slide(self,slide_number):
    #slidemasterからlayoutを指定しつつslideを追加
    self.__slide_source =self.__master_source[str(slide_number)]
    self.__shapes = self.__prs.slides.add_slide(self.__prs.slide_layouts[self.__slide_source['layout_number']]).shapes
    Slide(self.__shapes,self.__slide_source)

  def save(self):
    self.__prs.save(self.__presentation_name)
