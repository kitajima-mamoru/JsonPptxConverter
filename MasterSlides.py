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
    #レイアウトをコピーし別名で保存　将来的には設定用ディレクトリを作りそちらに格納する
    Presentation(template).save(self.__get_name())
    self.__prs = Presentation(self.__get_name())

  def __get_name(self):
    #pptxファイルの名前
    return self.__master_source['0']['presentation_name']
    
  def __get_source(self,slide_number):
    #pptxファイルのソース取得
    return self.__master_source.get(str(slide_number))
    
  def make_slide(self,slide_number):
    #スライド1枚を生成するのに必要なソースと共にslideインスタンスを生成する
    Slide(self.__prs,self.__get_source(slide_number))

  def save(self):
    self.__prs.save(self.__get_name())
