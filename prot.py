# coding:utf-8

from prot4 import getdata

if __name__ == '__main__':
  params = {
      "slide_sum":12,#新規pptxファイルのスライド数　将来的にはforeachで回すので不要になる
      "template":"L.pptx",
      "sourse":"test.json"
      }
  a=getdata(params)
  a.main()
  a.save()
