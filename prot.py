# coding:utf-8

from MasterSlides import MasterSlides

if __name__ == '__main__':
  slide_sum =12#新規pptxファイルのスライド数　将来的にはforeachで回すので不要になる
  template = "L.pptx"
  source = "test.json"

  presen=MasterSlides(template,source)
    #各ページを追加、保存していく
  for slide_number in range(1,slide_sum+1):
    presen.make_slide(slide_number)
  presen.save()
