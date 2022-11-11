Sub setpicsize() '设置图片尺寸

Dim n '图片个数

On Error Resume Next '忽略错误

For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes 类型 图片

ActiveDocument.InlineShapes(n).Height = 170.1 '设置图片高度为 6cm
'ActiveDocument.InlineShapes(n).Width = 283.5 '设置图片宽度 10cm

' Word中的尺寸单位默认是cm（厘米），而1cm等于28.35px（像素），
' 由于代码中换算设置的单位是px（像素）。所以就用尺寸高度或宽度值乘像素值。
' 即为：7*28.35=198.45；宽度换算方法与此相同。


Next n

End Sub