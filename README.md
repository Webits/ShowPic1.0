ShowPic1.0
==========

自動裁圖顯示 1.0

這是 ASP 引用 ASPJPEG 作出來的圖片顯示元件。
可以取代原本 img 的 src 路徑。

例如原本圖片是：

<img src="a.jpg">   

但是這張圖片比例跟尺寸都太大。

那就可以使用這支程式，並將標籤改為：

<img src="pics.asp?參數...>

這樣這張圖片就會以 ASP 產生的東西餵出來，以我們要的設定方式呈現。

====================


使用方法：

pics.asp?mode=[模式]&w=[寬]&h=[高]&src=[圖片路徑]

範例：

pics.asp?w=200&h=200&mode=resize&src=program/cmx/dir_pics/A1850.jpg


參數說明：
[模式] = crop  / resize
[w]=寬度
[h]=高度
[src]=圖片路徑

Crop 會自動將圖片的長邊調整為參數中的短邊，然後依照參數中的比例刪除不要的部分。
Resize 則會將圖片的長邊調整為參數中的長邊，然後比例上多餘的部分會補白。
大多情況下我們會選用 Crop ，而不會使用 Resize。










