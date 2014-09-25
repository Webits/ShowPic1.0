<%

'-----要加寫優先權（只有W or H）

p_width=request("w")
p_height=request("h")
Set Jpeg = Server.CreateObject("Persits.Jpeg")
    
Jpeg.Open Server.MapPath(request("src"))


'---Resize 模式： 在W/H的寬高範圍內限縮，以長邊為主，例如限高300的立式圖片，縮完後高度就是300，而寬度會變更小
if request("mode")="" or request("mode")="resize" then

	'---長型圖片（直幅）
	If Jpeg.OriginalHeight >= Jpeg.OriginalWidth Then

		Jpeg.height =p_height
		Jpeg.width = p_height * Jpeg.OriginalWidth / Jpeg.OriginalHeight
		
	
	else
	'---寬型圖片（橫幅）

		Jpeg.width =p_width
		Jpeg.height =p_width * Jpeg.OriginalHeight/Jpeg.OriginalWidth

	End If
	
	Jpeg.Crop 0, 0, Jpeg.width, Jpeg.height
	Jpeg.SendBinary


'---Crop 模式，將原圖以W/H的範圍內縮小後裁切，例如限高寬300的立式圖片，裁切完後寬就是300，但是高會被裁成300(正方形）並且置中
elseif request("mode")="crop" Then

			  
				SourceAspectRatio = Jpeg.OriginalWidth / Jpeg.OriginalHeight
				DesiredAspectRatio = p_width / p_height
			  
				if SourceAspectRatio > DesiredAspectRatio then
					'
					' Triggered when source image is wider
					'
					Jpeg.Height = p_height
					Jpeg.Width = p_height * SourceAspectRatio
				else
					'
					' Triggered otherwise (i.e. source image is similar or taller)
					'
					Jpeg.Width = p_width
					Jpeg.Height = p_width / SourceAspectRatio
				end if
			  

				'
				' Crop the image keeping the CENTER part intact
				'
				X0 = (Jpeg.Width - p_width) / 2
				Y0 = (Jpeg.Height - p_height) / 2
				X1 = X0 + p_width
				Y1 = Y0 + p_height
			  
				Jpeg.Crop X0, Y0, X1, Y1

				Jpeg.SendBinary



end if



%>
