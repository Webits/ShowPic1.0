<%

'-----�n�[�g�u���v�]�u��W or H�^

p_width=request("w")
p_height=request("h")
Set Jpeg = Server.CreateObject("Persits.Jpeg")
    
Jpeg.Open Server.MapPath(request("src"))


'---Resize �Ҧ��G �bW/H���e���d�򤺭��Y�A�H���䬰�D�A�Ҧp����300���ߦ��Ϥ��A�Y���ᰪ�״N�O300�A�Ӽe�׷|�ܧ�p
if request("mode")="" or request("mode")="resize" then

	'---�����Ϥ��]���T�^
	If Jpeg.OriginalHeight >= Jpeg.OriginalWidth Then

		Jpeg.height =p_height
		Jpeg.width = p_height * Jpeg.OriginalWidth / Jpeg.OriginalHeight
		
	
	else
	'---�e���Ϥ��]��T�^

		Jpeg.width =p_width
		Jpeg.height =p_width * Jpeg.OriginalHeight/Jpeg.OriginalWidth

	End If
	
	Jpeg.Crop 0, 0, Jpeg.width, Jpeg.height
	Jpeg.SendBinary


'---Crop �Ҧ��A�N��ϥHW/H���d���Y�p������A�Ҧp�����e300���ߦ��Ϥ��A��������e�N�O300�A���O���|�Q����300(����Ρ^�åB�m��
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
