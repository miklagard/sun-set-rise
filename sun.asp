<%
' Source: http://www.paulsadowski.com/wsh/suntimes.htm

Class Sun
	Private Pi 
	Private Radeg
	Private Degrad
	
	Private Sub Class_Initialize()
		pi = 3.1415926535897932384626433832795
		RADEG = 180 / Pi
		DEGRAD = Pi / 180
	End Sub 

	Public Function SunTimes (ReqDay, longitude, latitude, TZ, isdst)
		Dim d, n, i, w, m, l, e, e1, a, xv, yv, v, xs, ys, xe, ecl, lonsun, ye, ze, ra, dec, h 
		Dim GMST0, UT_Sun_in_south, LHA, hour_rise, hour_set, min_rise, min_set
		Dim Ret 
		
		Set Ret = Server.CreateObject("Scripting.Dictionary")
		
		strYear = Year(ReqDay)
		strMonth = Month(ReqDay)
		strDay = Day(ReqDay)
		
		'calculate days since 2000 jan 1
		d = (367 * (strYear) - int((7 * ((strYear) + (((strMonth) + 9) / 12))) / 4) + int((275 * (strMonth)) / 9) + (strDay) - 730530)

		' Orbital elements of the Sun:
		N = 0.0
		i = 0.0
		w = 282.9404 + 4.70935E-5 * d

		a = 1.000000
		e = 0.016709 - 1.151E-9 * d

		M = 356.0470 + 0.9856002585 * d
		M = rev(M)

		ecl = 23.4393 - 3.563E-7 * d
		L = w + M
		
		If (L < 0 OR L > 360) Then
			L = rev(L)
		End If

		' position of the Sun
		E1 = M + e*(180/pi) * sind(M) * ( 1.0 + e * cosd(M) )
		xv = cosd(E1) - e
		yv = sqrt(1.0 - e * e) * sind(E1)

		v = atan2d(yv, xv)
		r = sqrt(xv * xv + yv * yv) 
		lonsun = v + w
		
		If (lonsun < 0 OR lonsun > 360) Then
			lonsun = rev(lonsun)
		End If
		
		xs = r * cosd(lonsun)
		ys = r * sind(lonsun)
		xe = xs
		ye = ys * cosd(ecl)
		ze = ys * sind(ecl)
		RA = atan2d(ye, xe)
		Dec = atan2d(ze, (sqrt((xe * xe) + (ye * ye))))
		h = -0.833

		GMST0 = L + 180
		
		If (GMST0 < 0 OR GMST0 > 360) Then
			GMST0 = rev(GMST0)
		End If

		UT_Sun_in_south = (RA - GMST0 - longitude) / 15.0
		
		If (UT_Sun_in_south < 0) Then
			UT_Sun_in_south = UT_Sun_in_south + 24
		End If

		LHA = (sind(h) - (sind(latitude) * sind(Dec))) / (cosd(latitude) * cosd(Dec))
		
		If (LHA > -1 AND LHA < 1) Then
			LHA	= acosd(LHA) / 15
		Else 
			Ret.Add "rise", "No sunrise"
			Ret.Add "set", "No sunset"
			Set SunTimes = Ret
			Exit Function
		End If 
		
		hour_rise = UT_Sun_in_south - LHA
		hour_set = UT_Sun_in_south + LHA
		min_rise = int((hour_rise-int(hour_rise)) * 60)
		min_set = int((hour_set-int(hour_set)) * 60)

		hour_rise = (int(hour_rise) + (TZ + isdst))
		
		hour_set = (int(hour_set) + (TZ + isdst))
		
		If (min_rise < 10) Then
			min_rise = right("0000" & min_rise, 2)
		End If
		
		If (min_set < 10) Then
			min_set = right("0000" & min_set, 2)
		End If
		
		Ret.Add "rise", hour_rise & ":" & min_rise
		Ret.Add "set", hour_set & ":" & min_set
		Set SunTimes = Ret
	End Function

	Private Function sind(qqq)
		sind = sin((qqq) * DEGRAD)
	End Function

	Private Function cosd(qqq)
		cosd = cos((qqq) * DEGRAD)
	End Function

	Private Function tand(qqq)
		tand = tan((qqq) * DEGRAD)
	End Function

	Private Function atand(qqq)
		atand = (RADEG * atan(qqq))
	End Function

	Private Function asind(qqq)
		asind = (RADEG * asin(qqq))
	End Function

	Private Function acosd(qqq) 
		acosd = (RADEG * acos(qqq))
	End Function

	Private Function atan2d (qqq, qqq1)
		atan2d = (RADEG * atan2(qqq, qqq1))
	End Function

	Private Function rev(qqq) 
		Dim x
		
		x = (qqq - int(qqq / 360.0) * 360.0)
		
		If (x <= 0) Then
			x = x + 360
		End If
		
		rev = x
	End Function

	Private Function atan2(ys,xs)
		Dim theta
		
		If xs <> 0 Then
		
			theta = Atn(ys / xs)
			
			If xs < 0 Then
				theta = theta + pi
			End If
		Else
			If ys < 0 Then
				theta = 3 * Pi / 2 '90
			Else
				theta = Pi / 2 '270
			End If
		End If
		
		atan2 = theta
	End Function

	Private Function acos(x)
		acos = Atn(-X / Sqrt(-X * X + 1)) + 2 * Atn(1)
	End Function

	Private Function sqrt(x)
		If x > 0 Then
			sqrt = Sqr(x)
		Else
			sqrt = 0
		End If
	End Function
End Class

Dim s 
Dim Today 
Dim Longitude 
Dim Latitude
Dim TimeZoneOffset

Longitude = 28.963565826416016
Latitude = 41.134128520610744

TimeZoneOffset = 2

Set s = New Sun 
Set Today = s.SunTimes(Date(), Longitude, Latitude, TimeZoneOffset, 0)
	
With Response
	.Write "<pre>"
	.Write "<img src=""http://static-maps.yandex.ru/1.x/lang=en-US&?ll=" & Longitude & "," & Latitude & "&pt=" & Longitude & "," & Latitude & "&spn=10,10&size=300,300&l=map"" />" & vbCrLf
	.Write "Date" & vbTab & vbTab & " : " & Date() & vbCrLf
	.Write "Longitude" & vbTab & " : " & Longitude & vbCrLf
	.Write "Latitude" & vbTab & " : " & Latitude & vbCrLf
	.Write "TimeZoneOffset" & vbTab & " : GMT+" & TimeZoneOffset & vbCrLf
	.Write "Sunrise" & vbTab & vbTab & " : " & Today("rise") & vbCrLf
	.Write "Sunset" & vbTab & vbTab & " : " & Today("set") & vbCrLf
	.Write "</pre>"
End With 

Set Today = Nothing
Set s = Nothing

%>
