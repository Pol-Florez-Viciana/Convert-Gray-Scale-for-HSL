' Aquí Declaro una Estructura llamada RGB
Public Structure RGB
   Private _r As Byte
   Private _g As Byte
   Private _b As Byte
   Public Sub New(r As Byte, g As Byte, b As Byte)
     Me._r = r
     Me._g = g
     Me._b = b
   End Sub
   Public Property R() As Byte
     Get
     Return Me._r
     End Get
     Set(value As Byte)
     Me._r = value
     End Set
   End Property
   Public Property G() As Byte
     Get
     Return Me._g
     End Get
     Set(value As Byte)
     Me._g = value
     End Set
   End Property
   Public Property B() As Byte
     Get
     Return Me._b
     End Get
     Set(value As Byte)
     Me._b = value
     End Set
   End Property
   Public Overloads Function Equals(rgb As RGB) As Boolean
      Return (Me.R = rgb.R) AndAlso (Me.G = rgb.G) AndAlso (Me.B = rgb.B)
   End Function
End Structure

' Pol Flórez Viciana 16/04/2018
' Convertir Imagen a Escala de Grises con Visual Basic.NET

' Este procedimiento obtiene el Color en la Escala de Grises Correcta de la Imagen
Public Sub ConvertirEscalaDeGrises(ByVal Imagen As PictureBox)
  On Error Resume Next
  Dim imAgtEmp1 As New Bitmap(Imagen.Image)
  Dim imAgtEmp2 As New Bitmap(Imagen.Image)
  Dim elgrIs As Long
  Largo = imAgtEmp1.Width
  Alto = imAgtEmp1.Height
  For Me.i = 0 To Alto - 1
     For Me.j = 0 To Largo - 1
       Dim ColorActual As RGB
       ElColor = imAgtEmp1.GetPixel(Me.j, Me.i)
       ColorActual.R = CByte(ElColor.R.ToString)
       ColorActual.G = CByte(ElColor.G.ToString)
       ColorActual.B = CByte(ElColor.B.ToString)
       Dim Valor As Double
       If ColorActual.B > 128 Then
         If ColorActual.G < 128 Then
             If ColorActual.R < 128 Then
             ' Azules
             Valor = ((ColorActual.B - 127) * (ColorActual.G + 1) * (ColorActual.R + 1)) + 2097152
             Else
             ' Magentas
             Valor = ((ColorActual.B - 127) * (ColorActual.G + 1) * (ColorActual.R - 127)) + (2097152 * 2)
             End If
         Else
             If ColorActual.R < 128 Then
                ' Cyanes
                Valor = ((ColorActual.B - 127) * (ColorActual.G - 127) * (ColorActual.R + 1)) + (2097152 * 6)
             Else
                ' Blancos
                Valor = ((ColorActual.B - 127) * (ColorActual.G - 127) * (ColorActual.R - 127)) + (2097152 * 7)
             End If
         End If
       Else
         If ColorActual.G < 128 Then
             If ColorActual.R < 128 Then
                ' Negros
                Valor = (ColorActual.B + 1) * (ColorActual.G + 1) * (ColorActual.R + 1)
             Else
                ' Rojos
                Valor = ((ColorActual.B + 1) * (ColorActual.G + 1) * (ColorActual.R - 127)) + (2097152 * 3)
             End If
         Else
             If ColorActual.R < 128 Then
                ' Verdes
                Valor = ((ColorActual.B + 1) * (ColorActual.G - 127) * (ColorActual.R + 1)) + (2097152 * 5)
             Else
                ' Amarillos
                Valor = ((ColorActual.B + 1) * (ColorActual.G - 127) * (ColorActual.R - 127)) + (2097152 * 4)
             End If
         End If
       End If
       Valor = ((Valor / (256 ^ 3)) * 256) - 1
       ElColor = Color.FromArgb(CInt(Valor), CInt(Valor), CInt(Valor))
       imAgtEmp2.SetPixel(Me.j, Me.i, ElColor)
     Next
  Next
  Imagen.Image = imAgtEmp2.Clone
End Sub
