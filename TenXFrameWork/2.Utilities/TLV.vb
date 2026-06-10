Imports System
Imports System.Text

'Namespace XMLandQRCodeGeneration
Public Class TLV
		Private tag As Integer

		Private value As String

		Public Sub New(ByVal tag As Integer, ByVal value As String)
			MyBase.New()
			If (If(value Is Nothing, True, value.Trim().Equals(""))) Then
				Throw New Exception("Value cannot be null or empty")
			End If
			Me.tag = tag
			Me.value = value
		End Sub

		Private Function getLength() As Integer
			Return Me.value.Length
		End Function

		Private Function getTag() As Integer
			Return Me.tag
		End Function

		Private Function getValue() As String
			Return Me.value
		End Function

		Private Function toHex(ByVal value As Integer) As String
		'	Dim hex As String = String.Format("{0:X2}", value)
		'	Dim input As String = If(hex.Length Mod 2 = 0, hex, String.Concat(hex, "0"))
		'	Dim output As StringBuilder = New StringBuilder()
		'	Dim i As Integer = 0
		'	While i < input.Length
		'		Dim str As String = input.Substring(i, i + 2)
		'		'	output.Append(CChar(Convert.ToInt32(str, 16)))
		'		output.Append(CStr(Convert.ToInt32(str, 16)))
		'		i += 2
		'	End While
		'Return output.ToString()

		Dim str As String = String.Format("{0:X2}", value)
		Dim str1 As String = If(str.Length Mod 2 = 0, str, String.Concat(str, "0"))
		Dim stringBuilder As System.Text.StringBuilder = New System.Text.StringBuilder()
		Dim num As Integer = 0
		While num < str1.Length
			Dim str2 As String = str1.Substring(num, num + 2)
			'Naresh
			'stringBuilder.Append(CChar(Convert.ToInt32(str2, 16)))
			stringBuilder.Append(CChar(str2))
			'----------
			num += 2
		End While
		Return stringBuilder.ToString()
	End Function

		Public Overrides Function ToString() As String
			Dim str As String = String.Concat(Me.toHex(Me.getTag()), Me.toHex(Me.getLength()), Me.getValue())
			Return str
		End Function
	End Class
'End Namespace