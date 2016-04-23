#tag Class
Protected Class RARException
Inherits RuntimeException
	#tag Method, Flags = &h1000
		Sub Constructor(RARErrorNumber As Integer)
		  Me.ErrorNumber = RARErrorNumber
		  Me.Message = UnRAR.FormatError(Me.ErrorNumber)
		End Sub
	#tag EndMethod


End Class
#tag EndClass
