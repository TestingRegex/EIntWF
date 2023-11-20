Option Explicit

Sub accessMacro()

   Dim appAccess As New Access.Application

   Set appAccess = Access.Application

   appAccess.OpenCurrentDatabase "C:\blah.mdb"

   appAccess.Visible = True

   appAccess.DoCmd.RunMacro "RunQueries.RunQueries"
   appAccess.CloseCurrentDatabase

End Sub