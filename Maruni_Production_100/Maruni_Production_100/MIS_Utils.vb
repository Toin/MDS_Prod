Option Strict Off
Option Explicit On

Public Class MIS_Utils
    Public Function fctFormatDate(ByVal pdate As Date, ByVal oCompany As SAPbobsCOM.Company, Optional ByVal sngFormat As Integer = 5) As String
        Dim strSeparator As String
        Dim oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        fctFormatDate = ""

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate
        strSeparator = oAdminInfo.DateSeparator

        Select Case sngFormat
            Case 0
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "yy")
            Case 1
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + "20" + Format(pdate, "yy")
            Case 2
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + Format(pdate, "yy")
            Case 3
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + "20" + Format(pdate, "yy")
            Case 4
                fctFormatDate = "20" + Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
            Case 5
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MMMM") + strSeparator + Format(pdate, "yyyy")
            Case 6
                fctFormatDate = Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
        End Select

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
    End Function

End Class
