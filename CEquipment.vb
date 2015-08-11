Option Explicit On
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports System.IO

Public Class CEquipment
    Inherits NameRuleBase

    Public Sub WriteLog(Message As String)

        Dim FileName As String = Path.GetTempPath() & "NameRule_Equipment_Error.log"
        Dim objStreamWriter As StreamWriter

        If File.Exists(FileName) = False Then
            objStreamWriter = File.CreateText(FileName)
            objStreamWriter.Flush()
            objStreamWriter.Close()
        End If

        objStreamWriter = File.AppendText(FileName)
        objStreamWriter.WriteLine(DateTime.Now.ToShortDateString & " " & DateTime.Now.ToLongTimeString & " " & Message)
        objStreamWriter.Close()

    End Sub


    Public Overrides Sub ComputeName(oEntity As BusinessObject, oParents As ObjectModel.ReadOnlyCollection(Of BusinessObject), oActiveEntity As BusinessObject)

        Dim strEquipName As String

        Dim propTrainNumber As PropertyValueString
        Dim strTrainNumber As String

        Dim propProcessSubUnit As PropertyValueCodelist
        Dim strProcessSubUnit As String

        Dim propEquipmentTypeIdentification As PropertyValueCodelist
        Dim strEquipmentTypeIdentification As String

        Dim propEquipmentSeriesNumber As PropertyValueString
        Dim strEquipmentSeriesNumber As String

        Dim propSuffixLetter As PropertyValueString
        Dim strSuffixLetter As String

        Try

            propTrainNumber = oEntity.GetPropertyValue("IJUAEquipNumbering", "TrainNumber")
            strTrainNumber = Trim(propTrainNumber.PropValue)

            propProcessSubUnit = oEntity.GetPropertyValue("IJUAEquipNumbering", "ProcessSubUnit")
            strProcessSubUnit = Trim(propProcessSubUnit.PropertyInfo.CodeListInfo.GetCodelistItem(propProcessSubUnit.PropValue).Name)

            propEquipmentTypeIdentification = oEntity.GetPropertyValue("IJUAEquipNumbering", "EquipmentTypeIdentification")
            strEquipmentTypeIdentification = Trim(propEquipmentTypeIdentification.PropertyInfo.CodeListInfo.GetCodelistItem(propEquipmentTypeIdentification.PropValue).Name)

            propEquipmentSeriesNumber = oEntity.GetPropertyValue("IJUAEquipNumbering", "EquipmentSeriesNumber")
            strEquipmentSeriesNumber = Trim(propEquipmentSeriesNumber.PropValue)

            propSuffixLetter = oEntity.GetPropertyValue("IJUAEquipNumbering", "SuffixLetter")
            strSuffixLetter = Trim(propSuffixLetter.PropValue)

            strEquipName = strTrainNumber & strProcessSubUnit & strEquipmentTypeIdentification & strEquipmentSeriesNumber & strSuffixLetter

            oEntity.SetPropertyValue(strEquipName, "IJNamedItem", "Name")

        Catch ex As Exception

            WriteLog(ex.Message)

        End Try

    End Sub

    Public Overrides Function GetNamingParents(oEntity As BusinessObject) As ObjectModel.Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function
End Class
