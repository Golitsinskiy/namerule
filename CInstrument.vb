Option Explicit On
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports System.Collections.ObjectModel

Public Class CInstrument
    Inherits NameRuleBase

    Public Overrides Sub ComputeName(oEntity As BusinessObject, oParents As ReadOnlyCollection(Of BusinessObject), oActiveEntity As BusinessObject)

        Dim strInstrumentName As String

        Dim propTrainNumber As PropertyValueString
        Dim strTrainNumber As String

        Dim propProcessSubUnit As PropertyValueCodelist
        Dim strProcessSubUnit As String

        Dim propFunctionalIdentification1 As PropertyValueCodelist
        Dim strFunctionalIdentification1 As String

        Dim propFunctionalIdentification2 As PropertyValueCodelist
        Dim strFunctionalIdentification2 As String

        Dim propFunctionalIdentification3 As PropertyValueCodelist
        Dim strFunctionalIdentification3 As String

        Dim propFunctionalIdentification4 As PropertyValueCodelist
        Dim strFunctionalIdentification4 As String

        Dim propInstrumentSeriesNumber As PropertyValueString
        Dim strInstrumentSeriesNumber As String

        Dim propSuffixLetter As PropertyValueString
        Dim strSuffixLetter As String

        Try

            propTrainNumber = oEntity.GetPropertyValue("IJUAInstrNumbering", "TrainNumber")
            strTrainNumber = Trim(propTrainNumber.PropValue)

            propProcessSubUnit = oEntity.GetPropertyValue("IJUAInstrNumbering", "ProcessSubUnit")
            strProcessSubUnit = Trim(propProcessSubUnit.PropertyInfo.CodeListInfo.GetCodelistItem(propProcessSubUnit.PropValue).Name)

            propFunctionalIdentification1 = oEntity.GetPropertyValue("IJUAInstrNumbering", "FunctionalIdentification1")
            strFunctionalIdentification1 = Trim(propFunctionalIdentification1.PropertyInfo.CodeListInfo.GetCodelistItem(propFunctionalIdentification1.PropValue).Name)

            propFunctionalIdentification2 = oEntity.GetPropertyValue("IJUAInstrNumbering", "FunctionalIdentification2")
            strFunctionalIdentification2 = Trim(propFunctionalIdentification2.PropertyInfo.CodeListInfo.GetCodelistItem(propFunctionalIdentification2.PropValue).Name)

            propFunctionalIdentification3 = oEntity.GetPropertyValue("IJUAInstrNumbering", "FunctionalIdentification3")
            strFunctionalIdentification3 = Trim(propFunctionalIdentification3.PropertyInfo.CodeListInfo.GetCodelistItem(propFunctionalIdentification3.PropValue).Name)

            propFunctionalIdentification4 = oEntity.GetPropertyValue("IJUAInstrNumbering", "FunctionalIdentification4")
            strFunctionalIdentification4 = Trim(propFunctionalIdentification4.PropertyInfo.CodeListInfo.GetCodelistItem(propFunctionalIdentification4.PropValue).Name)

            propInstrumentSeriesNumber = oEntity.GetPropertyValue("IJUAInstrNumbering", "InstrumentSeriesNumber")
            strInstrumentSeriesNumber = Trim(propInstrumentSeriesNumber.PropValue)

            propSuffixLetter = oEntity.GetPropertyValue("IJUAInstrNumbering", "SuffixLetter")
            strSuffixLetter = Trim(propSuffixLetter.PropValue)

            strInstrumentName = strTrainNumber & strProcessSubUnit & strFunctionalIdentification1 & strFunctionalIdentification2 & strFunctionalIdentification3 & strFunctionalIdentification4 & strInstrumentSeriesNumber & strSuffixLetter

            oEntity.SetPropertyValue(strInstrumentName, "IJNamedItem", "Name")

        Catch ex As Exception

            Dim Log As New Logging()
            Log.WriteLog(ex.Message, "NameRule_Instrument_Error.log")

        End Try



    End Sub

    Public Overrides Function GetNamingParents(oEntity As BusinessObject) As Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function
End Class
