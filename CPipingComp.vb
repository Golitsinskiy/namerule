Option Explicit On
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports System.Collections.ObjectModel

Public Class CPipingComp
    Inherits NameRuleBase

    Public Overrides Sub ComputeName(oEntity As BusinessObject, oParents As ReadOnlyCollection(Of BusinessObject), oActiveEntity As BusinessObject)

        Dim strPipingCompName As String

        Dim propTrainNumber As PropertyValueString
        Dim strTrainNumber As String

        Dim propProcessSubUnit As PropertyValueCodelist
        Dim strProcessSubUnit As String

        Dim propElementType As PropertyValueCodelist
        Dim strElementType As String

        Dim propInstrumentSeriesNumber As PropertyValueString
        Dim strInstrumentSeriesNumber As String

        Dim propSuffixLetter As PropertyValueString
        Dim strSuffixLetter As String

        Try

            propTrainNumber = oEntity.GetPropertyValue("IJUAPipingCompNumbering", "TrainNumber")
            strTrainNumber = Trim(propTrainNumber.PropValue)

            propProcessSubUnit = oEntity.GetPropertyValue("IJUAPipingCompNumbering", "ProcessSubUnit")
            strProcessSubUnit = Trim(propProcessSubUnit.PropertyInfo.CodeListInfo.GetCodelistItem(propProcessSubUnit.PropValue).Name)

            propElementType = oEntity.GetPropertyValue("IJUAPipingCompNumbering", "ElementType")
            strElementType = Trim(propElementType.PropertyInfo.CodeListInfo.GetCodelistItem(propElementType.PropValue).Name)

            propInstrumentSeriesNumber = oEntity.GetPropertyValue("IJUAPipingCompNumbering", "InstrumentSeriesNumber")
            strInstrumentSeriesNumber = Trim(propInstrumentSeriesNumber.PropValue)

            propSuffixLetter = oEntity.GetPropertyValue("IJUAPipingCompNumbering", "SuffixLetter")
            strSuffixLetter = Trim(propSuffixLetter.PropValue)

            strPipingCompName = strTrainNumber & strProcessSubUnit & strElementType & strInstrumentSeriesNumber & strSuffixLetter

            oEntity.SetPropertyValue(strPipingCompName, "IJNamedItem", "Name")

        Catch ex As Exception

            Dim Log As New Logging()
            Log.WriteLog(ex.Message, "NameRule_PipingComp_Error.log")

        End Try


    End Sub

    Public Overrides Function GetNamingParents(oEntity As BusinessObject) As Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function
End Class
