Option Explicit On
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports System.Collections.ObjectModel

Public Class CPipeRun
    Inherits NameRuleBase

    Public Overrides Sub ComputeName(oEntity As BusinessObject, oParents As ReadOnlyCollection(Of BusinessObject), oActiveEntity As BusinessObject)

        Dim strPipeRunName As String

        Dim propTrainNumber As PropertyValueString
        Dim strTrainNumber As String

        Dim propProcessSubUnit As PropertyValueCodelist
        Dim strProcessSubUnit As String

        Dim propFluidCode As PropertyValueCodelist
        Dim strFluidCode As String

        Dim propLineNumber As PropertyValueString
        Dim strLineNumber As String

        Dim propNpd As PropertyValueDouble
        Dim strNpd As String

        Dim propNpdUnitType As PropertyValueString
        Dim strNpdUnitType As String

        Dim propPipingClass As PropertyValueString
        Dim strPipingClass As String

        Dim propInsulationPurpose As PropertyValueCodelist
        Dim strInsulationPurpose As String

        Dim propHeatTracingType As PropertyValueCodelist
        Dim strHeatTracingType As String

        Try

            propTrainNumber = oEntity.GetPropertyValue("IJUAPipeNumbering", "TrainNumber")
            strTrainNumber = Trim(propTrainNumber.PropValue)

            propProcessSubUnit = oEntity.GetPropertyValue("IJUAPipeNumbering", "ProcessSubUnit")
            strProcessSubUnit = Trim(propProcessSubUnit.PropertyInfo.CodeListInfo.GetCodelistItem(propProcessSubUnit.PropValue).Name)

            propFluidCode = oEntity.GetPropertyValue("IJUAPipeNumbering", "FluidCode")
            strFluidCode = Trim(propFluidCode.PropertyInfo.CodeListInfo.GetCodelistItem(propFluidCode.PropValue).Name)

            propLineNumber = oEntity.GetPropertyValue("IJUAPipeNumbering", "LineNumber")
            strLineNumber = Trim(propLineNumber.PropValue)

            propNpd = oEntity.GetPropertyValue("IJRtePipeRun", "NPD")
            strNpd = Trim(propNpd.PropValue)

            propNpdUnitType = oEntity.GetPropertyValue("IJRtePipeRun", "NPDUnitType")
            strNpdUnitType = Trim(propNpdUnitType.PropValue)

            'If strNpdUnitType = "mm" Then
            'strNpd = Math.Round(CInt(strNpd) / 25.4, 1)
            'End If

            'strNpd = strNpd & Chr(34)

            'If strNpdUnitType = "in" Then
            'strNpdUnitType = Chr(34) & Chr(34)
            'End If

            propPipingClass = oEntity.GetPropertyValue("IJUAPipeNumbering", "PipingClass")
            strPipingClass = Trim(propPipingClass.PropValue)

            propInsulationPurpose = oEntity.GetPropertyValue("IJRteInsulation", "Purpose")
            strInsulationPurpose = Trim(propInsulationPurpose.PropertyInfo.CodeListInfo.GetCodelistItem(propInsulationPurpose.PropValue).Name)

            propHeatTracingType = oEntity.GetPropertyValue("IJHeatTracingInfo", "HeatTracingType")
            strHeatTracingType = Trim(propHeatTracingType.PropertyInfo.CodeListInfo.GetCodelistItem(propHeatTracingType.PropValue).Name)

            strPipeRunName = strTrainNumber & strProcessSubUnit & strFluidCode & "-" & strLineNumber & "-" & strNpd & strNpdUnitType & "-" & strPipingClass & "-" & strInsulationPurpose & strHeatTracingType

            oEntity.SetPropertyValue(strPipeRunName, "IJNamedItem", "Name")

        Catch ex As Exception

            Dim Log As New Logging()
            Log.WriteLog(ex.Message, "NameRule_PipeRun_Error.log")

        End Try

    End Sub

    Public Overrides Function GetNamingParents(oEntity As BusinessObject) As Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function
End Class
