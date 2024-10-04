Attribute VB_Name = "ModStructure"
Public Type SetModel

    ModelName As String * 30
    Exposure(0 To 3) As Double

'Favis Edge Tool 정보 저장
    favEdgeTThresh(0 To 3, 0 To 29) As Double
    favEdgeTPorarity(0 To 3, 0 To 29) As Integer
    favEdgeTMode(0 To 3, 0 To 29) As Integer
    favEdgeTSideX(0 To 3, 0 To 29) As Long
    favEdgeTSideY(0 To 3, 0 To 29) As Long
    favEdgeTCenterX(0 To 3, 0 To 29) As Double
    favEdgeTCenterY(0 To 3, 0 To 29) As Double
    favEdgeTRotation(0 To 3, 0 To 29) As Double

    favFixEdgeTThresh(0 To 3, 0 To 3) As Double
    favFixEdgeTPorarity(0 To 3, 0 To 3) As Integer
    favFixEdgeTMode(0 To 3, 0 To 3) As Integer
    favFixEdgeTSideX(0 To 3, 0 To 3) As Long
    favFixEdgeTSideY(0 To 3, 0 To 3) As Long
    favFixEdgeTCenterX(0 To 3, 0 To 3) As Double
    favFixEdgeTCenterY(0 To 3, 0 To 3) As Double
    favFixEdgeTRotation(0 To 3, 0 To 3) As Double

    favCalEdgeTThresh(0 To 3, 0 To 3) As Double
    favCalEdgeTPorarity(0 To 3, 0 To 3) As Integer
    favCalEdgeTMode(0 To 3, 0 To 3) As Integer
    favCalEdgeTSideX(0 To 3, 0 To 3) As Long
    favCalEdgeTSideY(0 To 3, 0 To 3) As Long
    favCalEdgeTCenterX(0 To 3, 0 To 3) As Double
    favCalEdgeTCenterY(0 To 3, 0 To 3) As Double
    favCalEdgeTRotation(0 To 3, 0 To 3) As Double

'Favis Blob Tool 정보 저장
    favBlobTThreshold(0 To 3, 0 To 29) As Double
    favBlobTMin(0 To 3, 0 To 29) As Double
    favBlobTPolarity(0 To 3, 0 To 29) As Double
    favBlobTWidth(0 To 3, 0 To 29) As Long
    favBlobTHeight(0 To 3, 0 To 29) As Long
    favBlobTCenterX(0 To 3, 0 To 29) As Long
    favBlobTCenterY(0 To 3, 0 To 29) As Long
    favBlobTAngel(0 To 3, 0 To 29) As Long
End Type

Public Modelinfo As SetModel


'Public Type SetModel
'
'    ModelName As String * 30
'    Exposure(0 To 1) As Double
'
''Favis Edge Tool 정보 저장
'    favEdgeTThresh(0 To 1, 0 To 29) As Double
'    favEdgeTPorarity(0 To 1, 0 To 29) As Integer
'    favEdgeTMode(0 To 1, 0 To 29) As Integer
'    favEdgeTSideX(0 To 1, 0 To 29) As Long
'    favEdgeTSideY(0 To 1, 0 To 29) As Long
'    favEdgeTCenterX(0 To 1, 0 To 29) As Double
'    favEdgeTCenterY(0 To 1, 0 To 29) As Double
'    favEdgeTRotation(0 To 1, 0 To 29) As Double
'
'    favFixEdgeTThresh(0 To 1, 0 To 3) As Double
'    favFixEdgeTPorarity(0 To 1, 0 To 3) As Integer
'    favFixEdgeTMode(0 To 1, 0 To 3) As Integer
'    favFixEdgeTSideX(0 To 1, 0 To 3) As Long
'    favFixEdgeTSideY(0 To 1, 0 To 3) As Long
'    favFixEdgeTCenterX(0 To 1, 0 To 3) As Double
'    favFixEdgeTCenterY(0 To 1, 0 To 3) As Double
'    favFixEdgeTRotation(0 To 1, 0 To 3) As Double
'
'    favCalEdgeTThresh(0 To 1, 0 To 3) As Double
'    favCalEdgeTPorarity(0 To 1, 0 To 3) As Integer
'    favCalEdgeTMode(0 To 1, 0 To 3) As Integer
'    favCalEdgeTSideX(0 To 1, 0 To 3) As Long
'    favCalEdgeTSideY(0 To 1, 0 To 3) As Long
'    favCalEdgeTCenterX(0 To 1, 0 To 3) As Double
'    favCalEdgeTCenterY(0 To 1, 0 To 3) As Double
'    favCalEdgeTRotation(0 To 1, 0 To 3) As Double
'
''Favis Blob Tool 정보 저장
'    favBlobTThreshold(0 To 1, 0 To 29) As Double
'    favBlobTMin(0 To 1, 0 To 29) As Double
'    favBlobTPolarity(0 To 1, 0 To 29) As Double
'    favBlobTWidth(0 To 1, 0 To 29) As Long
'    favBlobTHeight(0 To 1, 0 To 29) As Long
'    favBlobTCenterX(0 To 1, 0 To 29) As Long
'    favBlobTCenterY(0 To 1, 0 To 29) As Long
'    favBlobTAngel(0 To 1, 0 To 29) As Long
'End Type
'
'Public Modelinfo As SetModel
