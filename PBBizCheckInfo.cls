VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBBizCheckInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public corpNum As String
Public checkDT As String
Public corpName As String
Public corpCode As Integer
Public corpScaleCode As Integer
Public personCorpCode As Integer
Public headOfficeCode As Integer
Public industryCode As String
Public companyRegNum As String
Public establishDate As String
Public establishCode As Integer
Public ceoname As String
Public workPlaceCode As Integer
Public addrCode As Integer
Public zipCode As String
Public addr As String
Public addrDetail As String
Public enAddr As String
Public bizClass As String
Public bizType As String
Public result As Integer
Public resultMessage As String
Public closeDownTaxType As Integer
Public closeDownTaxTypeDate As String
Public closeDownState As Integer
Public closeDownStateDate As String

Public Sub fromDictionary(dic As Variant)
    If IsNull(dic.Item("corpNum")) = False Then
        corpNum = dic.Item("corpNum")
    Else
        corpNum = ""
    End If
    If IsNull(dic.Item("checkDT")) = False Then
            checkDT = dic.Item("checkDT")
        Else
            checkDT = ""
        End If
    If IsNull(dic.Item("corpName")) = False Then
            corpName = dic.Item("corpName")
        Else
            corpName = ""
        End If
    If IsNull(dic.Item("corpCode")) = False Then
            corpCode = dic.Item("corpCode")
        End If
    If IsNull(dic.Item("corpScaleCode")) = False Then
            corpScaleCode = dic.Item("corpScaleCode")
        End If
    If IsNull(dic.Item("personCorpCode")) = False Then
            personCorpCode = dic.Item("personCorpCode")
        End If
    If IsNull(dic.Item("headOfficeCode")) = False Then
            headOfficeCode = dic.Item("headOfficeCode")
        End If
    If IsNull(dic.Item("industryCode")) = False Then
            industryCode = dic.Item("industryCode")
        Else
            industryCode = ""
        End If
    If IsNull(dic.Item("companyRegNum")) = False Then
            companyRegNum = dic.Item("companyRegNum")
        Else
            companyRegNum = ""
        End If
    If IsNull(dic.Item("establishDate")) = False Then
            establishDate = dic.Item("establishDate")
        Else
            establishDate = ""
        End If
    If IsNull(dic.Item("establishCode")) = False Then
            establishCode = dic.Item("establishCode")
        End If
    If IsNull(dic.Item("ceoname")) = False Then
            ceoname = dic.Item("ceoname")
        Else
            ceoname = ""
        End If
    If IsNull(dic.Item("workPlaceCode")) = False Then
            workPlaceCode = dic.Item("workPlaceCode")
        End If
    If IsNull(dic.Item("addrCode")) = False Then
            addrCode = dic.Item("addrCode")
        End If
    If IsNull(dic.Item("zipCode")) = False Then
            zipCode = dic.Item("zipCode")
        Else
            zipCode = ""
        End If
    If IsNull(dic.Item("addr")) = False Then
            addr = dic.Item("addr")
        Else
            addr = ""
        End If
    If IsNull(dic.Item("addrDetail")) = False Then
            addrDetail = dic.Item("addrDetail")
        Else
            addrDetail = ""
        End If
    If IsNull(dic.Item("enAddr")) = False Then
            enAddr = dic.Item("enAddr")
        Else
            enAddr = ""
        End If
    If IsNull(dic.Item("bizClass")) = False Then
            bizClass = dic.Item("bizClass")
        Else
            bizClass = ""
        End If
    If IsNull(dic.Item("bizType")) = False Then
            bizType = dic.Item("bizType")
        Else
            bizType = ""
        End If
    If IsNull(dic.Item("result")) = False Then
            result = dic.Item("result")
        End If
    If IsNull(dic.Item("resultMessage")) = False Then
            resultMessage = dic.Item("resultMessage")
        Else
            resultMessage = ""
        End If
    If IsNull(dic.Item("closeDownTaxType")) = False Then
            closeDownTaxType = dic.Item("closeDownTaxType")
        End If
    If IsNull(dic.Item("closeDownTaxTypeDate")) = False Then
            closeDownTaxTypeDate = dic.Item("closeDownTaxTypeDate")
        Else
            closeDownTaxTypeDate = ""
        End If
    If IsNull(dic.Item("closeDownState")) = False Then
            closeDownState = dic.Item("closeDownState")
        End If
    If IsNull(dic.Item("closeDownStateDate")) = False Then
            closeDownStateDate = dic.Item("closeDownStateDate")
        Else
            closeDownStateDate = ""
        End If
End Sub
