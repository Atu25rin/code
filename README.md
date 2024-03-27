Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports System.IO
Imports NXOpen.Utilities
Imports System.DateTime
Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Diagnostics
Imports System.Threading
Imports System.Data.OleDb
Imports NXOpen.Features
Imports System.Drawing
'Imports System.Windows.Forms
'Imports Microsoft.Office
'Imports Microsoft.Office.Interop.PowerPoint
'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
'Imports System.Windows.Forms
Imports System.Drawing.Imaging
Imports System.Collections
Imports NXOpen.GeometricUtilities

Imports NXOpen.UI
Imports K_Lib_Share

Public Class Zoukei_2023_DB0_069_PATAN1
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim theUFSession As UFSession = UFSession.GetUFSession()
    Dim displayPart As NXOpen.Part = theSession.Parts.Display
    Dim lw As ListingWindow = theSession.ListingWindow
    Dim Lib_Office As New Lib_Office
    Dim Lib_NX As New Lib_NX
    'Top Assy
    Dim iRootComp As Component = workPart.ComponentAssembly.RootComponent
    Dim iComTop As NXOpen.Assemblies.Component = Nothing
    Dim partLoadStatus1 As NXOpen.PartLoadStatus = Nothing
    Const ExcelLink As String = "D:\nxcad\scratch\VehiclePlanMacro\2023_DB_J00001\Character_line_CHECK_Input.xlsm"
    Const PPOut As String = "D:\nxcad\scratch\VehiclePlanMacro\2023_DB_J00001\Character_line_CHECK_Output.pptx"
    'Data Bango: DATA CL_TOOL
    'entire part, show exact

#Region "1_Varial"
    Dim Zoukei_Comp As NXOpen.Assemblies.Component
    Dim Patan1_Glass() As NXOpen.Body
    Dim Patan1_Handle() As NXOpen.Body
    Dim Patan1_Handle_buhin() As NXOpen.Body
    Dim CL_assy As Component
    Dim glassmen, saigaimen, Saigaimen_Draft As NXOpen.Body
    Dim glass1_CL, glass1_Draft As NXOpen.Body
    Dim Str_GRD, End_GRD As Point3d
    Dim extrude_space As Double = 2000
    Dim Glass_Offset_KCach As Double
    Dim Doday As Double
    Dim ZoukeiCL As Double
    Dim draft_pnt() As Point3d
    Dim saigai_sen, glass_sen, glassCL_sen As NXOpen.Curve
    Dim Ketluan As String
    Dim Handle_Inside(), Handle_Outside(), CW(), REINF As Body
    Dim PMI_size As Double = 15
    Const ImgeFolder1 As String = "D:\nxcad\scratch\Tool_Picture\"
    Dim saigaiY As Double = 990
#End Region

#Region "2_Varial"
    Dim PartSA_comp As Component
    Dim Csys_in As DatumCsys
#End Region

#Region "3_Varial"
    Dim Fr_SE_Pnt As NXOpen.Point
    Dim Fr_Center_pt As NXOpen.Point
    Dim Start_Grd_pnt As NXOpen.Point3d '= New Point3d(-65.4, 0, -135)
    Dim End_Grd_pnt As NXOpen.Point3d '= New Point3d(2554.6, 0, -121)
    Dim Grd_line As NXOpen.Line
    Dim Kcach, draft As Double
    Dim Center_Pnt_out As NXOpen.Point
    Dim line2 As NXOpen.Curve
    Dim help1 As NXOpen.Point3d
    Dim help2 As NXOpen.Point3d
    Dim iHead_circle As NXOpen.Arc
    Dim texthigh As Double = 50
    Dim Ground_line As NXOpen.Line
    Dim Condition_IN() As Double
    Const ImgeFolder3 As String = "D:\nxcad\scratch\Tool_Picture2\"


#End Region


    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Region "1_Clearance"

    Sub Main1(ByVal iZoukei As NXOpen.Assemblies.Component, ByVal iGlass As Body, ByVal Handle_cnt As NXOpen.Point, ByVal Input_Datum As NXOpen.DatumPlane(), ByVal In_Handle() As Body, ByVal Out_Handle() As Body, ByVal CW_in() As Body, ByVal REINF_in As Body, ByVal iinput() As Double)
        lw.Open()
        Call iResetFolder(ImgeFolder1)
        Call dodayduong(True)
        Glass_Offset_KCach = iinput(0)
        ZoukeiCL = iinput(1) + iinput(2)
        Doday = iinput(2)
        Try

            Zoukei_Comp = iZoukei
            For i As Integer = 0 To UBound(In_Handle)
                ReDim Preserve Handle_Inside(i)
                Handle_Inside(i) = In_Handle(i)
            Next

            For i As Integer = 0 To UBound(Out_Handle)
                ReDim Preserve Handle_Outside(i)
                Handle_Outside(i) = Out_Handle(i)
            Next

            For i As Integer = 0 To UBound(CW_in)
                ReDim Preserve CW(i)
                CW(i) = CW_in(i)
            Next

            REINF = CType(REINF_in, Body)


            theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)

            Call Lib_NX.Create_Assy("CL_Check", CL_assy, True)
            iGlass.Unblank()
            Call Patan1(iGlass, Handle_cnt, Input_Datum, CL_assy)
            Call dodayduong(False)
            Call View("Left")
            Try
                workPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.ShadedWithEdges
            Catch ex As Exception
                displayPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.ShadedWithEdges
            End Try
        Catch ex As Exception
            lw.WriteFullline(ex.ToString)
        End Try
        theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)
    End Sub
    Sub Patan1(ByVal Glass_zoukei As NXOpen.Body, ByVal Handle_cnt As NXOpen.Point, ByVal Datum_In As NXOpen.DatumPlane(), ByVal Assy_check As Component)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim Kcach1 As Double
        Dim bodytodel As NXOpen.Body
        'Wawe data vao assy
        Call Wave_Glass_saigaishoku(Glass_zoukei, glassmen, saigaimen)

        Dim men1, men2 As NXOpen.Body
        Call offset_3d(glassmen, Glass_Offset_KCach, True, men1)
        Call offset_3d(glassmen, Glass_Offset_KCach, False, men2)

        Call Get_MinDist(men1.Tag, saigaimen.Tag, Kcach, draft_pnt)
        Call Get_MinDist(men2.Tag, saigaimen.Tag, Kcach1, draft_pnt)
        If Kcach1 > Kcach Then
            Call Del_body(men2)
        Else
            Call Del_body(men1)
        End If


        Dim SA_danmen, SB_danmen, SC_danmen, SD_danmen As NXOpen.Assemblies.Component

        Call taomatcat(Datum_In, Assy_check, SA_danmen, SB_danmen, SC_danmen, SD_danmen)

        Try
            theUFSession.Obj.DeleteObject(men2.Tag)
            theUFSession.Obj.DeleteObject(men1.Tag)
        Catch ex As Exception

        End Try
        'tao them cac mat
        Call Lib_NX.SetWorkPart(Assy_check, workPart)

        Call Offset_Glass_And_saigaishoku(glass1_CL, glass1_Draft, Handle_cnt)


        'CHECK
        'Mat SA
        Call Danmen_Check(Assy_check, "SA", SA_danmen, "Z", Handle_cnt, Datum_In(0))
        'Mat SB
        Call Danmen_Check(Assy_check, "SB", SB_danmen, "X", Handle_cnt, Datum_In(1))
        'Mat SC
        Call Danmen_Check(Assy_check, "SC", SC_danmen, "X", Handle_cnt, Datum_In(2))
        ''Mat SD
        Call Danmen_Check(Assy_check, "SD", SD_danmen, "X", Handle_cnt, Datum_In(3))


        'Call Lib_NX.SetWorkPart(Assy_check, workPart)
        theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)
        Try
            theUFSession.Obj.DeleteObject(glass1_Draft.Tag)
            theUFSession.Obj.DeleteObject(Saigaimen_Draft.Tag)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try

        'Call Unblank(Datum_In(0).OwningComponent)
        'Glass_zoukei.Unblank()
    End Sub
    Sub Wave_Glass_saigaishoku(ByVal iGlass_in As NXOpen.Body, ByRef iGlass_out As NXOpen.Body, ByRef saigaishokumen As NXOpen.Body)
        Call Wave_Body(iGlass_in, iGlass_out)
        iGlass_in.Blank()
        Str_GRD = New Point3d(-1000, -1 * saigaiY, 0)
        End_GRD = New Point3d(4000, -1 * saigaiY, 0)
        'saigai_length_W = System.Math.Abs(Str_GRD.X) + System.Math.Abs(End_GRD.X)
        Dim Line_saigai As NXOpen.Line
        Line_saigai = workPart.Curves.CreateLine(Str_GRD, End_GRD)
        Call extrude(Line_saigai, CL_assy, saigaishokumen)
        Line_saigai.SetVisibility(SmartObject.VisibilityOption.Invisible)
    End Sub
    Sub taomatcat(ByVal Datum_ As NXOpen.DatumPlane(), ByVal Assy_cmp As Component, ByRef CL_DatumSA As Component, ByRef CL_DatumSB As Component, ByRef CL_DatumSC As Component, ByRef CL_DatumSD As Component)
        ''Mat SA
        Call Section_order("SA", CL_DatumSA, Datum_(0), Assy_cmp)
        ''Mat SB
        Call Section_order("SB", CL_DatumSB, Datum_(1), Assy_cmp)

        ''Mat SC
        Call Section_order("SC", CL_DatumSC, Datum_(2), Assy_cmp)
        ''Mat SD
        Call Section_order("SD", CL_DatumSD, Datum_(3), Assy_cmp)
    End Sub
    Sub Section_order(ByVal Ten_Comp As String, ByRef Comp As NXOpen.Assemblies.Component, ByVal Datum_input As DatumPlane, ByVal A_Comp As NXOpen.Assemblies.Component)

        Call Lib_NX.Create_Component(Ten_Comp, Comp, True)

        'Section 1: Saigaishoku & Glass, Glass CL
        Call Unblank(A_Comp)
        Call CreateSect(Datum_input)
        '------------------------------------------------
        'Section 2: Zoikei
        'An  tat ca
        For Each iChild As Component In iRootComp.GetChildren
            Call Blank(iChild)
        Next
        'Hien Zoukei
        Zoukei_Comp.Unblank()
        'Tao matcat
        Call CreateSect(Datum_input)
        Zoukei_Comp.Blank()
        '------------------------------------------------
        'Section 3: Handle Inside
        'Hien Handle Inside
        For i As Integer = 0 To UBound(Handle_Inside)
            Call Unblank(Handle_Inside(i).OwningComponent)
        Next
        'Tao matcat
        Call CreateSect(Datum_input)
        For i As Integer = 0 To UBound(Handle_Inside)
            Call Blank(Handle_Inside(i).OwningComponent)
        Next
        '------------------------------------------------
        'Section 4: Handle Outside
        'Hien Handle Outside
        For i As Integer = 0 To UBound(Handle_Outside)
            Call Unblank(Handle_Outside(i).OwningComponent)
        Next
        'Tao matcat
        Call CreateSect(Datum_input)
        For i As Integer = 0 To UBound(Handle_Outside)
            Call Blank(Handle_Outside(i).OwningComponent)
        Next
        '------------------------------------------------
        'Section 5: REINF & CW
        'Hien REINF & CW
        Call Unblank(REINF.OwningComponent)
        For i As Integer = 0 To UBound(CW)
            Call Unblank(CW(i).OwningComponent)
        Next
        'Tao matcat
        Call CreateSect(Datum_input)
        'An REINF & CW
        Call Blank(REINF.OwningComponent)
        For i As Integer = 0 To UBound(CW)
            Call Blank(CW(i).OwningComponent)
        Next
        '------------------------------------------------
        Call Unblank(A_Comp)
        Call Lib_NX.SetWorkPart(A_Comp, workPart)

        'Patan1_Zoukei_Comp.Blank()
    End Sub
    Sub Offset_Glass_And_saigaishoku(ByRef Glass_CL As NXOpen.Body, ByRef Glass_Draft As NXOpen.Body, ByVal center_Pnt As NXOpen.Point)
        Dim iBody1_Offset, iBody2_Offset As NXOpen.Body
        Call offset_3d(glassmen, Glass_Offset_KCach, False, iBody1_Offset)
        Call offset_3d(glassmen, Glass_Offset_KCach, True, iBody2_Offset)



        '--->>>tim mindist voi saigaishoku de phan biet 2 duowng
        Dim iip_1() As NXOpen.Point3d
        Dim iip_2() As NXOpen.Point3d
        Dim Sosanh As Double
        Call Get_MinDist(saigaimen.Tag, iBody1_Offset.Tag, Kcach, iip_1)
        Sosanh = Kcach
        Call Get_MinDist(saigaimen.Tag, iBody2_Offset.Tag, Kcach, iip_2)
        If Sosanh < Kcach Then
            Glass_CL = iBody1_Offset
            Glass_Draft = iBody2_Offset
        Else
            Glass_CL = iBody2_Offset
            Glass_Draft = iBody1_Offset
        End If

        '----------------------------------------------------------------
        Dim draft_saigai As NXOpen.Body
        Dim iSaigai1_Offset, iSaigai2_Offset As NXOpen.Body
        Call offset_3d(saigaimen, Glass_Offset_KCach, False, iSaigai1_Offset)
        Call offset_3d(saigaimen, Glass_Offset_KCach, True, iSaigai2_Offset)
        Call Get_MinDist(glassmen.Tag, iSaigai1_Offset.Tag, Kcach, iip_1)
        Sosanh = Kcach
        Call Get_MinDist(glassmen.Tag, iSaigai2_Offset.Tag, Kcach, iip_1)

        Dim bodytodel As NXOpen.Body
        If Sosanh < Kcach Then
            Saigaimen_Draft = iSaigai2_Offset
            bodytodel = iSaigai1_Offset
        Else
            Saigaimen_Draft = iSaigai1_Offset
            bodytodel = iSaigai2_Offset
        End If

        Try
            theUFSession.Obj.DeleteObject(bodytodel.Tag)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try
        'An het body
        Dim numberHidden1 As Integer = Nothing
        numberHidden1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_BODIES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)

        displayPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
    End Sub
    Sub Danmen_Check(ByVal Assy As Component, Section_Patan As String, ByVal iprt_Comp As Component, ByVal PMI_Plane As String, ByVal Handle_cnt As NXOpen.Point, ByVal Daplane As DatumPlane)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        For Each icomp As Component In Assy.GetChildren
            Call Blank(icomp)
        Next
        Call Unblank(iprt_Comp)
        Call Lib_NX.SetWorkPart(iprt_Comp, workPart)
        'Section 1------------------------------------------------------------
        'tim saigaisen & Glass sen & 'tim diemgan nhat voi saigaisen cua OH & diemgan nhat voi Glass sen cua OH
        Dim diemcantim1, diemcantim2, diemcantim3, diemcantim4, diemcantim5, diemcantim6, diemcantim7, diemcantim8, diemcantim9, diemcantim10, diemcantim11, diemcantim12, diemcantim13 As Point3d
        Dim tmpGrp As NXOpen.Tag = NXOpen.Tag.Null
        Dim myGroup As NXOpen.Group
        Dim member_count As Integer
        Dim group_member_list() As Tag
        Dim group_member_list1() As Tag
        Dim kcmin1 As Double = 30000
        Dim kcmin2 As Double = 30000
        Dim kcmin3 As Double = 30000
        Dim kcmin4 As Double = 30000
        Dim kcmin5 As Double = 30000
        Dim kcmin6 As Double = 30000
        Dim kcmin7 As Double = 30000
        Dim kcmin8 As Double = 30000
        Dim kcmin9 As Double = 30000
        Dim kcmin10 As Double = 30000
        Dim kcmin11 As Double = 30000
        Dim kcmin12 As Double = 30000
        Dim Do_REINF As Boolean = False

        'Draft line de tim diem thap nhat cua REINF OTR
        Dim drf_1 As Point3d = New Point3d(Handle_cnt.Coordinates.X, Handle_cnt.Coordinates.Y - 999, Handle_cnt.Coordinates.Z + 999)
        Dim drf_2 As Point3d = New Point3d(Handle_cnt.Coordinates.X, Handle_cnt.Coordinates.Y + 999, Handle_cnt.Coordinates.Z + 999)
        Dim Line_Draft_Upper As Line
        Line_Draft_Upper = workPart.Curves.CreateLine(drf_1, drf_2)
        Line_Draft_Upper.SetVisibility(SmartObject.VisibilityOption.Invisible)

        'Draft line de tim diem cao nhat cua CW
        Dim drf_3 As Point3d = New Point3d(Handle_cnt.Coordinates.X, Handle_cnt.Coordinates.Y - 999, Handle_cnt.Coordinates.Z - 999)
        Dim drf_4 As Point3d = New Point3d(Handle_cnt.Coordinates.X, Handle_cnt.Coordinates.Y + 999, Handle_cnt.Coordinates.Z - 999)
        Dim Line_Draft_lower As Line
        Line_Draft_lower = workPart.Curves.CreateLine(drf_3, drf_4)
        Line_Draft_lower.SetVisibility(SmartObject.VisibilityOption.Invisible)

        Dim saigai_length As Double
        Dim Zoukei_curve_Arr(1) As Curve, draft_curve As Curve, zoukei_min As Double = 9999
        Dim idem_Zoukei_curve As Integer = 0
        Dim BRKT_curve(), CW_curve(), REINF_curve() As Curve
        Dim idem_BRKT, idem_CW, idem_REINF As Integer

        'Do duong
        Do
            theUFSession.Obj.CycleObjsInPart(workPart.Tag, UFConstants.UF_group_type, tmpGrp)
            If tmpGrp = NXOpen.Tag.Null Then
                Continue Do
            End If
            Dim theType As Integer, theSubtype As Integer
            theUFSession.Obj.AskTypeAndSubtype(tmpGrp, theType, theSubtype)

            If theSubtype = UFConstants.UF_group_type Then
                myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            End If
            myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            Dim guess1(2) As Double
            Dim guess2(2) As Double
            Dim pt1(2) As Double
            Dim pt2(2) As Double
            Dim junk(2) As Double, trash As Double
            If InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) > 0 Then
                Continue Do
            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 1".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                'Section 1: Saigaishoku & Glass, Glass CL
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            Dim StartPoint(2) As Double
                            Dim midPt(2) As Double
                            Dim EndPoint(2) As Double

                            '------------------------------------------------------------------------------------------------------------------------
                            'tim saigaisen
                            Call Get_MinDist(iCurve.Tag, Saigaimen_Draft.Tag, Kcach, draft_pnt)
                            If System.Math.Abs(Kcach - Glass_Offset_KCach) < 0.001 Then
                                saigai_sen = CType(iCurve, Curve)
                            End If
                            '------------------------------------------------------------------------------------------------------------------------
                            'Tim Glass sen
                            Call Get_MinDist(iCurve.Tag, glass1_Draft.Tag, Kcach, draft_pnt)
                            If System.Math.Abs(Kcach - Glass_Offset_KCach) < 0.001 Then
                                glass_sen = CType(iCurve, Curve)
                            End If
                            '------------------------------------------------------------------------------------------------------------------------
                            'Tim Glass_CL sen
                            Call Get_MinDist(iCurve.Tag, glass1_Draft.Tag, Kcach, draft_pnt)
                            If System.Math.Abs(Kcach - 2 * Glass_Offset_KCach) < 0.001 Then
                                glassCL_sen = CType(iCurve, Curve)
                            End If
                        End If
                    Next
                Next

            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 2".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then
                'Section 2: Zoikei
                If PMI_Plane = "Z" Then 'ko lam gi neu la mat SA
                Else
                    theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                    For Each iGrp As Tag In group_member_list
                        theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                        For Each iTag As Tag In group_member_list1
                            Dim iObj As NXObject = Nothing
                            iObj = NXObjectManager.Get(iTag)

                            If iObj.GetType.ToString <> "NXOpen.Group" Then
                                Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                                'tim zoukei: Cac duong co diem dau, diem cuoi trong khoang -+150 so voi Handle center
                                Dim d1 As Point3d = timdiemdau("Z", True, iCurve)
                                Dim d2 As Point3d = timdiemdau("Z", False, iCurve)
                                If d1.Z > Handle_cnt.Coordinates.Z - 150 And d2.Z < Handle_cnt.Coordinates.Z + 150 Then
                                    ReDim Preserve Zoukei_curve_Arr(idem_Zoukei_curve)
                                    Zoukei_curve_Arr(idem_Zoukei_curve) = iCurve
                                    idem_Zoukei_curve = idem_Zoukei_curve + 1
                                End If
                            End If
                        Next
                    Next
                End If

            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 3".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                'Section 3: Handle Inside
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            Call Get_MinDist(iCurve.Tag, CW(1).Tag, Kcach, draft_pnt)
                            Call Get_MinDist(iCurve.Tag, CW(1).Tag, draft, draft_pnt)
                            If Kcach <> 0 And draft <> 0 Then
                                ReDim Preserve BRKT_curve(idem_BRKT)
                                BRKT_curve(idem_BRKT) = CType(iObj, NXOpen.Curve)
                                idem_BRKT = idem_BRKT + 1
                                Call Get_MinDist(iCurve.Tag, glass1_Draft.Tag, Kcach, draft_pnt)
                                'Tim diem do voi Glass
                                If Kcach < kcmin1 Then
                                    kcmin1 = Kcach
                                    If draft_pnt(0).Y < draft_pnt(1).Y Then
                                        diemcantim1 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
                                    Else
                                        diemcantim1 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
                                    End If
                                End If
                            End If
                            'Tim diem do voi REINF OTR
                            Call Get_MinDist(iCurve.Tag, Line_Draft_Upper.Tag, Kcach, draft_pnt)
                            If Kcach < kcmin7 Then
                                kcmin7 = Kcach
                                diemcantim10 = draft_pnt(0)
                            End If

                        End If
                    Next
                Next

            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 4".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then 'Do duong Outside Handle
                'Section 4: Handle Outside
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            'Tim diem do voi saigaisen
                            Call Get_MinDist(iCurve.Tag, Saigaimen_Draft.Tag, Kcach, draft_pnt)
                            If Kcach < kcmin2 Then
                                kcmin2 = Kcach
                                diemcantim2 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
                            End If
                        End If
                    Next
                Next

            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 5".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then 'Do duong Counter Weight(CW) & REINF OTR
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                'Section 5: REINF & CW
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            'Tim diem cao nhat cua 2 CW
                            'CW(0)
                            Call Get_MinDist(iCurve.Tag, CW(0).Tag, Kcach, draft_pnt)
                            If Kcach < 0.04 Then
                                ReDim Preserve CW_curve(idem_CW)
                                CW_curve(idem_CW) = CType(iCurve, NXOpen.Curve)
                                idem_CW = idem_CW + 1
                                Call Get_MinDist(iCurve.Tag, Line_Draft_Upper.Tag, Kcach, draft_pnt)
                                If Kcach < kcmin8 Then
                                    kcmin8 = Kcach
                                    diemcantim9 = draft_pnt(0)
                                End If
                            End If
                            'CW(1)
                            Call Get_MinDist(iCurve.Tag, CW(1).Tag, Kcach, draft_pnt)
                            If Kcach < 0.04 Then

                                ReDim Preserve CW_curve(idem_CW)
                                CW_curve(idem_CW) = CType(iCurve, NXOpen.Curve)
                                idem_CW = idem_CW + 1
                                Call Get_MinDist(iCurve.Tag, Line_Draft_Upper.Tag, Kcach, draft_pnt)
                                If Kcach < kcmin8 Then
                                    kcmin8 = Kcach
                                    diemcantim9 = draft_pnt(0)
                                End If
                            End If
                            'Tim duong REINF OTR
                            Call Get_MinDist(iCurve.Tag, REINF.Tag, Kcach, draft_pnt)
                            If Kcach < 0.04 Then
                                Do_REINF = True
                                ReDim Preserve REINF_curve(idem_REINF)
                                REINF_curve(idem_REINF) = CType(iCurve, NXOpen.Curve)
                                idem_REINF = idem_REINF + 1
                                Call Get_MinDist(iCurve.Tag, Line_Draft_lower.Tag, Kcach, draft_pnt)
                                If Kcach < kcmin9 Then
                                    kcmin9 = Kcach
                                    diemcantim11 = draft_pnt(0)
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        Loop Until tmpGrp = NXOpen.Tag.Null
        '-----------------------------------------------------
        Call Netdut(saigai_sen)
        Call Netdut(glassCL_sen)
        '-----------------------------------------------------
        'Tao diem do
        Dim Diem1_Handle, Diem2_saigai, Diem3_GlassCL, Diem4_Glass, Diem5_BRKT, Diem6_ZoukeiCL, Diem7_CW_01, Diem8_Zoukei, Diem9_REINF_01, Diem10_Handle_BRKT, Diem11_CW, Diem12_CW_02, Diem13_Glass_CL_02 As NXOpen.Point
        Call Lib_NX.Create_Point(diemcantim1.X, diemcantim1.Y, diemcantim1.Z, Diem1_Handle)
        Diem1_Handle.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Call Lib_NX.Create_Point(diemcantim2.X, diemcantim2.Y, diemcantim2.Z, Diem2_saigai)
        Diem2_saigai.SetVisibility(SmartObject.VisibilityOption.Invisible)
        '------------------------------------------------------------------------------------------------------------------------
        'Tim diem gan nhat tren Glass
        Call Get_MinDist(glassCL_sen.Tag, Diem1_Handle.Tag, Kcach, draft_pnt)
        If draft_pnt(0).Y <> Diem1_Handle.Coordinates.Y Then
            diemcantim3 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
        Else
            diemcantim3 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
        End If
        Call Lib_NX.Create_Point(diemcantim3.X, diemcantim3.Y, diemcantim3.Z, Diem3_GlassCL)
        Diem3_GlassCL.SetVisibility(SmartObject.VisibilityOption.Invisible)
        '-----------------------------------------------------
        'PMI
        'Saigaimen vs Handle
        If Diem2_saigai.Coordinates.Y > timdiemdau("Z", True, saigai_sen).Y Then
            Ketluan = "OK"
        Else
            Ketluan = "NG"
        End If
        Dim diemdat2 As Point3d = New Point3d(diemcantim2.X, diemcantim2.Y - 150, diemcantim2.Z)
        Call PMI_vertical_Pnt_Line(Diem2_saigai, saigai_sen, PMI_Plane, diemdat2, False, 0, 999, Ketluan, PMI_size)

        'Glass vs handle BRKT
        If Diem1_Handle.Coordinates.Y < Diem3_GlassCL.Coordinates.Y Then
            Ketluan = "OK"
        Else
            Ketluan = "NG"
        End If
        Dim diemdat3 As Point3d = New Point3d(diemcantim1.X, diemcantim1.Y + 50, diemcantim1.Z + 20)

        Call PMI_Pnt_Pnt(Diem1_Handle, Diem3_GlassCL, Daplane, PMI_Plane, diemdat3, False, 0, 999, Ketluan, PMI_size)

        'Khoang offset Glass
        Call Get_MinDist(glass_sen.Tag, Diem3_GlassCL.Tag, Kcach, draft_pnt)
        If draft_pnt(0).Y <> Diem3_GlassCL.Coordinates.Y Then
            diemcantim4 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
        Else
            diemcantim4 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
        End If
        Call Lib_NX.Create_Point(diemcantim4.X, diemcantim4.Y, diemcantim4.Z, Diem4_Glass)
        Diem4_Glass.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Dim diemdat4 As Point3d = New Point3d(diemcantim4.X + 100, diemcantim4.Y + 50, diemcantim4.Z - 60)

        Call PMI_Pnt_Pnt(Diem4_Glass, Diem3_GlassCL, Daplane, PMI_Plane, diemdat4, False, 0, 999, Nothing, PMI_size)

        '------------------------------------------------------------------------------------------------------------------------
        If PMI_Plane = "Z" And Section_Patan = "SA" Then 'Ko check mat SA
            Call View("Top")

        ElseIf PMI_Plane = "X" Then 'Check SB, SC, SD
            'Zoukei vs handle
            Call View("FR")
            Dim Zoukei_offset_curve(), Zoukei_Draft_curve(), Thickness_curve() As Curve
            Dim del_feature As Feature 'Xoa feature Nhap
            'Offset zoukei CL
            Do
                Try
                    Call Offset_nonSketch(Zoukei_curve_Arr, ZoukeiCL, Zoukei_offset_curve)
                    Exit Do
                Catch ex As Exception
                    Call Re_zoukei_curve(Zoukei_curve_Arr, Zoukei_curve_Arr)
                End Try
            Loop


            'Offset zoukei draft
            Call Offset_nonSketch_draft(Zoukei_curve_Arr, Glass_Offset_KCach, Zoukei_Draft_curve, del_feature)
            'Ofset Zoukei thickness
            Call Offset_nonSketch(Zoukei_curve_Arr, Doday, Thickness_curve)

            Call Lib_NX.Create_Point(diemcantim11.X, diemcantim11.Y, diemcantim11.Z, Diem9_REINF_01)
            Diem9_REINF_01.SetVisibility(SmartObject.VisibilityOption.Invisible)

            If Section_Patan = "SB" Then 'Check SB (Do REINF voi Handle BRKT)
                For Each ispline As Curve In Zoukei_Draft_curve
                    For i As Integer = 0 To UBound(BRKT_curve)
                        'tim diem do tren Handle BRKT
                        Call Get_MinDist(BRKT_curve(i).Tag, ispline.Tag, Kcach, draft_pnt)
                        If Kcach < kcmin3 Then
                            kcmin3 = Kcach
                            diemcantim5 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
                        End If
                    Next
                Next

                Call Lib_NX.Create_Point(diemcantim5.X, diemcantim5.Y, diemcantim5.Z, Diem5_BRKT)
                Diem5_BRKT.SetVisibility(SmartObject.VisibilityOption.Invisible)
                'tim diem do tren Zoukei
                For Each ispline As Curve In Zoukei_offset_curve
                    Call Get_MinDist(Diem5_BRKT.Tag, ispline.Tag, Kcach, draft_pnt)
                    If Kcach < kcmin4 Then
                        kcmin4 = Kcach
                        diemcantim6 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
                    End If
                Next

                Call Lib_NX.Create_Point(diemcantim6.X, diemcantim6.Y, diemcantim6.Z, Diem6_ZoukeiCL)
                Diem6_ZoukeiCL.SetVisibility(SmartObject.VisibilityOption.Invisible)
                If Diem5_BRKT.Coordinates.Y > Diem6_ZoukeiCL.Coordinates.Y Then
                    Ketluan = "OK"
                Else
                    Ketluan = "NG"
                End If
                Dim diemdat5 As Point3d = New Point3d(diemcantim6.X, diemcantim6.Y - 50, diemcantim6.Z + 10)

                'PMI BRKT vs Zoukei CL
                Call PMI_Pnt_Pnt(Diem5_BRKT, Diem6_ZoukeiCL, Daplane, PMI_Plane, diemdat5, False, 0, 999, Ketluan, PMI_size)

                'do voi REINF OTR
                Call Lib_NX.Create_Point(diemcantim10.X, diemcantim10.Y, diemcantim10.Z, Diem10_Handle_BRKT)
                Diem10_Handle_BRKT.SetVisibility(SmartObject.VisibilityOption.Invisible)
                Dim diemdat_03 As Point3d = New Point3d(diemcantim10.X, diemcantim10.Y + 150, diemcantim10.Z)
                Call PMI_Vertical_Pnt_Pnt(Diem9_REINF_01, Diem10_Handle_BRKT, Daplane, PMI_Plane, diemdat_03, True, 0, 999, Nothing, PMI_size)

            Else 'Check SC, SD
                'tim diem do tren CW

                For Each ispline As Curve In Zoukei_Draft_curve
                    For i As Integer = 0 To UBound(CW_curve)
                        Call Get_MinDist(CW_curve(i).Tag, ispline.Tag, Kcach, draft_pnt)
                        If Kcach < kcmin5 Then
                            kcmin5 = Kcach
                            diemcantim7 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
                        End If
                        Call Get_MinDist(CW_curve(i).Tag, glass1_Draft.Tag, Kcach, draft_pnt)
                        If Kcach < kcmin11 Then
                            kcmin11 = Kcach
                            diemcantim12 = New Point3d(draft_pnt(0).X, draft_pnt(0).Y, draft_pnt(0).Z)
                        End If
                    Next
                Next

                Call Lib_NX.Create_Point(diemcantim12.X, diemcantim12.Y, diemcantim12.Z, Diem12_CW_02)
                Diem12_CW_02.SetVisibility(SmartObject.VisibilityOption.Invisible)
                Call Lib_NX.Create_Point(diemcantim7.X, diemcantim7.Y, diemcantim7.Z, Diem7_CW_01)
                Diem7_CW_01.SetVisibility(SmartObject.VisibilityOption.Invisible)
                'tim diem do tren Zoikei
                For Each ispline As Curve In Zoukei_offset_curve
                    Call Get_MinDist(Diem7_CW_01.Tag, ispline.Tag, Kcach, draft_pnt)
                    If Kcach < kcmin10 Then
                        kcmin10 = Kcach
                        diemcantim8 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
                    End If
                    Call Get_MinDist(Diem12_CW_02.Tag, ispline.Tag, Kcach, draft_pnt)
                Next

                Call Lib_NX.Create_Point(diemcantim8.X, diemcantim8.Y, diemcantim8.Z, Diem8_Zoukei)
                Diem8_Zoukei.SetVisibility(SmartObject.VisibilityOption.Invisible)
                'BRKT vs Glass
                If Diem7_CW_01.Coordinates.Y > Diem8_Zoukei.Coordinates.Y Then
                    Ketluan = "OK"
                Else
                    Ketluan = "NG"
                End If
                Dim diemdat6 As Point3d = New Point3d(diemcantim7.X, diemcantim7.Y - 50, diemcantim7.Z)
                Call PMI_Pnt_Pnt(Diem7_CW_01, Diem8_Zoukei, Daplane, PMI_Plane, diemdat6, False, 0, 999, Ketluan, PMI_size)

                'CW vs Glass
                Call Get_MinDist(Diem12_CW_02.Tag, glassCL_sen.Tag, Kcach, draft_pnt)
                If Kcach < kcmin12 Then
                    kcmin12 = Kcach
                    diemcantim13 = New Point3d(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z)
                End If
                Call Lib_NX.Create_Point(diemcantim13.X, diemcantim13.Y, diemcantim13.Z, Diem13_Glass_CL_02)
                Diem13_Glass_CL_02.SetVisibility(SmartObject.VisibilityOption.Invisible)
                If Diem12_CW_02.Coordinates.Y < Diem13_Glass_CL_02.Coordinates.Y Then
                    Ketluan = "OK"
                Else
                    Ketluan = "NG"
                End If
                Dim diemdat7 As Point3d = New Point3d(diemcantim12.X, diemcantim12.Y + 50, diemcantim12.Z + 50)
                Call PMI_Pnt_Pnt(Diem12_CW_02, Diem13_Glass_CL_02, Daplane, PMI_Plane, diemdat7, False, 0, 999, Ketluan, PMI_size)

                'do voi REINF OTR
                If Do_REINF = True Then
                    Call Lib_NX.Create_Point(diemcantim9.X, diemcantim9.Y, diemcantim9.Z, Diem11_CW)
                    Diem11_CW.SetVisibility(SmartObject.VisibilityOption.Invisible)
                    Dim diemdat_04 As Point3d = New Point3d(diemcantim9.X, diemcantim9.Y + 150, diemcantim9.Z)

                    If Diem11_CW.Coordinates.Z < Diem9_REINF_01.Coordinates.Z Then
                        Call PMI_Vertical_Pnt_Pnt(Diem11_CW, Diem9_REINF_01, Daplane, PMI_Plane, diemdat_04, True, 0, 999, Nothing, PMI_size)
                    End If
                End If
            End If

            'Xoa draft
            Try
                theUFSession.Obj.DeleteObject(del_feature.Tag)
            Catch ex As Exception
                lw.WriteLine(ex.ToString)
            End Try
            'doi thanh net dut
            For Each ispline As Curve In Zoukei_offset_curve
                Call Netdut(ispline)
            Next



            'Do khoang offset va do day Zoukei
            Dim Diemdo_Zoukei, Diemdo_Zoukei_CL, Diemdo_Zoukei_Thickness As NXOpen.Point
            Dim diem_Zoukei, diem_Zoukei_CL, diem_Zoukei_Thickness, Diemdat_01, Diemdat_02, Drft_pnt As Point3d
            diem_Zoukei = New Point3d(1, 1, 99999)
            diem_Zoukei_CL = New Point3d(1, 1, 99999)
            diem_Zoukei_Thickness = New Point3d(1, 1, 99999)

            'Tim diem thap nhat Zoukei
            For i As Integer = 0 To UBound(Zoukei_curve_Arr)
                Drft_pnt = timdiemdau("Z", True, Zoukei_curve_Arr(i))
                If Drft_pnt.Z < diem_Zoukei.Z Then
                    diem_Zoukei = Drft_pnt
                End If
            Next

            'Tim diem thap nhat Zoukei offset curve
            For i As Integer = 0 To UBound(Zoukei_offset_curve)
                Drft_pnt = timdiemdau("Z", True, Zoukei_offset_curve(i))
                If Drft_pnt.Z < diem_Zoukei_CL.Z Then
                    diem_Zoukei_CL = Drft_pnt
                End If
            Next

            'Tim diem thap nhat Zoukei Thickness curve
            For i As Integer = 0 To UBound(Thickness_curve)
                Drft_pnt = timdiemdau("Z", True, Thickness_curve(i))
                If Drft_pnt.Z < diem_Zoukei_Thickness.Z Then
                    diem_Zoukei_Thickness = Drft_pnt
                End If
            Next

            Call Lib_NX.Create_Point(diem_Zoukei.X, diem_Zoukei.Y, diem_Zoukei.Z, Diemdo_Zoukei)
            Diemdo_Zoukei.SetVisibility(SmartObject.VisibilityOption.Invisible)
            Call Lib_NX.Create_Point(diem_Zoukei_CL.X, diem_Zoukei_CL.Y, diem_Zoukei_CL.Z, Diemdo_Zoukei_CL)
            Diemdo_Zoukei_CL.SetVisibility(SmartObject.VisibilityOption.Invisible)
            Call Lib_NX.Create_Point(diem_Zoukei_Thickness.X, diem_Zoukei_Thickness.Y, diem_Zoukei_Thickness.Z, Diemdo_Zoukei_Thickness)
            Diemdo_Zoukei_Thickness.SetVisibility(SmartObject.VisibilityOption.Invisible)
            Diemdat_01 = New Point3d(diem_Zoukei.X, diem_Zoukei.Y + 50, diem_Zoukei.Z)
            Diemdat_02 = New Point3d(diem_Zoukei.X, diem_Zoukei.Y - 50, diem_Zoukei.Z + 25)
            'PMI Zoukei CL va PMI Zoukei Thickness
            Call PMI_Pnt_Pnt(Diemdo_Zoukei_Thickness, Diemdo_Zoukei_CL, Daplane, PMI_Plane, Diemdat_01, False, 0, 999, Nothing, PMI_size)
            Call PMI_Pnt_Pnt(Diemdo_Zoukei, Diemdo_Zoukei_Thickness, Daplane, PMI_Plane, Diemdat_02, False, 0, 999, Nothing, PMI_size)

        Else

        End If

        Call Xoay_Datum(Handle_cnt, Daplane)

        Call Fit_to_PMI()
        Call zentaizu(ImgeFolder1, Section_Patan & ".png")
        Call Blank(iprt_Comp)
    End Sub
    Sub Re_zoukei_curve(ByVal zoukei_curve() As Curve, ByRef Out_curve() As Curve)

        Dim Top_line As Curve = zoukei_curve(0)
        For Each icurve As Curve In zoukei_curve
            If timdiemdau("Z", False, (icurve)).Z > timdiemdau("Z", False, Top_line).Z Then
                Top_line = icurve
            End If
        Next

        Dim idem As Integer
        For Each icurve As Curve In zoukei_curve
            If icurve.Tag <> Top_line.Tag Then
                ReDim Preserve Out_curve(idem)
                Out_curve(idem) = icurve
                idem = idem + 1
            End If
        Next

    End Sub
    Sub Do_PMI_Angle(ByVal line1 As Curve, ByVal line2 As Curve)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUfSession As UFSession = UFSession.GetUFSession()
        Call Lib_NX.SetWorkPart(line1.OwningComponent, workPart)
        Call PMI_Angle(line1, line2, PMI_size)
    End Sub
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#End Region

#Region "2_Tsukaiyasui"
    Sub Main2(ByVal Handle_cnt As NXOpen.Point, ByVal Input_Datum As NXOpen.DatumPlane(), ByVal In_Handle() As Body, ByVal Out_Handle() As Body, ByVal iZoukei As NXOpen.Assemblies.Component)
        lw.Open()
        Try
            theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)
            Call Lib_NX.Create_Assy("Tsukaiyasui", CL_assy, True)
            For i As Integer = 0 To UBound(In_Handle)
                ReDim Preserve Handle_Inside(i)
                Handle_Inside(i) = In_Handle(i)
            Next

            For i As Integer = 0 To UBound(Out_Handle)
                ReDim Preserve Handle_Outside(i)
                Handle_Outside(i) = Out_Handle(i)
            Next
            Zoukei_Comp = iZoukei
            Call SA_2(Input_Datum, Handle_cnt, CL_assy)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try

    End Sub
    Sub SA_2(ByVal Datum_In As NXOpen.DatumPlane(), ByVal Handle_center As NXOpen.Point, ByVal Acomp As Assemblies.Component)
        Call Sect_Order2("SA", PartSA_comp, Datum_In(0), Acomp)
        Dim diemthuocZoukei As Point3d
        'Dim zoukei_curve_out() As Curve
        Dim Handle_curve1, Handle_curve2 As Curve
        Dim Joint_Zoukei, Joint_Handle As Feature
        Call Zoukei(Handle_center, diemthuocZoukei, Nothing, Joint_Zoukei)
        Call Handle_OS(diemthuocZoukei, Nothing, Joint_Handle, Handle_curve1, Handle_curve2)
        Call check(Datum_In(0), Handle_center, Joint_Zoukei, Joint_Handle, Handle_curve1, Handle_curve2)
        Joint_Zoukei.HideBody()
        Joint_Handle.HideBody()
    End Sub
    Sub Handle_OS(ByVal diem_Zoukei3D As Point3d, ByRef H_curve_Arr() As Curve, ByRef ifeature As Feature, ByRef H_curve1 As Curve, ByRef H_curve2 As Curve)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim tmpGrp As NXOpen.Tag = NXOpen.Tag.Null
        Dim myGroup As NXOpen.Group
        Dim member_count As Integer
        Dim group_member_list() As Tag
        Dim group_member_list1() As Tag

        Dim H_group As NXOpen.Group
        Dim Diem_Zoukei As NXOpen.Point
        Call Lib_NX.Create_Point(diem_Zoukei3D.X, diem_Zoukei3D.Y, diem_Zoukei3D.Z, Diem_Zoukei)
        Diem_Zoukei.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Dim idem_z As Integer = 0
        Dim kc1 As Double = 9999
        Dim H_curve As Curve


        Do
            theUFSession.Obj.CycleObjsInPart(workPart.Tag, UFConstants.UF_group_type, tmpGrp)
            If tmpGrp = NXOpen.Tag.Null Then
                Continue Do
            End If
            Dim theType As Integer, theSubtype As Integer
            theUFSession.Obj.AskTypeAndSubtype(tmpGrp, theType, theSubtype)

            If theSubtype = UFConstants.UF_group_type Then
                myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            End If
            myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            Dim guess1(2) As Double
            Dim guess2(2) As Double
            Dim pt1(2) As Double
            Dim pt2(2) As Double
            Dim junk(2) As Double, trash As Double
            If InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) > 0 Then
                Continue Do
            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 3".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                'Section 1: Saigaishoku & Glass, Glass CL
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            Dim StartPoint(2) As Double
                            Dim midPt(2) As Double
                            Dim EndPoint(2) As Double
                            ''------------------------------------------------------------------------------------------------------------------------

                            Call Get_MinDist(iCurve.Tag, Diem_Zoukei.Tag, Kcach, draft_pnt)
                            If Kcach < kc1 Then
                                kc1 = Kcach
                                H_curve = iCurve
                                Dim obj As NXOpen.TaggedObject = NXOpen.Utilities.NXObjectManager.Get(iGrp)
                                H_group = CType(obj, NXOpen.Group)

                            End If
                            ''------------------------------------------------------------------------------------------------------------------------

                        End If
                    Next
                Next
            End If
        Loop Until tmpGrp = NXOpen.Tag.Null


        theUFSession.Group.AskGroupData(H_group.Tag, group_member_list, member_count)
        Dim curve_Arr0() As Curve
        Dim idemc As Integer
        For Each itag As Tag In group_member_list
            Dim iObj As NXObject = Nothing
            iObj = NXObjectManager.Get(itag)

            If iObj.GetType.ToString <> "NXOpen.Group" Then
                Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                If iCurve.Tag <> H_curve.Tag Then
                    Call Get_MinDist(iCurve.Tag, H_curve.Tag, Kcach, draft_pnt)
                    If Kcach = 0 And timdiemdau("X", False, iCurve).X > diem_Zoukei3D.X Then
                        H_curve1 = iCurve
                        'H_curve1.Highlight()
                    ElseIf Kcach = 0 And timdiemdau("X", True, iCurve).X < diem_Zoukei3D.X Then
                        H_curve2 = iCurve
                        'H_curve2.Highlight()
                    End If
                End If
            End If
        Next



        Call Join_curve_Group(H_group, H_curve, ifeature)

        Exit Sub
        ''''chọn tangent Curve
        Dim iCurveTangentRule As NXOpen.CurveChainRule = Nothing
        iCurveTangentRule = workPart.ScRuleFactory.CreateRuleCurveChain(H_curve, Nothing, False, 0.0095)
        Dim iRules(0) As NXOpen.SelectionIntentRule
        iRules(0) = iCurveTangentRule
        Dim scCollector1 As ScCollector = workPart.ScCollectors.CreateCollector
        scCollector1.ReplaceRules(iRules, False)
        Dim tmpCurve As Curve
        Dim Curve_tag() As Tag
        Dim selectionEdges() As TaggedObject
        selectionEdges = scCollector1.GetObjects
        For Each obj As TaggedObject In selectionEdges
            tmpCurve = CType(obj, Curve)
            tmpCurve.Highlight()
            ReDim Preserve H_curve_Arr(idemc)
            H_curve_Arr(idemc) = tmpCurve
            ReDim Preserve Curve_tag(idemc)
            Curve_tag(idemc) = tmpCurve.Tag
            H_curve_Arr(idemc).Highlight()
            idemc = idemc + 1
        Next
        Dim out_curve_tag As Tag
        'Call SelectjoinCurve(Curve_tag, idemc, out_curve_tag)
        Call Joint_curve_Arr(H_curve_Arr, ifeature)

    End Sub
    Sub SelectjoinCurve(ByVal curves() As Tag, ByVal num_curves As Integer, ByRef iCurve2 As String)

        Dim inx As Integer = 0
        'Dim curves() As NXOpen.Tag
        Dim num_joined As Integer
        Dim n As String = vbCrLf
        Dim join_type As Integer = 2

        If (num_curves) > 0 Then

            Dim joined(num_curves) As NXOpen.Tag

            theUFSession.Curve.AutoJoinCurves(curves, num_curves,
                                     join_type, joined, num_joined)
            theUFSession.Ui.OpenListingWindow()
            'theUFSession.Ui.WriteListingWindow("Joined curve output count: " & _
            '                                    num_joined.ToString & n)

            For inx = 0 To num_joined - 1
                'theUFSession.Ui.WriteListingWindow("Joined: " & joined(inx).ToString & n)
                iCurve2 = joined(inx).ToString
            Next

        End If

    End Sub
    Sub Zoukei(ByVal H_Center As NXOpen.Point, ByRef diem_Zoukei As Point3d, ByRef Zoukei_curve_Arr() As Curve, ByRef ifeature As Feature)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim tmpGrp As NXOpen.Tag = NXOpen.Tag.Null
        Dim myGroup As NXOpen.Group
        Dim member_count As Integer
        Dim group_member_list() As Tag
        Dim group_member_list1() As Tag


        Dim ZoukeiGroup As NXOpen.Group
        Dim idem_z As Integer = 0
        Dim kc1 As Double = 9999


        Do
            theUFSession.Obj.CycleObjsInPart(workPart.Tag, UFConstants.UF_group_type, tmpGrp)
            If tmpGrp = NXOpen.Tag.Null Then
                Continue Do
            End If
            Dim theType As Integer, theSubtype As Integer
            theUFSession.Obj.AskTypeAndSubtype(tmpGrp, theType, theSubtype)

            If theSubtype = UFConstants.UF_group_type Then
                myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            End If
            myGroup = Utilities.NXObjectManager.Get(tmpGrp)
            Dim guess1(2) As Double
            Dim guess2(2) As Double
            Dim pt1(2) As Double
            Dim pt2(2) As Double
            Dim junk(2) As Double, trash As Double
            If InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) > 0 Then
                Continue Do
            ElseIf InStr(myGroup.Name.ToString.ToUpper, "Section 1".ToUpper) > 0 And InStr(myGroup.Name.ToString.ToUpper, "Dynamic Sections".ToUpper) = 0 Then
                theUFSession.Group.AskGroupData(myGroup.Tag, group_member_list, member_count)
                ZoukeiGroup = myGroup
                'Section 1: Saigaishoku & Glass, Glass CL
                For Each iGrp As Tag In group_member_list
                    theUFSession.Group.AskGroupData(iGrp, group_member_list1, member_count)
                    For Each iTag As Tag In group_member_list1
                        Dim iObj As NXObject = Nothing
                        iObj = NXObjectManager.Get(iTag)

                        If iObj.GetType.ToString <> "NXOpen.Group" Then
                            Dim iCurve As NXOpen.Curve = CType(iObj, NXOpen.Curve)
                            Dim StartPoint(2) As Double
                            Dim midPt(2) As Double
                            Dim EndPoint(2) As Double
                            '------------------------------------------------------------------------------------------------------------------------

                            ReDim Preserve Zoukei_curve_Arr(idem_z)
                            Zoukei_curve_Arr(idem_z) = iCurve
                            idem_z = idem_z + 1
                            Call Get_MinDist(iCurve.Tag, H_Center.Tag, Kcach, draft_pnt)
                            If Kcach < kc1 Then
                                kc1 = Kcach
                                diem_Zoukei = draft_pnt(0)
                            End If

                            '------------------------------------------------------------------------------------------------------------------------
                        End If
                    Next
                Next
            End If
        Loop Until tmpGrp = NXOpen.Tag.Null
        Call Join_curve_Group(ZoukeiGroup, Zoukei_curve_Arr(0), ifeature)


    End Sub
    Sub Sect_Order2(ByVal Ten_Comp As String, ByVal Comp As Assemblies.Component, ByVal Datum_input As NXOpen.DatumPlane, ByVal A_comp As Assemblies.Component)

        Call Lib_NX.Create_Component(Ten_Comp, Comp, True)
        For Each iChild As Component In iRootComp.GetChildren
            Call Blank(iChild)
        Next
        Call Unblank(A_comp)
        'Hien Zoukei
        Zoukei_Comp.Unblank()
        'Tao matcat
        Call CreateSect(Datum_input)
        Zoukei_Comp.Blank()
        '------------------------------------------------
        'Section 2: Handle Inside
        'Hien Handle Inside
        For i As Integer = 0 To UBound(Handle_Inside)
            Call Unblank(Handle_Inside(i).OwningComponent)
        Next
        'Tao matcat
        Call CreateSect(Datum_input)
        For i As Integer = 0 To UBound(Handle_Inside)
            Call Blank(Handle_Inside(i).OwningComponent)
        Next
        '------------------------------------------------
        'Section 4: Handle Outside
        'Hien Handle Outside
        For i As Integer = 0 To UBound(Handle_Outside)
            Call Unblank(Handle_Outside(i).OwningComponent)
        Next
        'Tao matcat
        Call CreateSect(Datum_input)
        For i As Integer = 0 To UBound(Handle_Outside)
            Call Blank(Handle_Outside(i).OwningComponent)
        Next
        '------------------------------------------------
    End Sub
    Sub check(ByVal iplane As DatumPlane, ByVal SE_pnt As NXOpen.Point, ByVal Zoukei_Feat As Feature, ByVal Hanle_Feat As Feature, ByVal curve_01 As Curve, ByVal curve_02 As Curve)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        theSession.BeginTaskEnvironment()
        ' ----------------------------------------------

        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin

        sketchInPlaceBuilder1.OriginOptionInfer = NXOpen.OriginMethod.WorkPartOrigin

        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)

        Dim simpleSketchInPlaceBuilder1 As NXOpen.SimpleSketchInPlaceBuilder = Nothing
        simpleSketchInPlaceBuilder1 = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()

        sketchAlongPathBuilder1.PlaneLocation.Expression.SetFormula("0")

        simpleSketchInPlaceBuilder1.UseWorkPartOrigin = True

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim waveDatumBuilder1 As NXOpen.Features.WaveDatumBuilder = Nothing
        waveDatumBuilder1 = workPart.Features.CreateWaveDatumBuilder(nullNXOpen_Features_Feature)

        waveDatumBuilder1.Associative = True

        Dim selectObjectList1 As NXOpen.SelectObjectList = Nothing
        selectObjectList1 = waveDatumBuilder1.Datums

        'Dim component1 As NXOpen.Assemblies.Component = CType(displayPart.ComponentAssembly.RootComponent.FindObject("COMPONENT @DB/NML49842728/AA 1"), NXOpen.Assemblies.Component)

        Dim datumPlane1 As NXOpen.DatumPlane = iplane 'CType(component1.FindObject("PROTO#ENTITY 197 1 1"), NXOpen.DatumPlane)

        Dim added1 As Boolean = Nothing
        added1 = selectObjectList1.Add(datumPlane1)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = waveDatumBuilder1.CommitCreateOnTheFly()

        Dim waveLinkRepository1 As NXOpen.GeometricUtilities.WaveLinkRepository = Nothing
        waveLinkRepository1 = workPart.CreateWavelinkRepository()

        waveLinkRepository1.SetNonFeatureApplication(False)

        waveLinkRepository1.SetBuilder(simpleSketchInPlaceBuilder1)

        Dim waveDatum1 As NXOpen.Features.WaveDatum = CType(feature1, NXOpen.Features.WaveDatum)

        waveLinkRepository1.SetLink(waveDatum1)

        Dim waveDatumBuilder2 As NXOpen.Features.WaveDatumBuilder = Nothing
        waveDatumBuilder2 = workPart.Features.CreateWaveDatumBuilder(waveDatum1)

        waveDatumBuilder2.Associative = False

        Dim feature2 As NXOpen.Features.Feature = Nothing
        feature2 = waveDatumBuilder2.CommitCreateOnTheFly()

        waveDatumBuilder2.Destroy()

        Dim waveDatum2 As NXOpen.Features.WaveDatum = CType(feature2, NXOpen.Features.WaveDatum)

        waveDatum2.SetName("Datum")

        waveDatumBuilder1.Destroy()

        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(SE_pnt.Coordinates.X, SE_pnt.Coordinates.Y, SE_pnt.Coordinates.Z)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) X axis"), NXOpen.DatumAxis)

        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim success1 As Boolean = Nothing
        success1 = direction1.ReverseDirection()

        Dim datumPlane2 As NXOpen.DatumPlane = CType(workPart.Datums.FindObject("LINKED_DATUM_PLANE(3)"), NXOpen.DatumPlane)

        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(datumPlane2, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, False)
        '--------------------------------------------------------------------------------
        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        simpleSketchInPlaceBuilder1.CoordinateSystem = cartesianCoordinateSystem1

        Dim iDirX As Vector3d, iDirY As Vector3d

        cartesianCoordinateSystem1.GetDirections(iDirX, iDirY)

        '--------------------------------------------------------------------------------
        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1

        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1

        Dim nErrs1 As Integer = Nothing
        nErrs1 = theSession.UpdateManager.AddToDeleteList(point1)

        theSession.Preferences.Sketch.CreateInferredConstraints = True

        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False

        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression

        theSession.Preferences.Sketch.TextSizeFixed = True

        theSession.Preferences.Sketch.FixedTextSize = 3.0

        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True

        theSession.Preferences.Sketch.DisplayReferenceGeometry = True

        theSession.Preferences.Sketch.DisplayShadedRegions = True

        theSession.Preferences.Sketch.FindMovableObjects = True

        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0

        theSession.Preferences.Sketch.DisplayObjectColor = True

        theSession.Preferences.Sketch.DisplayObjectName = True

        theSession.Preferences.Sketch.EditDimensionOnCreation = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = simpleSketchInPlaceBuilder1.Commit()

        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)

        Dim feature3 As NXOpen.Features.Feature = Nothing
        feature3 = sketch1.Feature

        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)

        sketchInPlaceBuilder1.Destroy()

        sketchAlongPathBuilder1.Destroy()

        simpleSketchInPlaceBuilder1.Destroy()

        plane1.DestroyPlane()

        waveLinkRepository1.Destroy()

        theSession.ActiveSketch.SetName("SKETCH_000")

        theSession.CleanUpFacetedFacesAndEdges()
        ' ----------------------------------------------
        'Main
        Dim diem1 As NXOpen.Point
        Call Lib_NX.Create_Point(SE_pnt.Coordinates.X, SE_pnt.Coordinates.Y, SE_pnt.Coordinates.Z, diem1)
        Dim p1_1 As Point3d = New Point3d(SE_pnt.Coordinates.X, SE_pnt.Coordinates.Y, SE_pnt.Coordinates.Z)
        Dim vt001 As Vector3d = New Vector3d(iDirX.X, iDirX.Y, iDirX.Z)
        Dim vt002 As Vector3d = New Vector3d(-1 * iDirX.X, -1 * iDirX.Y, -1 * iDirX.Z)
        Dim i_long As Double = 20
        Dim p1_2_01 As Point3d = New Point3d(p1_1.X + i_long * vt001.X, p1_1.Y + i_long * vt001.Y, p1_1.Z + i_long * vt001.Z)
        Dim p1_2_02 As Point3d = New Point3d(p1_1.X + i_long * vt002.X, p1_1.Y + i_long * vt002.Y, p1_1.Z + i_long * vt002.Z)
        Dim idraft_line1 As NXOpen.Line = Nothing
        idraft_line1 = workPart.Curves.CreateLine(p1_1, p1_2_01)
        idraft_line1.SetVisibility(SmartObject.VisibilityOption.Invisible)
        'theSession.ActiveSketch.AddGeometry(idraft_line1, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        'idraft_line1.Highlight()
        Dim idraft_line2 As NXOpen.Line = Nothing
        idraft_line2 = workPart.Curves.CreateLine(p1_1, p1_2_02)
        idraft_line2.SetVisibility(SmartObject.VisibilityOption.Invisible)
        'theSession.ActiveSketch.AddGeometry(idraft_line1, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        'idraft_line2.Highlight()
        Dim center1, center2 As Point3d
        center1 = timdiemdauKhacgoc(idraft_line1, p1_1)
        center2 = timdiemdauKhacgoc(idraft_line2, p1_1)
        Dim arc1, arc2 As NXOpen.Arc
        Dim bankinh As Double = 19 / 2
        Dim nXMatrix1 As NXOpen.NXMatrix = Nothing
        nXMatrix1 = theSession.ActiveSketch.Orientation
        arc1 = workPart.Curves.CreateArc(center1, nXMatrix1, bankinh, 0.0, (360.0 * Math.PI / 180.0))
        arc2 = workPart.Curves.CreateArc(center2, nXMatrix1, bankinh, 0.0, (360.0 * Math.PI / 180.0))
        theSession.ActiveSketch.AddGeometry(arc1, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        theSession.ActiveSketch.AddGeometry(arc2, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        Call Rangbuocduongkinh(19, arc1, 1)
        Call Rangbuocduongkinh(19, arc2, 2)

        Dim kc1 As Integer = 9999
        Dim kc2 As Integer = 9999
        Dim Zoukei_curve1, Zoukei_curve2, Handle_curve1, Handle_curve2 As Curve
        For Each i_curve As Curve In Zoukei_Feat.GetEntities
            Call Get_MinDist(i_curve.Tag, arc1.Tag, Kcach, draft_pnt)
            If Kcach < kc1 Then
                kc1 = Kcach
                Zoukei_curve1 = i_curve
            End If
            Call Get_MinDist(i_curve.Tag, arc2.Tag, Kcach, draft_pnt)
            If Kcach < kc2 Then
                kc2 = Kcach
                Zoukei_curve2 = i_curve
            End If
        Next

        Call ttcurvevsarc(Zoukei_curve1, arc1)
        Call ttcurvevsarc(Zoukei_curve2, arc2)

        'For Each i_curve As Curve In Hanle_Feat.GetEntities
        '    Call Get_MinDist(i_curve.Tag, arc1.Tag, Kcach, draft_pnt)
        '    If Kcach < kc1 Then
        '        kc1 = Kcach
        '        Handle_curve1 = i_curve
        '    End If
        '    Call Get_MinDist(i_curve.Tag, arc2.Tag, Kcach, draft_pnt)
        '    If Kcach < kc2 Then
        '        kc2 = Kcach
        '        Handle_curve2 = i_curve
        '    End If
        'Next

        Call ttcurvevsarc(curve_01, arc1)
        Call ttcurvevsarc(curve_02, arc2)
        ' ----------------------------------------------
        i_long = 100
        Dim p1_2_03 As Point3d = New Point3d(p1_1.X + i_long * vt001.X, p1_1.Y + i_long * vt001.Y, p1_1.Z + i_long * vt001.Z)
        Dim p1_2_04 As Point3d = New Point3d(p1_1.X + i_long * vt002.X, p1_1.Y + i_long * vt002.Y, p1_1.Z + i_long * vt002.Z)

        Dim vecto001 As Vector3d = New Vector3d(iDirY.X, iDirY.Y, iDirY.Z)
        Dim vecto002 As Vector3d = New Vector3d(-1 * iDirY.X, -1 * iDirY.Y, -1 * iDirY.Z)

        Dim p1_2_07 As Point3d = New Point3d(p1_2_04.X + i_long * vecto001.X, p1_2_04.Y + i_long * vecto001.Y, p1_2_04.Z + i_long * vecto001.Z)
        Dim p1_2_08 As Point3d = New Point3d(p1_2_04.X + i_long * vecto002.X, p1_2_04.Y + i_long * vecto002.Y, p1_2_04.Z + i_long * vecto002.Z)
        Dim Line001, line002 As Line
        Line001 = workPart.Curves.CreateLine(p1_2_07, p1_2_08)
        Line001.SetVisibility(SmartObject.VisibilityOption.Invisible)

        Dim p1_2_05 As Point3d = New Point3d(p1_2_03.X + i_long * vecto001.X, p1_2_03.Y + i_long * vecto001.Y, p1_2_03.Z + i_long * vecto001.Z)
        Dim p1_2_06 As Point3d = New Point3d(p1_2_03.X + i_long * vecto002.X, p1_2_03.Y + i_long * vecto002.Y, p1_2_03.Z + i_long * vecto002.Z)
        line002 = workPart.Curves.CreateLine(p1_2_05, p1_2_06)
        line002.SetVisibility(SmartObject.VisibilityOption.Invisible)


        Call Get_MinDist(Line001.Tag, arc1.Tag, Kcach, draft_pnt)
        Dim diemdo1, diemdo2 As NXOpen.Point
        Call Lib_NX.Create_Point(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z, diemdo1)
        diemdo1.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Call Get_MinDist(line002.Tag, arc2.Tag, Kcach, draft_pnt)
        Call Lib_NX.Create_Point(draft_pnt(1).X, draft_pnt(1).Y, draft_pnt(1).Z, diemdo2)
        diemdo2.SetVisibility(SmartObject.VisibilityOption.Invisible)

        Dim diemdat001 As Point3d = New Point3d(diemdo2.Coordinates.X, diemdo2.Coordinates.Y - 50, diemdo2.Coordinates.Z)
        Call PMI_Pnt_Pnt(diemdo1, diemdo2, iplane, "Z", diemdat001, True, 0, 9999, Nothing, 5)
        ' ----------------------------------------------
        theSession.Preferences.Sketch.SectionView = False
        theSession.ActiveSketch.Deactivate(NXOpen.Sketch.ViewReorient.False, NXOpen.Sketch.UpdateLevel.Model)
        'theSession.DeleteUndoMarksSetInTaskEnvironment()
        theSession.EndTaskEnvironment()
        Call PMI_Diameter(arc1, TIMDIEMTHUOCDUONGTRON(arc1), iplane, "Z", 9999)
        Call PMI_Diameter(arc2, TIMDIEMTHUOCDUONGTRON(arc2), iplane, "Z", 9999)
    End Sub
#End Region

#Region "3_Doorzone"
    Sub Main3(ByVal SE1() As TaggedObject, ByVal Center_pnt() As TaggedObject, ByVal Hinge_axis() As TaggedObject, ByVal Back_curve() As TaggedObject, ByVal B_Center_pnt() As TaggedObject, ByVal double00() As Double)
        lw.Open()
        Call iResetFolder(ImgeFolder3)
        Call dodayduong(True)
        Start_Grd_pnt = New Point3d(-1 * double00(1), 0, -1 * double00(0))
        End_Grd_pnt = New Point3d(double00(3), 0, -1 * double00(2))

        For ii As Integer = 0 To UBound(double00)
            ReDim Preserve Condition_IN(ii)
            Condition_IN(ii) = double00(ii)
        Next

        Try
            Dim p1 As NXOpen.Point
            Dim p2 As NXOpen.Point
            Dim p3 As NXOpen.Point
            Dim Axis_in As NXOpen.Curve
            Dim Back_Curve_in As NXOpen.Curve
            Axis_in = CType(Hinge_axis(0), NXOpen.Curve)
            Back_Curve_in = CType(Back_curve(0), NXOpen.Curve)


            Try
                p1 = CType(SE1(0), NXOpen.Point)
            Catch ex As Exception
                lw.WriteLine("23" & ex.ToString)
            End Try

            Try
                p2 = CType(Center_pnt(0), NXOpen.Point)
            Catch ex As Exception
                lw.WriteLine(ex.ToString)
            End Try

            Try
                p3 = CType(B_Center_pnt(0), NXOpen.Point)
            Catch ex As Exception
                lw.WriteLine(ex.ToString)
            End Try


            'active Top Assy
            Dim iComTop As NXOpen.Assemblies.Component = Nothing
            Dim partLoadStatus1 As NXOpen.PartLoadStatus = Nothing
            theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)

            'Dim Assy_check As NXOpen.Assemblies.Component
            'Call Lib_NX.Create_Assy("Handle Check", Assy_check, True)


            Call Doorzone(p1, p2, p3, Axis_in, Back_Curve_in)
            theSession.Parts.SetWorkComponent(iComTop, partLoadStatus1)
            Call dodayduong(False)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try




    End Sub
    Sub Doorzone(ByVal p1_1 As NXOpen.Point, ByVal p1_2 As NXOpen.Point, ByVal p1_3 As NXOpen.Point, ByVal Hinge_Center_axis As NXOpen.Curve, ByVal Back_line As NXOpen.Curve)
        'Dim In_Out As NXOpen.Assemblies.Component
        'Call Lib_NX.Create_Assy("In_Out", In_Out, True)
        Dim Front_Check As NXOpen.Assemblies.Component
        Call Lib_NX.Create_Component("Door zone", Front_Check, True)

        Try
            Call Fronto_Check(p1_1, p1_2, Hinge_Center_axis)
        Catch ex As Exception
            lw.WriteLine("FrontCheck: " & ex.ToString)
        End Try

        Try
            Call Rear_check(p1_3, Back_line)
        Catch ex As Exception
            lw.WriteLine("Rear Check: " & ex.ToString)
        End Try
    End Sub

    '------------------------Front Check----------------------------------------------------
    Sub Fronto_Check(ByVal SE_pt As NXOpen.Point, ByVal Center_pt As NXOpen.Point, ByVal Axis_in As NXOpen.Curve)


        '------------------------------------------------------Front View
        'tao diem da chon
        Call Lib_NX.Create_Point(SE_pt.Coordinates.X, SE_pt.Coordinates.Y, SE_pt.Coordinates.Z, Fr_SE_Pnt)
        Fr_SE_Pnt.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Call Lib_NX.Create_Point(Center_pt.Coordinates.X, Center_pt.Coordinates.Y, Center_pt.Coordinates.Z, Fr_Center_pt)
        Fr_Center_pt.SetVisibility(SmartObject.VisibilityOption.Invisible)
        've duong ground

        'hinge axis
        Dim Axis_out As NXOpen.Curve
        Call wave_link(Axis_in, Axis_out)
        'tao datum vung mat
        Dim DaPlane As NXOpen.DatumPlane
        Call TaoMatSE(Axis_out, Fr_SE_Pnt, DaPlane)
        Dim Headzone__arc As NXOpen.Arc
        Call Head_zone(DaPlane, Axis_out, Fr_SE_Pnt, Headzone__arc)
        Dim numberHidden1 As Integer = Nothing
        numberHidden1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_DATUM_PLANES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
        workPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
        'Plan view
        Dim Fr_Center_pt1 As NXOpen.Point
        Dim Fr_SE_Pnt1 As NXOpen.Point
        Call chieu_headzone(Headzone__arc, Fr_SE_Pnt, Fr_Center_pt, Fr_SE_Pnt1, Fr_Center_pt1)

        Dim Dimen3 As Point3d = New NXOpen.Point3d(Fr_SE_Pnt1.Coordinates.X - 400, Fr_SE_Pnt1.Coordinates.Y + 300, Fr_SE_Pnt1.Coordinates.Z) '(936.13446125343967, -309.523651364787, 663.69329279106523)
        Call PMI_Horizontal_Pnt_Pnt(Fr_SE_Pnt1, Fr_Center_pt1, Nothing, Dimen3, False, Condition_IN(10), Nothing, texthigh)
        Dim point3_diemdat As NXOpen.Point3d = New NXOpen.Point3d(Fr_SE_Pnt1.Coordinates.X + 300, Fr_SE_Pnt1.Coordinates.Y + 250, Fr_SE_Pnt1.Coordinates.Z) '(936.13446125343967, -309.523651364787, 663.69329279106523)
        Call PMI_Vertical_Pnt_Pnt(Fr_SE_Pnt1, Fr_Center_pt1, Nothing, Nothing, point3_diemdat, False, 0, Condition_IN(12), Nothing, texthigh)
        Call PMI_Radial_arc(iHead_circle, help2, Nothing, Condition_IN(11))


        '------------------------------------------------------Side View
        Call side_view(Fr_Center_pt, Fr_SE_Pnt)
        Call Toado_Handle(Fr_Center_pt, True)

        '------------------------------------------------------Chup anh
        Call View("Top")

        Call zentaizu(ImgeFolder3, "FR" & ".png")
        Dim numbershow1 As Integer = Nothing
        numbershow1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_DATUM_PLANES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
        workPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.ShowOnly)
    End Sub
    Sub TaoMatSE(ByVal line1 As Line, ByVal pnt1 As NXOpen.Point, ByRef datumPlane1 As NXOpen.DatumPlane)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim pnt3d As NXOpen.Point3d = New Point3d(pnt1.Coordinates.X, pnt1.Coordinates.Y, pnt1.Coordinates.Z)
        ' ----------------------------------------------
        '   Menu: Insert->Datum->Datum Plane...
        ' ---------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim datumPlaneBuilder1 As NXOpen.Features.DatumPlaneBuilder = Nothing
        datumPlaneBuilder1 = workPart.Features.CreateDatumPlaneBuilder(nullNXOpen_Features_Feature)

        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = datumPlaneBuilder1.GetPlane()

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        plane1.SetUpdateOption(NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim expression3 As NXOpen.Expression = Nothing
        expression3 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim point2 As NXOpen.Point = pnt1

        Dim nullNXOpen_Xform As NXOpen.Xform = Nothing

        Dim point3 As NXOpen.Point = Nothing
        point3 = workPart.Points.CreatePoint(point2, nullNXOpen_Xform, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim expression4 As NXOpen.Expression = Nothing
        expression4 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim point4 As NXOpen.Point = Nothing
        point4 = workPart.Points.CreatePoint(point2, nullNXOpen_Xform, NXOpen.SmartObject.UpdateOption.WithinModeling)

        'Dim compositeCurve1 As NXOpen.Features.CompositeCurve = CType(workPart.Features.FindObject("LINKED_CURVE(1)"), NXOpen.Features.CompositeCurve)

        'Dim line1 As NXOpen.Line = CType(compositeCurve1.FindObject("CURVE 1"), NXOpen.Line)

        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(line1, NXOpen.Sense.Reverse, NXOpen.SmartObject.UpdateOption.WithinModeling)

        plane1.SetMethod(NXOpen.PlaneTypes.MethodType.PointDir)

        Dim geom1(1) As NXOpen.NXObject
        geom1(0) = point3
        geom1(1) = direction1
        plane1.SetGeometry(geom1)

        plane1.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)

        plane1.Evaluate()

        Dim coordinates2 As NXOpen.Point3d = pnt3d
        Dim point5 As NXOpen.Point = Nothing
        point5 = workPart.Points.CreatePoint(coordinates2)

        workPart.Points.DeletePoint(point1)

        Dim coordinates3 As NXOpen.Point3d = pnt3d
        Dim point6 As NXOpen.Point = Nothing
        point6 = workPart.Points.CreatePoint(coordinates3)

        workPart.Points.DeletePoint(point5)

        plane1.RemoveOffsetData()

        plane1.Evaluate()

        datumPlaneBuilder1.ResizeDuringUpdate = True

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = datumPlaneBuilder1.CommitFeature()

        Dim datumPlaneFeature1 As NXOpen.Features.DatumPlaneFeature = CType(feature1, NXOpen.Features.DatumPlaneFeature)

        'Dim datumPlane1 As NXOpen.DatumPlane = Nothing
        datumPlane1 = datumPlaneFeature1.DatumPlane

        datumPlane1.SetReverseSection(False)

        datumPlaneBuilder1.Destroy()

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression2)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression1)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression3)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression4)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.Points.DeletePoint(point4)

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub Head_zone(ByVal iplane As DatumPlane, ByVal Axis_curve As NXOpen.Curve, ByVal SE_pnt As NXOpen.Point, ByRef Headzone_ARC As NXOpen.Arc)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Sketch
        ' ----------------------------------------------
        'mo sketch
        theSession.BeginTaskEnvironment()

        ' ----------------------------------------------
        '   Menu: Application->Document->PMI
        ' ----------------------------------------------

        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin

        sketchInPlaceBuilder1.OriginOptionInfer = NXOpen.OriginMethod.WorkPartOrigin

        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)

        Dim simpleSketchInPlaceBuilder1 As NXOpen.SimpleSketchInPlaceBuilder = Nothing
        simpleSketchInPlaceBuilder1 = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()

        sketchAlongPathBuilder1.PlaneLocation.Expression.SetFormula("0")

        simpleSketchInPlaceBuilder1.UseWorkPartOrigin = True

        'Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 50.516901170581455, 1112.4685482059983)
        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(SE_pnt.Coordinates.X, SE_pnt.Coordinates.Y, SE_pnt.Coordinates.Z)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) X axis"), NXOpen.DatumAxis)

        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim datumPlane1 As NXOpen.DatumPlane = iplane 'CType(workPart.Datums.FindObject("DATUM_PLANE(2)"), NXOpen.DatumPlane)

        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(datumPlane1, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, False)

        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        simpleSketchInPlaceBuilder1.CoordinateSystem = cartesianCoordinateSystem1

        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1

        Dim nErrs1 As Integer = Nothing
        nErrs1 = theSession.UpdateManager.AddToDeleteList(point1)

        theSession.Preferences.Sketch.CreateInferredConstraints = True

        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False

        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression

        theSession.Preferences.Sketch.TextSizeFixed = True

        theSession.Preferences.Sketch.FixedTextSize = 3.0

        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True

        theSession.Preferences.Sketch.DisplayReferenceGeometry = True

        theSession.Preferences.Sketch.DisplayShadedRegions = True

        theSession.Preferences.Sketch.FindMovableObjects = True

        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0

        theSession.Preferences.Sketch.DisplayObjectColor = True

        theSession.Preferences.Sketch.DisplayObjectName = True

        theSession.Preferences.Sketch.EditDimensionOnCreation = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = simpleSketchInPlaceBuilder1.Commit()

        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketch1.Feature

        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)

        sketchInPlaceBuilder1.Destroy()

        sketchAlongPathBuilder1.Destroy()

        simpleSketchInPlaceBuilder1.Destroy()

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression2)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression1)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        plane1.DestroyPlane()

        theSession.ActiveSketch.SetName("SKETCH_000")

        theSession.CleanUpFacetedFacesAndEdges()
        ' ----------------------------------------------
        'Main
        'tim tam duong tron

        Call Laygiaodiemvoimatsketch(Axis_curve, Center_Pnt_out)
        'lw.WriteLine(Center_Pnt_out.Coordinates.ToString)

        Dim Center_Pnt3d As NXOpen.Point3d = New Point3d(Center_Pnt_out.Coordinates.X, Center_Pnt_out.Coordinates.Y, Center_Pnt_out.Coordinates.Z)

        'tim ban kinh
        Dim Center_draft() As NXOpen.Point3d
        Call Get_MinDist(Center_Pnt_out.Tag, SE_pnt.Tag, Kcach, Center_draft)

        Call veduongtron_tamvadiem(Center_Pnt_out, Kcach, Headzone_ARC)


        ' ----------------------------------------------
        theSession.Preferences.Sketch.SectionView = False
        theSession.ActiveSketch.Deactivate(NXOpen.Sketch.ViewReorient.False, NXOpen.Sketch.UpdateLevel.Model)
        'theSession.DeleteUndoMarksSetInTaskEnvironment()
        theSession.EndTaskEnvironment()



    End Sub
    Sub chieu_headzone(ByVal arc01 As NXOpen.Arc, ByVal SE_p As NXOpen.Point, ByVal Center_p As NXOpen.Point, ByRef SE_p_chieu As NXOpen.Point,
       ByRef Center_p_chieu As NXOpen.Point)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        theSession.BeginTaskEnvironment()
        ' ----------------------------------------------

        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin
        sketchInPlaceBuilder1.OriginOptionInfer = NXOpen.OriginMethod.WorkPartOrigin
        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)
        Dim simpleSketchInPlaceBuilder1 As NXOpen.SimpleSketchInPlaceBuilder = Nothing
        simpleSketchInPlaceBuilder1 = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()
        sketchAlongPathBuilder1.PlaneLocation.Expression.SetFormula("0")
        simpleSketchInPlaceBuilder1.UseWorkPartOrigin = True
        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)
        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) X axis"), NXOpen.DatumAxis)
        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim datumPlane1 As NXOpen.DatumPlane = CType(workPart.Datums.FindObject("DATUM_CSYS(0) XY plane"), NXOpen.DatumPlane)
        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(datumPlane1, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, False)
        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        simpleSketchInPlaceBuilder1.CoordinateSystem = cartesianCoordinateSystem1
        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1
        theSession.Preferences.Sketch.CreateInferredConstraints = True
        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False
        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression
        theSession.Preferences.Sketch.TextSizeFixed = True
        theSession.Preferences.Sketch.FixedTextSize = 3.0
        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True
        theSession.Preferences.Sketch.DisplayReferenceGeometry = True
        theSession.Preferences.Sketch.DisplayShadedRegions = True
        theSession.Preferences.Sketch.FindMovableObjects = True
        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0
        theSession.Preferences.Sketch.DisplayObjectColor = True
        theSession.Preferences.Sketch.DisplayObjectName = True
        theSession.Preferences.Sketch.EditDimensionOnCreation = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = simpleSketchInPlaceBuilder1.Commit()
        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)
        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketch1.Feature
        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)
        sketchInPlaceBuilder1.Destroy()
        sketchAlongPathBuilder1.Destroy()
        simpleSketchInPlaceBuilder1.Destroy()

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression2)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression1)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        plane1.DestroyPlane()
        theSession.ActiveSketch.SetName("SKETCH_001")
        theSession.CleanUpFacetedFacesAndEdges()
        ' ----------------------------------------------
        'Main
        Dim arc_chieu As NXOpen.Ellipse
        Call chieuduong(arc01, arc_chieu, feature1)
        Call Netdut(arc_chieu)
        Call chieudiem(Center_p, Center_p_chieu, feature1)
        Call chieudiem(SE_p, SE_p_chieu, feature1)
        Call HEAD(Center_p_chieu, arc_chieu)
        ' ----------------------------------------------
        theSession.Preferences.Sketch.SectionView = False

        theSession.ActiveSketch.Deactivate(NXOpen.Sketch.ViewReorient.False, NXOpen.Sketch.UpdateLevel.Model)

        'theSession.DeleteUndoMarksSetInTaskEnvironment()

        theSession.EndTaskEnvironment()
    End Sub
    Sub HEAD(ByVal center_p2 As NXOpen.Point, ByVal iBig_circle As NXOpen.Ellipse)
        Dim Head_pnt As NXOpen.Point
        Call Lib_NX.Create_Point(center_p2.Coordinates.X + 350, center_p2.Coordinates.Y - 300, center_p2.Coordinates.Z, Head_pnt)
        Head_pnt.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Call veduongtron_tamvadiem(Head_pnt, 5, iHead_circle)
        Dim drf_p As NXOpen.Point3d = New Point3d(Center_Pnt_out.Coordinates.X, Center_Pnt_out.Coordinates.Y, 0)
        Dim center_p1 As NXOpen.Point = workPart.Points.CreatePoint(drf_p)

        Call Tim_helppoint(center_p1, Head_pnt, iBig_circle, iHead_circle, help1, help2)

        Call MakeTiepTuyen(iBig_circle, iHead_circle, help1, help2)

        line2.SetVisibility(SmartObject.VisibilityOption.Invisible)
    End Sub
    Sub Tim_helppoint(ByVal ipoint1 As NXOpen.Point, ByVal ipoint2 As NXOpen.Point, ByVal Arc1 As NXOpen.Ellipse, ByVal Arc2 As NXOpen.Arc, ByRef p1 As NXOpen.Point3d, ByRef p2 As NXOpen.Point3d)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        line2 = workPart.Curves.CreateLine(ipoint1, ipoint2)
        line2.SetVisibility(SmartObject.VisibilityOption.Visible)

        Dim iip1() As NXOpen.Point3d
        Dim iip2() As NXOpen.Point3d
        Call Get_MinDist(line2.Tag, Arc1.Tag, Kcach, iip1)
        p1 = iip1(0)
        Call Get_MinDist(line2.Tag, Arc2.Tag, Kcach, iip2)
        p2 = iip2(0)
    End Sub
    Sub side_view(ByVal Handle_center As NXOpen.Point, ByVal FR_SE As NXOpen.Point)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        theSession.BeginTaskEnvironment()
        ' ----------------------------------------------
        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin
        sketchInPlaceBuilder1.OriginOptionInfer = NXOpen.OriginMethod.WorkPartOrigin
        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)
        Dim simpleSketchInPlaceBuilder1 As NXOpen.SimpleSketchInPlaceBuilder = Nothing
        simpleSketchInPlaceBuilder1 = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()
        sketchAlongPathBuilder1.PlaneLocation.Expression.SetFormula("0")
        simpleSketchInPlaceBuilder1.UseWorkPartOrigin = True
        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) X axis"), NXOpen.DatumAxis)
        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim datumPlane1 As NXOpen.DatumPlane = CType(workPart.Datums.FindObject("DATUM_CSYS(0) XZ plane"), NXOpen.DatumPlane)
        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(datumPlane1, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, True)

        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        simpleSketchInPlaceBuilder1.CoordinateSystem = cartesianCoordinateSystem1
        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1
        theSession.Preferences.Sketch.CreateInferredConstraints = True
        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False
        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression
        theSession.Preferences.Sketch.TextSizeFixed = True
        theSession.Preferences.Sketch.FixedTextSize = 3.0
        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True
        theSession.Preferences.Sketch.DisplayReferenceGeometry = True
        theSession.Preferences.Sketch.DisplayShadedRegions = True
        theSession.Preferences.Sketch.FindMovableObjects = True
        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0
        theSession.Preferences.Sketch.DisplayObjectColor = True
        theSession.Preferences.Sketch.DisplayObjectName = True
        theSession.Preferences.Sketch.EditDimensionOnCreation = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = simpleSketchInPlaceBuilder1.Commit()

        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketch1.Feature
        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)

        sketchInPlaceBuilder1.Destroy()
        sketchAlongPathBuilder1.Destroy()
        simpleSketchInPlaceBuilder1.Destroy()

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression2)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression1)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        plane1.DestroyPlane()
        theSession.ActiveSketch.SetName("SKETCH_002")
        theSession.CleanUpFacetedFacesAndEdges()
        ' ----------------------------------------------
        've ground line
        Dim start_Point As NXOpen.Point = workPart.Points.CreatePoint(Start_Grd_pnt)
        theSession.ActiveSketch.AddGeometry(start_Point, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        Dim end_Point As NXOpen.Point = workPart.Points.CreatePoint(End_Grd_pnt)
        theSession.ActiveSketch.AddGeometry(end_Point, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
        Call rangbuocdiem(start_Point, -1 * Start_Grd_pnt.Z, True, "p0")
        Call rangbuocdiem(start_Point, -1 * Start_Grd_pnt.X, False, "p1")
        Call rangbuocdiem(end_Point, -1 * End_Grd_pnt.Z, True, "p2")
        Call rangbuocdiem(end_Point, End_Grd_pnt.X, False, "p3")
        Call veduongtrongsketch(start_Point, end_Point, Ground_line)
        start_Point.SetVisibility(SmartObject.VisibilityOption.Invisible)
        end_Point.SetVisibility(SmartObject.VisibilityOption.Invisible)



        Dim Sideview_FR_SE As NXOpen.Point
        Call chieudiem(FR_SE, Sideview_FR_SE, feature1)
        Dim dimen3 As NXOpen.Point3d = New NXOpen.Point3d(Sideview_FR_SE.Coordinates.X + 250, Sideview_FR_SE.Coordinates.Y, Sideview_FR_SE.Coordinates.Z - 500) '(1639.3594590296225, 0.0, 269.8119400167011)
        Call PMI_vertical_Pnt_Line(Sideview_FR_SE, Ground_line, Nothing, dimen3, True, 0, Condition_IN(9), Nothing, texthigh)

        Dim Sideview_Handle_pnt As NXOpen.Point
        Call chieudiem(Handle_center, Sideview_Handle_pnt, feature1)
        Dim dimen2 As NXOpen.Point3d = New NXOpen.Point3d(Sideview_Handle_pnt.Coordinates.X - 250, Sideview_Handle_pnt.Coordinates.Y, Sideview_Handle_pnt.Coordinates.Z - 100) '(1639.3594590296225, 0.0, 269.8119400167011)
        Call PMI_vertical_Pnt_Line(Sideview_Handle_pnt, Ground_line, Nothing, dimen2, False, Condition_IN(4), Condition_IN(5), Nothing, texthigh)


        ' ----------------------------------------------
        theSession.Preferences.Sketch.SectionView = False
        theSession.ActiveSketch.Deactivate(NXOpen.Sketch.ViewReorient.False, NXOpen.Sketch.UpdateLevel.Model)
        'theSession.DeleteUndoMarksSetInTaskEnvironment()
        theSession.EndTaskEnvironment()



        Dim dimen1 As NXOpen.Point3d = New NXOpen.Point3d(Sideview_FR_SE.Coordinates.X + 150, 0, Sideview_FR_SE.Coordinates.Z + 100)
        Call PMI_Horizontal_Pnt_Pnt(Sideview_FR_SE, Sideview_Handle_pnt, "Y", dimen1, True, Condition_IN(8), Nothing, texthigh)
    End Sub
    Sub Set_ground(ByVal line1 As NXOpen.Curve, ByVal point_In As NXOpen.Point, ByVal ivetical As Boolean)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim pointln As Point3d = New Point3d(point_In.Coordinates.X, point_In.Coordinates.Y, point_In.Coordinates.Z)
        ' ----------------------------------------------
        '   Menu: Insert->Dimensions->Linear...
        ' ----------------------------------------------
        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing

        Dim sketchLinearDimensionBuilder1 As NXOpen.SketchLinearDimensionBuilder = Nothing
        sketchLinearDimensionBuilder1 = workPart.Sketches.CreateLinearDimensionBuilder(nullNXOpen_Annotations_Dimension)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)
        If ivetical = False Then
            sketchLinearDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Horizontal
        Else
            sketchLinearDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Vertical
        End If

        sketchLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchLinearDimensionBuilder1.Driving.DrivingMethod = NXOpen.Annotations.DrivingValueBuilder.DrivingValueMethod.Driving

        sketchLinearDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing

        sketchLinearDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing

        sketchLinearDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View

        sketchLinearDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        sketchLinearDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Vertical

        sketchLinearDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        sketchLinearDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim datumCsys1 As NXOpen.Features.DatumCsys = CType(workPart.Features.FindObject("DATUM_CSYS(0)"), NXOpen.Features.DatumCsys)

        Dim point1 As NXOpen.Point = CType(datumCsys1.FindObject("POINT 1"), NXOpen.Point)

        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchLinearDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim dimensionlinearunits21 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits21 = sketchLinearDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim dimensionlinearunits22 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits22 = sketchLinearDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        'Dim line1 As NXOpen.Line = CType(theSession.ActiveSketch.FindObject("Curve Line1"), NXOpen.Line)

        Dim point1_3 As NXOpen.Point3d = pointln
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Start, line1, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = pointln
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Start, line1, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim point1_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchLinearDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_5, Nothing, nullNXOpen_View, point2_5)

        Dim point1_6 As NXOpen.Point3d = pointln
        Dim point2_6 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Start, line1, displayPart.ModelingViews.WorkView, point1_6, Nothing, nullNXOpen_View, point2_6)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometryFromLeader(True)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.RelativeToGeometry
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = displayPart.ModelingViews.WorkView
        Dim point2 As NXOpen.Point = CType(workPart.Points.FindObject("ENTITY 2 2"), NXOpen.Point)

        assocOrigin1.PointOnGeometry = point2
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        sketchLinearDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d = New NXOpen.Point3d(-111.78754067552983, 0.0, -65.051474752499544)
        sketchLinearDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)

        sketchLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchLinearDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Left

        sketchLinearDimensionBuilder1.Style.DimensionStyle.TextCentered = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchLinearDimensionBuilder1.Commit()
    End Sub

    '------------------------Rear Check----------------------------------------------------
    Sub Rear_check(ByVal Selected_Pnt As NXOpen.Point, ByVal Back_Curve_in As NXOpen.Curve)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        ' ----------------------------------------------
        Dim Back_Curve_out As NXOpen.Curve
        Call wave_link(Back_Curve_in, Back_Curve_out)

        ' ----------------------------------------------
        theSession.BeginTaskEnvironment()
        ' ----------------------------------------------
        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin
        sketchInPlaceBuilder1.OriginOptionInfer = NXOpen.OriginMethod.WorkPartOrigin
        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)
        Dim simpleSketchInPlaceBuilder1 As NXOpen.SimpleSketchInPlaceBuilder = Nothing
        simpleSketchInPlaceBuilder1 = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()
        sketchAlongPathBuilder1.PlaneLocation.Expression.SetFormula("0")
        simpleSketchInPlaceBuilder1.UseWorkPartOrigin = True
        Dim coordinates1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) X axis"), NXOpen.DatumAxis)
        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim datumPlane1 As NXOpen.DatumPlane = CType(workPart.Datums.FindObject("DATUM_CSYS(0) XZ plane"), NXOpen.DatumPlane)
        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(datumPlane1, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, True)

        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        simpleSketchInPlaceBuilder1.CoordinateSystem = cartesianCoordinateSystem1
        simpleSketchInPlaceBuilder1.HorizontalReference.Value = datumAxis1
        theSession.Preferences.Sketch.CreateInferredConstraints = True
        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False
        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression
        theSession.Preferences.Sketch.TextSizeFixed = True
        theSession.Preferences.Sketch.FixedTextSize = 3.0
        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True
        theSession.Preferences.Sketch.DisplayReferenceGeometry = True
        theSession.Preferences.Sketch.DisplayShadedRegions = True
        theSession.Preferences.Sketch.FindMovableObjects = True
        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0
        theSession.Preferences.Sketch.DisplayObjectColor = True
        theSession.Preferences.Sketch.DisplayObjectName = True
        theSession.Preferences.Sketch.EditDimensionOnCreation = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = simpleSketchInPlaceBuilder1.Commit()

        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketch1.Feature
        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)

        sketchInPlaceBuilder1.Destroy()
        sketchAlongPathBuilder1.Destroy()
        simpleSketchInPlaceBuilder1.Destroy()

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression2)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        Try
            ' Expression is still in use.
            workPart.Expressions.Delete(expression1)
        Catch ex As NXException
            ex.AssertErrorCode(1050029)
        End Try

        plane1.DestroyPlane()
        theSession.ActiveSketch.SetName("SKETCH_003")
        theSession.CleanUpFacetedFacesAndEdges()
        ' ----------------------------------------------
        Dim RR_center_pnt As NXOpen.Point
        Call Lib_NX.Create_Point(Selected_Pnt.Coordinates.X, Selected_Pnt.Coordinates.Y, Selected_Pnt.Coordinates.Z, RR_center_pnt)
        RR_center_pnt.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Dim RR_center_dim As NXOpen.Point
        Call chieudiem(RR_center_pnt, RR_center_dim, feature1)

        Dim curve_chieu As NXOpen.Curve
        Call chieuduong(Back_Curve_out, curve_chieu, feature1)

        Dim startPoint1 As NXOpen.Point3d = New Point3d(RR_center_dim.Coordinates.X + 1000, 0, RR_center_dim.Coordinates.Z + 500)
        Dim endPoint1 As NXOpen.Point3d = New Point3d(RR_center_dim.Coordinates.X + 1000, 0, 0)
        Dim idraft_line As NXOpen.Line = Nothing
        idraft_line = workPart.Curves.CreateLine(startPoint1, endPoint1)
        theSession.ActiveSketch.AddGeometry(idraft_line, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)

        Call Make_Vuonggoc(idraft_line, Ground_line)
        Dim iBackPnt() As Point3d
        Dim iBack_point As NXOpen.Point
        Dim iHelppnt1 As NXOpen.Point3d
        Dim iHelppnt2 As NXOpen.Point3d
        iHelppnt1 = timdiemdau("Z", True, idraft_line)
        iHelppnt2 = timdiemdau("Z", True, curve_chieu)
        Call MakeTiepTuyen(curve_chieu, idraft_line, iHelppnt1, iHelppnt2)
        Call Get_MinDist(curve_chieu.Tag, idraft_line.Tag, Kcach, iBackPnt)
        Call Lib_NX.Create_Point(iBackPnt(0).X, iBackPnt(0).Y, iBackPnt(0).Z, iBack_point)
        iBack_point.SetVisibility(SmartObject.VisibilityOption.Invisible)
        Call Netdut(idraft_line)


        ' ----------------------------------------------
        theSession.Preferences.Sketch.SectionView = False
        theSession.ActiveSketch.Deactivate(NXOpen.Sketch.ViewReorient.False, NXOpen.Sketch.UpdateLevel.Model)
        'theSession.DeleteUndoMarksSetInTaskEnvironment()
        theSession.EndTaskEnvironment()


        Dim dimen4 As NXOpen.Point3d = New Point3d(RR_center_dim.Coordinates.X - 50, 0, RR_center_dim.Coordinates.Z - 200)
        Call PMI_vertical_Pnt_Line(RR_center_dim, Ground_line, Nothing, dimen4, False, Condition_IN(6), Condition_IN(7), Nothing, texthigh)
        Dim dimen5 As NXOpen.Point3d = New Point3d(RR_center_dim.Coordinates.X, 0, RR_center_dim.Coordinates.Z + 200)
        Call PMI_Horizontal_Pnt_Pnt(RR_center_dim, iBack_point, "Y", dimen5, True, Condition_IN(10), Nothing, texthigh)
        Call Toado_Handle(RR_center_pnt, False)

        '------------------------------------------------------Chup anh
        Call View("Left")
        Call zentaizu(ImgeFolder3, "RR" & ".png")
    End Sub
    Sub Toado_Handle(ByVal Handle_pnt As NXOpen.Point, ByVal iLoc_FR As Boolean)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim drf_pnt As Point3d = New Point3d(Handle_pnt.Coordinates.X, Handle_pnt.Coordinates.Y, Handle_pnt.Coordinates.Z)

        ' ----------------------------------------------
        '   Menu: PMI->Specialized->Coordinate Note...
        ' ----------------------------------------------

        Dim nullNXOpen_Annotations_CoordinateNote As NXOpen.Annotations.CoordinateNote = Nothing

        Dim coordinateNoteBuilder1 As NXOpen.Annotations.CoordinateNoteBuilder = Nothing
        coordinateNoteBuilder1 = workPart.PmiManager.PmiAttributes.CreateCoordinateNoteBuilder(nullNXOpen_Annotations_CoordinateNote)
        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)
        coordinateNoteBuilder1.Title = ""
        coordinateNoteBuilder1.Category = ""
        coordinateNoteBuilder1.Identifier = ""
        coordinateNoteBuilder1.Revision = ""
        coordinateNoteBuilder1.ToggleX = True
        coordinateNoteBuilder1.StringPrefixX = "X="
        coordinateNoteBuilder1.ToggleY = True
        coordinateNoteBuilder1.StringPrefixY = "Y="
        coordinateNoteBuilder1.ToggleZ = True
        coordinateNoteBuilder1.StringPrefixZ = "Z="
        coordinateNoteBuilder1.StringPrefixI = ""
        coordinateNoteBuilder1.StringPrefixJ = ""
        coordinateNoteBuilder1.StringPrefixK = ""
        coordinateNoteBuilder1.StringPrefixLabel = ""
        coordinateNoteBuilder1.StringPrefixLevel = ""
        coordinateNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim leaderData1 As NXOpen.Annotations.LeaderData = Nothing
        leaderData1 = workPart.Annotations.CreateLeaderData()
        leaderData1.StubSize = 2.0
        leaderData1.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.OpenArrow
        leaderData1.VerticalAttachment = NXOpen.Annotations.LeaderVerticalAttachment.Center
        coordinateNoteBuilder1.Leader.Leaders.Append(leaderData1)
        coordinateNoteBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        coordinateNoteBuilder1.Title = "Coordinate Note"
        coordinateNoteBuilder1.Category = "User Defined"
        coordinateNoteBuilder1.Identifier = "User Defined"
        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)
        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        coordinateNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane

        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim point1 As NXOpen.Point = Handle_pnt

        Dim nullNXOpen_Xform As NXOpen.Xform = Nothing

        Dim point2 As NXOpen.Point = Nothing
        point2 = workPart.Points.CreatePoint(point1, nullNXOpen_Xform, NXOpen.SmartObject.UpdateOption.AfterModeling)

        coordinateNoteBuilder1.TrackingPoint = point2

        Dim nullNXOpen_View As NXOpen.View = Nothing

        Dim point1_1 As NXOpen.Point3d = drf_pnt
        Dim point2_1 As NXOpen.Point3d = drf_pnt
        leaderData1.Leader.SetValue(NXOpen.InferSnapType.SnapType.None, point2, nullNXOpen_View, point1_1, Nothing, nullNXOpen_View, point2_1)

        coordinateNoteBuilder1.Style.LetteringStyle.GeneralTextSize = texthigh
        '----------------------------------
        coordinateNoteBuilder1.DecimalPlace = 1
        '----------------------------------
        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        coordinateNoteBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d
        If iLoc_FR = True Then
            point3 = New NXOpen.Point3d(drf_pnt.X - 500, drf_pnt.Y, drf_pnt.Z + 250) '(1363.7334658618661, -861.0769381313745, 444.20932637597065)
            leaderData1.StubSide = NXOpen.Annotations.LeaderSide.Right
        Else
            point3 = New NXOpen.Point3d(drf_pnt.X - 500, drf_pnt.Y, drf_pnt.Z + 250) '(1363.7334658618661, -861.0769381313745, 444.20932637597065)
            leaderData1.StubSide = NXOpen.Annotations.LeaderSide.Right
        End If
        coordinateNoteBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)

        coordinateNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = coordinateNoteBuilder1.Commit()

        Dim iPmi As NXOpen.Annotations.Annotation = CType(nXObject1, NXOpen.Annotations.Annotation)
        Call PMI_doidoday(iPmi)
        coordinateNoteBuilder1.Destroy()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression1)

        workPart.MeasureManager.ClearPartTransientModification()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
#End Region

#Region "Cong cu"
    '------------------------Sub tinh toan------------------------
    Function timdiemdau(ByVal Dir As String, ByVal lower As Boolean, ByVal Line1 As Curve) As Point3d
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUfSession As UFSession = UFSession.GetUFSession()
        Dim junk(2) As Double, trash As Double
        Dim EndPt(2) As Double
        Dim StrPt(2) As Double
        theUfSession.Modl.AskCurveProps(Line1.Tag, 1, EndPt, junk, junk, junk, trash, trash)
        theUfSession.Modl.AskCurveProps(Line1.Tag, 0, StrPt, junk, junk, junk, trash, trash)

        Dim StrPoint As Point3d = New Point3d(StrPt(0), StrPt(1), StrPt(2))
        Dim EndPoint As Point3d = New Point3d(EndPt(0), EndPt(1), EndPt(2))
        Dim tempval As Double
        If Dir = "Z" Then
            tempval = 2
        ElseIf Dir = "Y" Then
            tempval = 1
        ElseIf Dir = "X" Then
            tempval = 0
        End If

        If lower = True Then
            If StrPt(tempval) < EndPt(tempval) Then
                timdiemdau = StrPoint
            Else
                timdiemdau = EndPoint
            End If
        Else
            If StrPt(tempval) > EndPt(tempval) Then
                timdiemdau = StrPoint
            Else
                timdiemdau = EndPoint
            End If
        End If

        'iVector = New Vector3d(StrPoint.X - EndPoint.X, StrPoint.Y - EndPoint.Y, StrPoint.Z - EndPoint.Z)
    End Function

    Function TIMDIEMTHUOCDUONGTRON(ByVal Line1 As Curve) As Point3d
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUfSession As UFSession = UFSession.GetUFSession()

        Dim junk(2) As Double, trash As Double
        Dim EndPt(2) As Double
        Dim StrPt(2) As Double
        Dim Midpt(2) As Double
        theUfSession.Modl.AskCurveProps(Line1.Tag, 0.5, Midpt, junk, junk, junk, trash, trash)
        Dim MidPoint As Point3d = New Point3d(Midpt(0), Midpt(1), Midpt(2))
        TIMDIEMTHUOCDUONGTRON = MidPoint
    End Function
    Function timdiemdauKhacgoc(ByVal Line1 As Curve, ByVal Tmppt As Point3d) As Point3d
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUfSession As UFSession = UFSession.GetUFSession()

        Dim junk(2) As Double, trash As Double
        Dim EndPt(2) As Double
        Dim StrPt(2) As Double
        theUfSession.Modl.AskCurveProps(Line1.Tag, 1, EndPt, junk, junk, junk, trash, trash)
        theUfSession.Modl.AskCurveProps(Line1.Tag, 0, StrPt, junk, junk, junk, trash, trash)

        Dim StrPoint As Point3d = New Point3d(StrPt(0), StrPt(1), StrPt(2))
        Dim EndPoint As Point3d = New Point3d(EndPt(0), EndPt(1), EndPt(2))
        If System.Math.Abs(StrPt(2) - Tmppt.Z) > 0.1 Then
            timdiemdauKhacgoc = StrPoint
        Else
            timdiemdauKhacgoc = EndPoint
        End If
    End Function
    Sub tinhvecto(ByVal Line1 As Curve, ByRef iVector As NXOpen.Vector3d)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUfSession As UFSession = UFSession.GetUFSession()

        Dim junk(2) As Double, trash As Double
        Dim EndPt(2) As Double
        Dim StrPt(2) As Double
        theUfSession.Modl.AskCurveProps(Line1.Tag, 1, EndPt, junk, junk, junk, trash, trash)
        theUfSession.Modl.AskCurveProps(Line1.Tag, 0, StrPt, junk, junk, junk, trash, trash)

        Dim StrPoint As Point3d = New Point3d(StrPt(0), StrPt(1), StrPt(2))
        Dim EndPoint As Point3d = New Point3d(EndPt(0), EndPt(1), EndPt(2))

        iVector = New Vector3d(StrPoint.X - EndPoint.X, StrPoint.Y - EndPoint.Y, StrPoint.Z - EndPoint.Z)
    End Sub
    Public Sub Get_MinDist(ByVal Obj1 As NXOpen.Tag, ByVal Obj2 As NXOpen.Tag, ByRef minDist As Double, ByRef PtArr() As NXOpen.Point3d)
        Dim guess1(2) As Double
        Dim guess2(2) As Double
        Dim pt1(2) As Double
        Dim pt2(2) As Double

        theUFSession.Modl.AskMinimumDist(Obj1, Obj2, 0, guess1, 0, guess2, minDist, pt1, pt2)
        ReDim Preserve PtArr(0)
        PtArr(0) = New NXOpen.Point3d(pt1(0), pt1(1), pt1(2))

        ReDim Preserve PtArr(1)
        PtArr(1) = New NXOpen.Point3d(pt2(0), pt2(1), pt2(2))
    End Sub

    '------------------------rang buoc trong sketch------------------------------
    Sub Make_Vuonggoc(ByVal curve_1 As NXOpen.Curve, ByVal Curve_2 As NXOpen.Curve)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Geometric Constraints...
        ' ----------------------------------------------
        Dim sketchConstraintBuilder1 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder1 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.PerpendicularToString

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Coincident

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Tangent

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Perpendicular

        'Dim sketchFeature1 As NXOpen.Features.SketchFeature = CType(workPart.Features.FindObject("SKETCH(5)"), NXOpen.Features.SketchFeature)

        'Dim sketch1 As NXOpen.Sketch = CType(sketchFeature1.FindObject("SKETCH_002"), NXOpen.Sketch)

        Dim line1 As NXOpen.Line = curve_1

        Dim point1 As NXOpen.Point3d = timdiemdau("Z", True, curve_1)
        Dim added1 As Boolean = Nothing
        added1 = sketchConstraintBuilder1.GeometryToConstrain.Add(line1, displayPart.ModelingViews.WorkView, point1)

        Dim line2 As NXOpen.Line = Curve_2

        Dim point2 As NXOpen.Point3d = timdiemdau("Z", True, Curve_2)
        sketchConstraintBuilder1.GeometryToConstrainTo.SetValue(line2, displayPart.ModelingViews.WorkView, point2)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchConstraintBuilder1.Commit()

        Dim objects1() As NXOpen.NXObject
        objects1 = sketchConstraintBuilder1.GetCommittedObjects()

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.GeometryToConstrainTo.Value = Nothing

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.Destroy()

        Dim sketchConstraintBuilder2 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder2 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder2.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Perpendicular

        sketchConstraintBuilder2.Destroy()

        Dim sketchConstraintBuilder3 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder3 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.PerpendicularToString

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Coincident

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Perpendicular

        sketchConstraintBuilder3.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub rangbuocdiem(ByVal Point_0 As NXOpen.Point, ByVal giatri As Double, ByVal ivertical As Double, ByVal px As String)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Dimensions->Rapid...
        ' ----------------------------------------------
        Dim sketchRapidDimensionBuilder1 As NXOpen.SketchRapidDimensionBuilder = Nothing
        sketchRapidDimensionBuilder1 = workPart.Sketches.CreateRapidDimensionBuilder()
        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        If ivertical = True Then
            sketchRapidDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Vertical
        Else
            sketchRapidDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Horizontal
        End If

        sketchRapidDimensionBuilder1.Driving.ExpressionMode = NXOpen.Annotations.DrivingValueBuilder.DrivingExpressionMode.KeepExpression
        sketchRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane
        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)
        sketchRapidDimensionBuilder1.Driving.DrivingMethod = NXOpen.Annotations.DrivingValueBuilder.DrivingValueMethod.Driving
        sketchRapidDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        sketchRapidDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing
        sketchRapidDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View
        sketchRapidDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        Dim datumCsys1 As NXOpen.Features.DatumCsys '= CType(workPart.Features.FindObject("SKETCH(5:1B)"), NXOpen.Features.DatumCsys)

        Dim iAllFeats As FeatureCollection = workPart.Features
        For Each iFeat1 As Feature In iAllFeats
            If TypeOf (iFeat1) Is NXOpen.Features.DatumCsys Then
                datumCsys1 = iFeat1
            End If
        Next


        Dim point1 As NXOpen.Point = CType(datumCsys1.FindObject("POINT 1"), NXOpen.Point)

        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim point2 As NXOpen.Point = Point_0 'CType(theSession.ActiveSketch.FindObject("Point Point8"), NXOpen.Point)
        Dim Point0_3d As Point3d = New Point3d(Point_0.Coordinates.X, Point_0.Coordinates.Y, Point_0.Coordinates.Z)

        Dim point1_3 As NXOpen.Point3d = Point0_3d
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = Point0_3d
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim point1_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_5, Nothing, nullNXOpen_View, point2_5)

        Dim point1_6 As NXOpen.Point3d = Point0_3d
        Dim point2_6 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point2, displayPart.ModelingViews.WorkView, point1_6, Nothing, nullNXOpen_View, point2_6)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        sketchRapidDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d = Point0_3d
        sketchRapidDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchRapidDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Right

        sketchRapidDimensionBuilder1.Style.DimensionStyle.TextCentered = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchRapidDimensionBuilder1.Commit()

        sketchRapidDimensionBuilder1.Destroy()

        Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject(px), NXOpen.Expression)

        expression1.SetFormula(giatri)

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub MakeTiepTuyen(ByVal ellipse1 As NXOpen.Curve, ByVal arc1 As NXOpen.Curve, ByVal point1 As NXOpen.Point3d, ByVal point2 As NXOpen.Point3d) 'ByVal ellipse1 As NXOpen.Ellipse, ByVal arc1 As NXOpen.Arc,
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Geometric Constraints...
        ' ----------------------------------------------
        Dim sketchConstraintBuilder1 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder1 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.PerpendicularToString

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Coincident

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Tangent

        'Dim arc1 As NXOpen.Arc = CType(theSession.ActiveSketch.FindObject("Curve Arc2"), NXOpen.Arc)

        'Dim point1 As NXOpen.Point3d = New NXOpen.Point3d(1605.1346262335669, -1157.4326411893132, 0.0)
        If point1.X = 0 And point1.Y = 0 And point1.Z = 0 Then
            point1 = timdiemdau("Z", True, ellipse1)
        End If

        Dim added1 As Boolean = Nothing
        added1 = sketchConstraintBuilder1.GeometryToConstrain.Add(arc1, displayPart.ModelingViews.WorkView, point1)

        'Dim projectCurve1 As NXOpen.Features.ProjectCurve = CType(workPart.Features.FindObject("SKETCH(4:2B)"), NXOpen.Features.ProjectCurve)

        'Dim ellipse1 As NXOpen.Ellipse = CType(projectCurve1.FindObject("CURVE 1"), NXOpen.Ellipse)

        If point2.X = 0 And point2.Y = 0 And point2.Z = 0 Then
            point2 = timdiemdau("Z", True, arc1)
        End If
        'Dim point2 As NXOpen.Point3d = New NXOpen.Point3d(1531.3788569490487, -1115.0156206686891, 0.0)
        sketchConstraintBuilder1.GeometryToConstrainTo.SetValue(ellipse1, displayPart.ModelingViews.WorkView, point2)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchConstraintBuilder1.Commit()

        Dim objects1() As NXOpen.NXObject
        objects1 = sketchConstraintBuilder1.GetCommittedObjects()

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.GeometryToConstrainTo.Value = Nothing

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.Destroy()

        Dim sketchConstraintBuilder2 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder2 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder2.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Tangent

        sketchConstraintBuilder2.Destroy()

        Dim sketchConstraintBuilder3 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder3 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.PerpendicularToString

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Coincident

        sketchConstraintBuilder3.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Tangent

        sketchConstraintBuilder3.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub ttcurvevsarc(ByVal icurvein As Curve, ByVal iarcin As Arc)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        ' ----------------------------------------------
        '   Menu: Insert->Geometric Constraints...
        ' ----------------------------------------------
        Dim sketchConstraintBuilder1 As NXOpen.SketchConstraintBuilder = Nothing
        sketchConstraintBuilder1 = workPart.Sketches.CreateConstraintBuilder()

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.PerpendicularToString

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Coincident

        sketchConstraintBuilder1.ConstraintType = NXOpen.SketchConstraintBuilder.Constraint.Tangent

        Dim line1 As NXOpen.Curve = icurvein
        Dim arc1 As NXOpen.Arc = iarcin
        Dim PtAr() As Point3d
        Lib_NX.Get_MinDist(line1.Tag, arc1.Tag, Nothing, PtAr)

        Dim point1 As NXOpen.Point3d = PtAr(0)
        Dim added1 As Boolean = Nothing
        added1 = sketchConstraintBuilder1.GeometryToConstrain.Add(line1, displayPart.ModelingViews.WorkView, point1)

        Dim point2 As NXOpen.Point3d = PtAr(1)
        sketchConstraintBuilder1.GeometryToConstrainTo.SetValue(arc1, displayPart.ModelingViews.WorkView, point2)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchConstraintBuilder1.Commit()

        Dim objects1() As NXOpen.NXObject
        objects1 = sketchConstraintBuilder1.GetCommittedObjects()

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.GeometryToConstrainTo.Value = Nothing

        sketchConstraintBuilder1.GeometryToConstrain.Clear()

        sketchConstraintBuilder1.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub Rangbuocduongkinh(ByVal duongkinh As Double, ByVal arc1 As Arc, ByVal thutu As Integer)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Dimensions->Rapid...
        ' ----------------------------------------------
        Dim sketchRapidDimensionBuilder1 As NXOpen.SketchRapidDimensionBuilder = Nothing
        sketchRapidDimensionBuilder1 = workPart.Sketches.CreateRapidDimensionBuilder()

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchRapidDimensionBuilder1.Driving.ExpressionMode = NXOpen.Annotations.DrivingValueBuilder.DrivingExpressionMode.KeepExpression

        sketchRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchRapidDimensionBuilder1.Driving.DrivingMethod = NXOpen.Annotations.DrivingValueBuilder.DrivingValueMethod.Driving

        sketchRapidDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing

        sketchRapidDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing

        sketchRapidDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View

        sketchRapidDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        'Dim arc1 As NXOpen.Arc = CType(theSession.ActiveSketch.FindObject("Curve Arc1"), NXOpen.Arc)

        Dim point1_1 As NXOpen.Point3d = TIMDIEMTHUOCDUONGTRON(arc1) 'New NXOpen.Point3d(1238.1752789026423, -821.5980990436741, 818.20637703814532)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Center, arc1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = point1_1 'New NXOpen.Point3d(1238.1752789026423, -821.5980990436741, 818.20637703814532)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Center, arc1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim point1_3 As NXOpen.Point3d = point1_1 'New NXOpen.Point3d(1238.1752789026423, -821.5980990436741, 818.20637703814532)
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, arc1, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        sketchRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, Nothing, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        sketchRapidDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point1 As NXOpen.Point3d = point1_1 'New NXOpen.Point3d(1203.6733748228967, -794.55601146038771, 810.96642741886637)
        sketchRapidDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point1)

        sketchRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)

        sketchRapidDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Left

        sketchRapidDimensionBuilder1.Style.DimensionStyle.TextCentered = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchRapidDimensionBuilder1.Commit()

        sketchRapidDimensionBuilder1.Destroy()
        Exit Sub
        Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p" & thutu), NXOpen.Expression)

        expression1.SetFormula(duongkinh)

        theSession.ActiveSketch.LocalUpdate()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub

    '------------------------PMI------------------------------
    Sub PMI_Horizontal_Pnt_Pnt(ByVal point_1 As NXOpen.Point, ByVal point_2 As NXOpen.Point, ByVal Plane_dir As String, ByVal diemdatPMI As NXOpen.Point3d, ByVal text_Center As Boolean, ByVal dieukien As Double, ByVal doimau As String, ByVal text_size As Double)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing
        Dim pmiRapidDimensionBuilder1 As NXOpen.Annotations.PmiRapidDimensionBuilder = Nothing
        pmiRapidDimensionBuilder1 = workPart.Dimensions.CreatePmiRapidDimensionBuilder(nullNXOpen_Annotations_Dimension)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Horizontal

        If Plane_dir = "Y" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane
            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "Z" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "X" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane

            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        Else
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If

        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        pmiRapidDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction
        Dim nullNXOpen_View As NXOpen.View = Nothing
        pmiRapidDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View

        pmiRapidDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        Dim point1 As NXOpen.Point = point_1 'CType(workPart.Points.FindObject("ENTITY 2 6 1"), NXOpen.Point)
        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing
        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)
        Dim point2 As NXOpen.Point = point_2 'CType(workPart.Points.FindObject("ENTITY 2 7 1"), NXOpen.Point)
        Dim point1_3 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim point1_5 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_5, Nothing, nullNXOpen_View, point2_5)

        Dim point1_6 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_6 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point2, displayPart.ModelingViews.WorkView, point1_6, Nothing, nullNXOpen_View, point2_6)

        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects2(1) As NXOpen.NXObject
        objects2(0) = point2
        objects2(1) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiRapidDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d = diemdatPMI 'New NXOpen.Point3d(point_1.Coordinates.X - 200, point_1.Coordinates.Y + 300, point_1.Coordinates.Z) '(936.13446125343967, -309.523651364787, 663.69329279106523)
        pmiRapidDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Right
        pmiRapidDimensionBuilder1.Style.DimensionStyle.TextCentered = text_Center
        pmiRapidDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = text_size
        pmiRapidDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True

        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiRapidDimensionBuilder1.Commit()

        pmiRapidDimensionBuilder1.Destroy()

        '-------------------------------------------------------------------------------------------------------------------------------
        Dim iPmi As NXOpen.Annotations.PmiHorizontalDimension = CType(nXObject1, NXOpen.Annotations.PmiHorizontalDimension)
        Dim ii() As String
        Dim ikhoangcach As String
        iPmi.GetDimensionText(ii, Nothing)
        For Each istring As String In ii
            ikhoangcach = CType(istring, Double)
        Next
        Call PMI_doidoday(iPmi)
        '-----------------------------------------------------------------------
        'doi mau
        Dim icolor As Double
        If ikhoangcach < dieukien Then
            icolor = 108
        Else
            icolor = 186
        End If

        If doimau = "OK" Then
            icolor = 108
        ElseIf doimau = "NG" Then
            icolor = 186
        End If

        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewColor = icolor

        Dim objects10(0) As NXOpen.DisplayableObject

        objects10(0) = iPmi
        displayModification1.Apply(objects10)

        displayModification1.Dispose()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_Vertical_Pnt_Pnt(ByVal point_1 As NXOpen.Point, ByVal point_2 As NXOpen.Point, ByVal datumplane1 As DatumPlane, ByVal Plane_dir As String, ByVal Diem_dat As Point3d, ByVal text_center As Boolean, ByVal dk As Double, ByVal dk2 As Double, ByVal doimau As String, ByVal text_size As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing
        Dim pmiRapidDimensionBuilder1 As NXOpen.Annotations.PmiRapidDimensionBuilder = Nothing
        pmiRapidDimensionBuilder1 = workPart.Dimensions.CreatePmiRapidDimensionBuilder(nullNXOpen_Annotations_Dimension)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Vertical

        If datumplane1 Is Nothing Then
        Else
            Dim xform1 As NXOpen.Xform = Nothing
            xform1 = workPart.Xforms.CreateXform(datumplane1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
            cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.UserDefined

            pmiRapidDimensionBuilder1.Origin.Plane.UserDefinedPlane = xform1
        End If

        If Plane_dir = "Y" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane
        ElseIf Plane_dir = "Z" Or Plane_dir = "SA" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane
        ElseIf Plane_dir = "X" Or Plane_dir = "SB" Or Plane_dir = "SC" Or Plane_dir = "SD" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane
        Else
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If


        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        pmiRapidDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction
        Dim nullNXOpen_View As NXOpen.View = Nothing
        pmiRapidDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View
        pmiRapidDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        Dim point1 As NXOpen.Point = point_1 'CType(workPart.Points.FindObject("ENTITY 2 6 1"), NXOpen.Point)
        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing
        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)
        Dim point2 As NXOpen.Point = point_2 'CType(workPart.Points.FindObject("ENTITY 2 7 1"), NXOpen.Point)
        Dim point1_3 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim point1_5 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, -862.7699486541685, 663.69329279106523)
        Dim point2_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_5, Nothing, nullNXOpen_View, point2_5)

        Dim point1_6 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z) '(1562.0191452238341, -622.66952709460372, 1143.0377650354451)
        Dim point2_6 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point2, displayPart.ModelingViews.WorkView, point1_6, Nothing, nullNXOpen_View, point2_6)

        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects2(1) As NXOpen.NXObject
        objects2(0) = point2
        objects2(1) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiRapidDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d = Diem_dat 'New NXOpen.Point3d(point_1.Coordinates.X + 300, point_1.Coordinates.Y + 50, point_1.Coordinates.Z) '(936.13446125343967, -309.523651364787, 663.69329279106523)
        pmiRapidDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Right
        pmiRapidDimensionBuilder1.Style.DimensionStyle.TextCentered = text_center
        pmiRapidDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = text_size
        pmiRapidDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiRapidDimensionBuilder1.Commit()

        pmiRapidDimensionBuilder1.Destroy()
        '-------------------------------------------------------------------------------------------------------------------------------
        Dim iPmi As NXOpen.Annotations.PmiVerticalDimension = CType(nXObject1, NXOpen.Annotations.PmiVerticalDimension)
        Dim ii() As String
        Dim ikhoangcach As String
        iPmi.GetDimensionText(ii, Nothing)
        For Each istring As String In ii
            ikhoangcach = CType(istring, Double)
        Next
        Call PMI_doidoday(iPmi)
        '-----------------------------------------------------------------------
        'doi mau
        Dim icolor As Double
        If ikhoangcach < dk2 Then
            icolor = 108
        Else
            icolor = 186
        End If

        If doimau = "OK" Then
            icolor = 108
        ElseIf doimau = "NG" Then
            icolor = 186
        End If

        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewColor = icolor

        Dim objects10(0) As NXOpen.DisplayableObject

        objects10(0) = iPmi
        displayModification1.Apply(objects10)

        displayModification1.Dispose()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_Radial_arc(ByVal arc1 As NXOpen.Arc, ByVal point1 As NXOpen.Point3d, ByVal Plane_dir As String, ByVal dk As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim Chuvi As Double = CType(arc1.GetLength, Double)
        Dim bankinh As Double = Chuvi / (2 * Math.PI)
        ' ----------------------------------------------
        '   Menu: PMI->Dimension->Radial...
        ' ----------------------------------------------
        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing

        Dim pmiRadialDimensionBuilder1 As NXOpen.Annotations.PmiRadialDimensionBuilder = Nothing
        pmiRadialDimensionBuilder1 = workPart.Dimensions.CreatePmiRadialDimensionBuilder(nullNXOpen_Annotations_Dimension)
        If Plane_dir = "Y" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane

            pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "Z" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

            pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "X" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane

            pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        Else
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(True)
        pmiRadialDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Radial
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        pmiRadialDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing
        pmiRadialDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View
        pmiRadialDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None


        pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(arc1, displayPart.ModelingViews.WorkView, point1)

        Dim point1_1 As NXOpen.Point3d = point1 'New NXOpen.Point3d(1772.8821961992139, -1180.8311593265621, 0.0)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, arc1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing

        pmiRadialDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = arc1
        pmiRadialDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)

        Dim dimensionlinearunits31 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits31 = pmiRadialDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim dimensionlinearunits32 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits32 = pmiRadialDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiRadialDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point2 As NXOpen.Point3d = New NXOpen.Point3d(point1.X + 2 * bankinh + 15, point1.Y - 150, 0) '(2046.1712012738183, -1368.1597921605073, 0.0)
        pmiRadialDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point2)
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.DiameterDimensionDimLineAngle = 0.0
        pmiRadialDimensionBuilder1.Style.DimensionStyle.TextAngle = 332.27899907725464
        pmiRadialDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Left
        pmiRadialDimensionBuilder1.Style.DimensionStyle.TextCentered = False
        pmiRadialDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = texthigh
        pmiRadialDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = texthigh
        pmiRadialDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = texthigh
        pmiRadialDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = texthigh
        pmiRadialDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiRadialDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim objects2(0) As NXOpen.NXObject
        objects2(0) = arc1
        pmiRadialDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiRadialDimensionBuilder1.Commit()

        pmiRadialDimensionBuilder1.Destroy()
        '-------------------------------------------------------------------------------------------------------------------------------
        Dim iPmi As NXOpen.Annotations.PmiRadiusDimension = CType(nXObject1, NXOpen.Annotations.PmiRadiusDimension)
        Dim ii() As String
        Dim ikhoangcach As String
        iPmi.GetDimensionText(ii, Nothing)
        For Each istring As String In ii
            ikhoangcach = CType(istring, Double)
        Next
        Call PMI_doidoday(iPmi)
        '-----------------------------------------------------------------------
        'doi mau
        Dim icolor As Double
        If ikhoangcach > dk Then
            icolor = 108
        Else
            icolor = 186
        End If
        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewColor = icolor

        Dim objects10(0) As NXOpen.DisplayableObject

        objects10(0) = iPmi
        displayModification1.Apply(objects10)

        displayModification1.Dispose()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_vertical_Pnt_Line(ByVal point_1 As NXOpen.Point, ByVal line_1 As NXOpen.Curve, ByVal Plane_dir As String, ByVal diemdatPMI As Point3d, ByVal text_Center As Boolean, ByVal dk_min As Double, ByVal dk_max As Double, ByVal doimau As String, ByVal Text_size As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: PMI->Dimension->Rapid...
        ' ----------------------------------------------
        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing

        Dim pmiRapidDimensionBuilder1 As NXOpen.Annotations.PmiRapidDimensionBuilder = Nothing
        pmiRapidDimensionBuilder1 = workPart.Dimensions.CreatePmiRapidDimensionBuilder(nullNXOpen_Annotations_Dimension)
        If Plane_dir = "Y" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane

            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "SA" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        ElseIf Plane_dir = "SC" Or Plane_dir = "SD" Then
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane

            pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        Else
            pmiRapidDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        pmiRapidDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing
        pmiRapidDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View
        pmiRapidDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        'Dim sketchFeature1 As NXOpen.Features.SketchFeature = CType(workPart.Features.FindObject("SKETCH(5)"), NXOpen.Features.SketchFeature)

        'Dim sketch1 As NXOpen.Sketch = CType(sketchFeature1.FindObject("SKETCH_002"), NXOpen.Sketch)

        Dim point1 As NXOpen.Point = point_1 'CType(sketch1.FindObject("Point Point5"), NXOpen.Point)

        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, 0.0, 663.69329279106523)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, 0.0, 663.69329279106523)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing

        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)

        Dim line1 As NXOpen.Curve = line_1 'CType(sketch1.FindObject("Curve Line1"), NXOpen.Line)

        Dim point2 As NXOpen.Point3d = timdiemdau("Z", True, line1) 'New NXOpen.Point3d(1465.496666091185, -0.0000000000043200998334214091, -126.81963613538807)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(line1, displayPart.ModelingViews.WorkView, point2)

        Dim point1_3 As NXOpen.Point3d = point2 'New NXOpen.Point3d(1465.496666091185, -0.0000000000043200998334214091, -126.81963613538807)
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, line1, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z) '(1298.5375912757734, 0.0, 663.69329279106523)
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiRapidDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        pmiRapidDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects2(1) As NXOpen.NXObject
        objects2(0) = line1
        objects2(1) = point1
        pmiRapidDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiRapidDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d = diemdatPMI 'New NXOpen.Point3d(point_1.Coordinates.X + 100, point_1.Coordinates.Y, point_1.Coordinates.Z - 50) '(1639.3594590296225, 0.0, 269.8119400167011)
        pmiRapidDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Right
        pmiRapidDimensionBuilder1.Style.DimensionStyle.TextCentered = text_Center
        pmiRapidDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = Text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = Text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = Text_size
        pmiRapidDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = Text_size
        pmiRapidDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRapidDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True
        pmiRapidDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiRapidDimensionBuilder1.Commit()

        pmiRapidDimensionBuilder1.Destroy()
        '-------------------------------------------------------------------------------------------------------------------------------
        Dim iPmi As NXOpen.Annotations.PmiPerpendicularDimension = CType(nXObject1, NXOpen.Annotations.PmiPerpendicularDimension)
        Dim ii() As String
        Dim ikhoangcach As String
        iPmi.GetDimensionText(ii, Nothing)
        For Each istring As String In ii
            ikhoangcach = CType(istring, Double)
        Next
        Call PMI_doidoday(iPmi)
        '-----------------------------------------------------------------------
        'doi mau
        Dim icolor As Double
        If ikhoangcach > dk_min And ikhoangcach < dk_max Then
            icolor = 108
        Else
            icolor = 186
        End If

        If doimau = "OK" Then
            icolor = 108
        ElseIf doimau = "NG" Then
            icolor = 186
        End If
        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewColor = icolor

        Dim objects10(0) As NXOpen.DisplayableObject

        objects10(0) = iPmi
        displayModification1.Apply(objects10)

        displayModification1.Dispose()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_Pnt_Pnt(ByVal point_1 As NXOpen.Point, point_2 As NXOpen.Point, ByVal datumplane1 As DatumPlane, ByVal Plane_dir As String, ByVal diemdatPMI As Point3d, ByVal text_Center As Boolean, ByVal dk_min As Double, ByVal dk_max As Double, ByVal doimau As String, ByVal Text_size As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: PMI->Dimension->Linear...
        ' ---------------------------------------------
        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing

        Dim pmiLinearDimensionBuilder1 As NXOpen.Annotations.PmiLinearDimensionBuilder = Nothing
        pmiLinearDimensionBuilder1 = workPart.Dimensions.CreatePmiLinearDimensionBuilder(nullNXOpen_Annotations_Dimension)

        pmiLinearDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.PointToPoint
        If datumplane1 Is Nothing Then
        Else
            Dim xform1 As NXOpen.Xform = Nothing
            xform1 = workPart.Xforms.CreateXform(datumplane1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
            cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            pmiLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.UserDefined

            pmiLinearDimensionBuilder1.Origin.Plane.UserDefinedPlane = xform1
        End If

        If Plane_dir = "Y" Then
            pmiLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane
        ElseIf Plane_dir = "Z" Or Plane_dir = "SA" Then
            pmiLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane
        ElseIf Plane_dir = "X" Or Plane_dir = "SB" Or Plane_dir = "SC" Or Plane_dir = "SD" Then
            pmiLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane
        Else
            pmiLinearDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If


        pmiLinearDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        pmiLinearDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        'Dim pointFeature1 As NXOpen.Features.PointFeature = CType(workPart.Features.FindObject("POINT(9)"), NXOpen.Features.PointFeature)

        Dim point1 As NXOpen.Point = point_1

        Dim point1_1 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z)
        Dim nullNXOpen_View As NXOpen.View = Nothing

        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = New NXOpen.Point3d(point_1.Coordinates.X, point_1.Coordinates.Y, point_1.Coordinates.Z)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point1, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing

        pmiLinearDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = point1
        pmiLinearDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)

        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim dimensionlinearunits21 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits21 = pmiLinearDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim dimensionlinearunits22 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits22 = pmiLinearDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        'Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT @DB/NML49515911/AA 1"), NXOpen.Assemblies.Component)

        Dim point2 As NXOpen.Point = point_2

        Dim point1_3 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z)
        Dim point2_3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_3, Nothing, nullNXOpen_View, point2_3)

        Dim point1_4 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z)
        Dim point2_4 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Exist, point2, displayPart.ModelingViews.WorkView, point1_4, Nothing, nullNXOpen_View, point2_4)

        Dim point1_5 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z)
        Dim point2_5 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point1, displayPart.ModelingViews.WorkView, point1_5, Nothing, nullNXOpen_View, point2_5)

        Dim point3 As NXOpen.Point = point_2

        Dim point1_6 As NXOpen.Point3d = New NXOpen.Point3d(point_2.Coordinates.X, point_2.Coordinates.Y, point_2.Coordinates.Z)
        Dim point2_6 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiLinearDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.Mid, point3, displayPart.ModelingViews.WorkView, point1_6, Nothing, nullNXOpen_View, point2_6)

        pmiLinearDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects2(1) As NXOpen.NXObject
        objects2(0) = point3
        objects2(1) = point1
        pmiLinearDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiLinearDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point4 As NXOpen.Point3d = diemdatPMI
        pmiLinearDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point4)

        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        pmiLinearDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Left

        pmiLinearDimensionBuilder1.Style.DimensionStyle.TextCentered = text_Center
        pmiLinearDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = Text_size
        pmiLinearDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = Text_size
        pmiLinearDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = Text_size
        pmiLinearDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = Text_size
        pmiLinearDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiLinearDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiLinearDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True
        pmiLinearDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        Dim objects3(1) As NXOpen.NXObject
        objects3(0) = point3
        objects3(1) = point1
        pmiLinearDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects3)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiLinearDimensionBuilder1.Commit()

        pmiLinearDimensionBuilder1.Destroy()
        Dim iPmi As NXOpen.Annotations.PmiParallelDimension = CType(nXObject1, NXOpen.Annotations.PmiParallelDimension)
        Dim ii() As String
        Dim ikhoangcach As String
        iPmi.GetDimensionText(ii, Nothing)
        For Each istring As String In ii
            ikhoangcach = CType(istring, Double)
        Next
        Call PMI_doidoday(iPmi)
        '-----------------------------------------------------------------------
        If doimau Is Nothing And dk_max = 999 And dk_min = 0 Then
            Exit Sub
        End If

        'doi mau
        Dim icolor As Double
        If ikhoangcach > dk_min And ikhoangcach < dk_max Then
            icolor = 108
        Else
            icolor = 186
        End If

        If doimau = "OK" Then
            icolor = 108
        ElseIf doimau = "NG" Then
            icolor = 186
        End If
        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewColor = icolor

        Dim objects10(0) As NXOpen.DisplayableObject

        objects10(0) = iPmi
        displayModification1.Apply(objects10)

        displayModification1.Dispose()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_Angle(ByVal line1 As Curve, ByVal line2 As Curve, ByVal Text_size As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: PMI->Dimension->Angular...
        ' ----------------------------------------------

        Dim nullNXOpen_Annotations_BaseAngularDimension As NXOpen.Annotations.BaseAngularDimension = Nothing

        Dim pmiMinorAngularDimensionBuilder1 As NXOpen.Annotations.PmiMinorAngularDimensionBuilder = Nothing
        pmiMinorAngularDimensionBuilder1 = workPart.Dimensions.CreatePmiMinorAngularDimensionBuilder(nullNXOpen_Annotations_BaseAngularDimension)

        pmiMinorAngularDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView

        pmiMinorAngularDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

        pmiMinorAngularDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        pmiMinorAngularDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing

        pmiMinorAngularDimensionBuilder1.FirstVector = nullNXOpen_Direction

        pmiMinorAngularDimensionBuilder1.SecondVector = nullNXOpen_Direction

        'Dim line1 As NXOpen.Line = CType(workPart.Lines.FindObject("ENTITY 3 1 1"), NXOpen.Line)

        Dim point1 As NXOpen.Point3d = timdiemdau("Z", False, line1) 'New NXOpen.Point3d(1443.9819202107833, -807.50325144451972, 677.03807383575077)
        pmiMinorAngularDimensionBuilder1.FirstAssociativity.SetValue(line1, displayPart.ModelingViews.WorkView, point1)

        pmiMinorAngularDimensionBuilder1.FirstVector = nullNXOpen_Direction

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = line1
        pmiMinorAngularDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)

        'Dim line2 As NXOpen.Line = CType(workPart.Lines.FindObject("ENTITY 3 21 1"), NXOpen.Line)

        Dim point2 As NXOpen.Point3d = timdiemdau("Z", False, line2) 'New NXOpen.Point3d(1444.4413280583231, -822.62159465433228, 674.38371940724574)
        pmiMinorAngularDimensionBuilder1.SecondAssociativity.SetValue(line2, displayPart.ModelingViews.WorkView, point2)

        pmiMinorAngularDimensionBuilder1.SecondVector = nullNXOpen_Direction

        Dim objects2(1) As NXOpen.NXObject
        objects2(0) = line2
        objects2(1) = line1
        pmiMinorAngularDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim point1_1 As NXOpen.Point3d = point1 'New NXOpen.Point3d(1443.4189596708518, -806.69041852173348, 687.5584510674762)
        Dim nullNXOpen_View As NXOpen.View = Nothing

        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiMinorAngularDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, line1, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)

        Dim point1_2 As NXOpen.Point3d = point2 'New NXOpen.Point3d(1444.1923439276566, -831.95052014880775, 683.011776581489)
        Dim point2_2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        pmiMinorAngularDimensionBuilder1.SecondAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, line2, displayPart.ModelingViews.WorkView, point1_2, Nothing, nullNXOpen_View, point2_2)

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiMinorAngularDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point3 As NXOpen.Point3d
        If point1.Z > point2.Z Then
            point3 = New NXOpen.Point3d(point1.X, point1.Y + 30, point1.Z + 50)
        Else
            point3 = New NXOpen.Point3d(point1.X, point1.Y + 30, point1.Z + 50)
        End If

        pmiMinorAngularDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point3)

        pmiMinorAngularDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        pmiMinorAngularDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Left

        pmiMinorAngularDimensionBuilder1.Style.DimensionStyle.TextCentered = True
        pmiMinorAngularDimensionBuilder1.Style.LetteringStyle.DimensionTextSize = Text_size
        pmiMinorAngularDimensionBuilder1.Style.LetteringStyle.AppendedTextSize = Text_size
        pmiMinorAngularDimensionBuilder1.Style.LetteringStyle.ToleranceTextSize = Text_size
        pmiMinorAngularDimensionBuilder1.Style.LetteringStyle.TwoLineToleranceTextSize = Text_size
        pmiMinorAngularDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 1
        pmiMinorAngularDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 1
        pmiMinorAngularDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiMinorAngularDimensionBuilder1.Style.UnitsStyle.DisplayTrailingZeros = True
        pmiMinorAngularDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiMinorAngularDimensionBuilder1.Commit()
        Dim iPmi As NXOpen.Annotations.BaseAngularDimension = CType(nXObject1, NXOpen.Annotations.BaseAngularDimension)
        Call PMI_doidoday(iPmi)
        pmiMinorAngularDimensionBuilder1.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_Diameter(ByVal arc1 As NXOpen.Arc, ByVal point1 As NXOpen.Point3d, ByVal datumplane1 As DatumPlane, ByVal Plane_dir As String, ByVal dk As Double)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        Dim nullNXOpen_Annotations_Dimension As NXOpen.Annotations.Dimension = Nothing
        Dim pmiRadialDimensionBuilder1 As NXOpen.Annotations.PmiRadialDimensionBuilder = Nothing
        pmiRadialDimensionBuilder1 = workPart.Dimensions.CreatePmiRadialDimensionBuilder(nullNXOpen_Annotations_Dimension)
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Radial
        pmiRadialDimensionBuilder1.IsRadiusToCenter = True
        pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)

        Dim nullNXOpen_Direction As NXOpen.Direction = Nothing
        pmiRadialDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction

        Dim nullNXOpen_View As NXOpen.View = Nothing
        pmiRadialDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View
        pmiRadialDimensionBuilder1.Style.DimensionStyle.NarrowDisplayType = NXOpen.Annotations.NarrowDisplayOption.None
        pmiRadialDimensionBuilder1.Measurement.Method = NXOpen.Annotations.DimensionMeasurementBuilder.MeasurementMethod.Diametral

        If datumplane1 Is Nothing Then
        Else
            Dim xform1 As NXOpen.Xform = Nothing
            xform1 = workPart.Xforms.CreateXform(datumplane1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
            cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.AfterModeling)

            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.UserDefined

            pmiRadialDimensionBuilder1.Origin.Plane.UserDefinedPlane = xform1
        End If

        If Plane_dir = "Y" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XzPlane
        ElseIf Plane_dir = "Z" Or Plane_dir = "SA" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane
        ElseIf Plane_dir = "X" Or Plane_dir = "SB" Or Plane_dir = "SC" Or Plane_dir = "SD" Then
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.YzPlane
        Else
            pmiRadialDimensionBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.ModelView
        End If

        pmiRadialDimensionBuilder1.Measurement.Direction = nullNXOpen_Direction
        pmiRadialDimensionBuilder1.Measurement.DirectionView = nullNXOpen_View

        'Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT @DB/NML26776909/AA.004 1"), NXOpen.Assemblies.Component)
        'Dim component2 As NXOpen.Assemblies.Component = CType(component1.FindObject("COMPONENT @DB/NML26776903/AA.002 1"), NXOpen.Assemblies.Component)
        'Dim arc1 As NXOpen.Arc = CType(component2.FindObject("PROTO#.Features|SKETCH(2)|SKETCH_000|HANDLE R-8918"), NXOpen.Arc)
        'Dim point1 As NXOpen.Point3d = New NXOpen.Point3d(1241.04598257328, -819.37835775830786, 817.87461075577232)
        Try
            pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(arc1, workPart.ModelingViews.WorkView, point1)
        Catch ex As Exception
            pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(arc1, displayPart.ModelingViews.WorkView, point1)
        End Try


        Dim arc2 As NXOpen.Arc = arc1 'CType(component2.FindObject("PROTO#.Features|SKETCH(2)|SKETCH_000|HANDLE R-8918"), NXOpen.Arc)

        Dim point1_1 As NXOpen.Point3d = point1 'New NXOpen.Point3d(1241.04598257328, -819.37835775830786, 817.87461075577232)
        Dim point2_1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Try
            pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, arc2, workPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)
        Catch ex As Exception
            pmiRadialDimensionBuilder1.FirstAssociativity.SetValue(NXOpen.InferSnapType.SnapType.None, arc2, displayPart.ModelingViews.WorkView, point1_1, Nothing, nullNXOpen_View, point2_1)
        End Try


        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing

        pmiRadialDimensionBuilder1.Measurement.PartOccurrence = nullNXOpen_Assemblies_Component

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = arc2
        pmiRadialDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects1)

        Dim dimensionlinearunits31 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits31 = pmiRadialDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim dimensionlinearunits32 As NXOpen.Annotations.DimensionUnit = Nothing
        dimensionlinearunits32 = pmiRadialDimensionBuilder1.Style.UnitsStyle.DimensionLinearUnits

        Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData = Nothing
        assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
        assocOrigin1.View = nullNXOpen_View
        assocOrigin1.ViewOfGeometry = nullNXOpen_View
        Dim nullNXOpen_Point As NXOpen.Point = Nothing

        assocOrigin1.PointOnGeometry = nullNXOpen_Point
        Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

        assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.DimensionLine = 0
        assocOrigin1.AssociatedView = nullNXOpen_View
        assocOrigin1.AssociatedPoint = nullNXOpen_Point
        assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
        assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
        assocOrigin1.XOffsetFactor = 0.0
        assocOrigin1.YOffsetFactor = 0.0
        assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
        pmiRadialDimensionBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

        Dim point2 As NXOpen.Point3d = New NXOpen.Point3d(1229.905727330171, -821.65159594100533, 817.83687587895724)
        pmiRadialDimensionBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point2)
        pmiRadialDimensionBuilder1.Origin.SetInferRelativeToGeometry(False)
        pmiRadialDimensionBuilder1.DiameterDimensionDimLineAngle = 70.064528513116684
        pmiRadialDimensionBuilder1.Style.DimensionStyle.TextAngle = 54.401777375601149
        pmiRadialDimensionBuilder1.Style.LineArrowStyle.LeaderOrientation = NXOpen.Annotations.LeaderSide.Right
        pmiRadialDimensionBuilder1.Style.DimensionStyle.TextCentered = True

        Dim objects2(0) As NXOpen.NXObject
        objects2(0) = arc2
        pmiRadialDimensionBuilder1.AssociatedObjects.Nxobjects.SetArray(objects2)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = pmiRadialDimensionBuilder1.Commit()
        Dim iPmi As NXOpen.Annotations.Dimension = CType(nXObject1, NXOpen.Annotations.Dimension)
        Call PMI_doidoday(iPmi)
        pmiRadialDimensionBuilder1.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub PMI_doidoday(ByVal pmiPerpendicularDimension1 As NXOpen.DisplayableObject)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Edit->Settings...
        ' ----------------------------------------------
        Dim objects1(0) As NXOpen.DisplayableObject
        'Dim pmiPerpendicularDimension1 As NXOpen.Annotations.PmiPerpendicularDimension = CType(workPart.FindObject("ENTITY 26 3 1"), NXOpen.Annotations.PmiPerpendicularDimension)

        objects1(0) = pmiPerpendicularDimension1
        Dim editSettingsBuilder1 As NXOpen.Annotations.EditSettingsBuilder = Nothing
        editSettingsBuilder1 = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects1)

        Dim editsettingsbuilders1(0) As NXOpen.Drafting.BaseEditSettingsBuilder
        editsettingsbuilders1(0) = editSettingsBuilder1
        workPart.SettingsManager.ProcessForMultipleObjectsSettings(editsettingsbuilders1)

        Dim fontIndex1 As Integer = Nothing
        fontIndex1 = workPart.Fonts.AddFont("kanji2", NXOpen.FontCollection.Type.Nx)

        editSettingsBuilder1.AnnotationStyle.LetteringStyle.DimensionTextLineWidth = NXOpen.Annotations.LineWidth.Thick
        'editSettingsBuilder1.AnnotationStyle.LetteringStyle.AppendedTextColor = displayPart.Colors.Find("Emerald")
        editSettingsBuilder1.AnnotationStyle.LetteringStyle.AppendedTextLineWidth = NXOpen.Annotations.LineWidth.Thick
        'editSettingsBuilder1.AnnotationStyle.LetteringStyle.AppendedTextSize = 15.0
        'editSettingsBuilder1.AnnotationStyle.LetteringStyle.ToleranceTextColor = displayPart.Colors.Find("Emerald")
        editSettingsBuilder1.AnnotationStyle.LetteringStyle.ToleranceTextLineWidth = NXOpen.Annotations.LineWidth.Thick
        editSettingsBuilder1.AnnotationStyle.LetteringStyle.GeneralTextLineWidth = NXOpen.Annotations.LineWidth.Thick
        'editSettingsBuilder1.AnnotationStyle.LetteringStyle.ToleranceTextSize = 15.0
        'editSettingsBuilder1.AnnotationStyle.LetteringStyle.TwoLineToleranceTextSize = 15.0

        ' ----------------------------------------------
        '   Dialog Begin Settings
        ' ----------------------------------------------
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.SecondExtensionLineWidth = NXOpen.Annotations.LineWidth.Four
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.FirstExtensionLineWidth = NXOpen.Annotations.LineWidth.Four
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.FirstArrowheadWidth = NXOpen.Annotations.LineWidth.Four
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.SecondArrowheadWidth = NXOpen.Annotations.LineWidth.Four
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.FirstArrowLineWidth = NXOpen.Annotations.LineWidth.Four
        editSettingsBuilder1.AnnotationStyle.LineArrowStyle.SecondArrowLineWidth = NXOpen.Annotations.LineWidth.Four

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = editSettingsBuilder1.Commit()

        editSettingsBuilder1.Destroy()
        theSession.CleanUpFacetedFacesAndEdges()

    End Sub
    '------------------------thao tac Sketch------------------------------
    Sub mosketch(ByVal DatumPlane1 As DatumPlane, ByVal Sketch_No As Double, coordinates1 As NXOpen.Point3d)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        'mo sketch
        Dim nullNXOpen_Sketch As NXOpen.Sketch = Nothing

        Dim sketchInPlaceBuilder1 As NXOpen.SketchInPlaceBuilder = Nothing
        sketchInPlaceBuilder1 = workPart.Sketches.CreateSketchInPlaceBuilder2(nullNXOpen_Sketch)

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.PlaneReference = plane1

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        sketchInPlaceBuilder1.OriginOption = NXOpen.OriginMethod.WorkPartOrigin

        Dim sketchAlongPathBuilder1 As NXOpen.SketchAlongPathBuilder = Nothing
        sketchAlongPathBuilder1 = workPart.Sketches.CreateSketchAlongPathBuilder(nullNXOpen_Sketch)

        sketchAlongPathBuilder1.PlaneLocation.Expression.RightHandSide = "0"

        Dim point1 As NXOpen.Point = Nothing
        point1 = workPart.Points.CreatePoint(coordinates1)

        Dim datumAxis1 As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) Y axis"), NXOpen.DatumAxis)

        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(datumAxis1, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim xform1 As NXOpen.Xform = Nothing
        xform1 = workPart.Xforms.CreateXformByPlaneXDirPoint(DatumPlane1, direction1, point1, NXOpen.SmartObject.UpdateOption.WithinModeling, 0.5, False, False)

        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
        cartesianCoordinateSystem1 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        sketchInPlaceBuilder1.Csystem = cartesianCoordinateSystem1

        Dim origin2 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal2 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane2 As NXOpen.Plane = Nothing
        plane2 = workPart.Planes.CreatePlane(origin2, normal2, NXOpen.SmartObject.UpdateOption.WithinModeling)

        plane2.SetMethod(NXOpen.PlaneTypes.MethodType.Coincident)

        Dim geom1(0) As NXOpen.NXObject
        geom1(0) = DatumPlane1
        plane2.SetGeometry(geom1)

        plane2.SetFlip(False)

        plane2.SetExpression(Nothing)

        plane2.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)

        plane2.Evaluate()

        Dim origin3 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim normal3 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim plane3 As NXOpen.Plane = Nothing
        plane3 = workPart.Planes.CreatePlane(origin3, normal3, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim expression3 As NXOpen.Expression = Nothing
        expression3 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression4 As NXOpen.Expression = Nothing
        expression4 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        plane3.SynchronizeToPlane(plane2)

        plane3.Evaluate()

        plane3.SetMethod(NXOpen.PlaneTypes.MethodType.Coincident)
        'chon datum
        Dim geom2(0) As NXOpen.NXObject
        geom2(0) = DatumPlane1
        plane3.SetGeometry(geom2)

        plane3.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)

        plane3.Evaluate()

        theSession.Preferences.Sketch.CreateInferredConstraints = True

        theSession.Preferences.Sketch.ContinuousAutoDimensioning = False

        theSession.Preferences.Sketch.DimensionLabel = NXOpen.Preferences.SketchPreferences.DimensionLabelType.Expression

        theSession.Preferences.Sketch.TextSizeFixed = True

        theSession.Preferences.Sketch.FixedTextSize = 3.0

        theSession.Preferences.Sketch.DisplayParenthesesOnReferenceDimensions = True

        theSession.Preferences.Sketch.DisplayReferenceGeometry = True

        theSession.Preferences.Sketch.ConstraintSymbolSize = 3.0

        theSession.Preferences.Sketch.DisplayObjectColor = True

        theSession.Preferences.Sketch.DisplayObjectName = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchInPlaceBuilder1.Commit()

        Dim sketch1 As NXOpen.Sketch = CType(nXObject1, NXOpen.Sketch)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketch1.Feature

        Dim nErrs1 As Long = Nothing

        sketch1.Activate(NXOpen.Sketch.ViewReorient.False)


        sketchInPlaceBuilder1.Destroy()

        sketchAlongPathBuilder1.Destroy()

        plane3.DestroyPlane()
        theSession.ActiveSketch.SetName("SKETCH_00" & Sketch_No)

    End Sub
    Sub Laygiaodiemvoimatsketch(ByVal line1 As NXOpen.Line, ByRef cent_pnt As NXOpen.Point)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Curve from Curves->Intersection Point...
        ' ----------------------------------------------

        Dim nullNXOpen_SketchIntersectionPoint As NXOpen.SketchIntersectionPoint = Nothing

        Dim sketchIntersectionPointBuilder1 As NXOpen.SketchIntersectionPointBuilder = Nothing
        sketchIntersectionPointBuilder1 = workPart.Sketches.CreateIntersectionPointBuilder(nullNXOpen_SketchIntersectionPoint)

        Dim section1 As NXOpen.Section = Nothing
        section1 = sketchIntersectionPointBuilder1.Rail

        section1.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim curves1(0) As NXOpen.IBaseCurve
        'Dim compositeCurve1 As NXOpen.Features.CompositeCurve = CType(workPart.Features.FindObject("LINKED_CURVE(1)"), NXOpen.Features.CompositeCurve)

        'Dim line1 As NXOpen.Line = CType(compositeCurve1.FindObject("CURVE 1"), NXOpen.Line)

        curves1(0) = line1
        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        section1.AllowSelfIntersection(True)

        section1.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", True, line1) 'New NXOpen.Point3d(442.99999999999704, -835.39251672714317, 381.17768651902975)
        section1.AddToSection(rules1, line1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        sketchIntersectionPointBuilder1.UpdateData()

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = sketchIntersectionPointBuilder1.CommitFeature()
        'Dim iintesec As NXOpen.Features.PointFeature = CType(feature1, NXOpen.Features.PointFeature)

        'cent_pnt = CType(iintesec, NXOpen.Point)
        'lw.WriteLine(feature1.GetFeatureColor.ToString)
        Dim ipoint1 As NXOpen.Point
        For Each iObj As NXOpen.NXObject In feature1.GetEntities
            If TypeOf (iObj) Is NXOpen.Point Then
                cent_pnt = iObj
                'ElseIf TypeOf (iObj) Is CartesianCoordinateSystem Then
                '    tmpCs = iObj
            End If
        Next

        sketchIntersectionPointBuilder1.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub veduongtron_tamvadiem(ByVal center_p As NXOpen.Point, ByVal bankinh As Double, ByRef arc1 As NXOpen.Arc)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        Dim nXMatrix1 As NXOpen.NXMatrix = Nothing
        nXMatrix1 = theSession.ActiveSketch.Orientation

        Dim center1 As NXOpen.Point3d = New NXOpen.Point3d(center_p.Coordinates.X, center_p.Coordinates.Y, center_p.Coordinates.Z)
        'Dim arc1 As NXOpen.Arc = Nothing
        arc1 = workPart.Curves.CreateArc(center1, nXMatrix1, bankinh, 0.0, (360.0 * Math.PI / 180.0))

        theSession.ActiveSketch.AddGeometry(arc1, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)

        Dim geom1_1 As NXOpen.Sketch.ConstraintGeometry = Nothing
        geom1_1.Geometry = arc1
        geom1_1.PointType = NXOpen.Sketch.ConstraintPointType.ArcCenter
        geom1_1.SplineDefiningPointIndex = 0
        Dim geom2_1 As NXOpen.Sketch.ConstraintGeometry = Nothing
        Dim point1 As NXOpen.Point = CType(center_p, NXOpen.Point)

        geom2_1.Geometry = point1
        geom2_1.PointType = NXOpen.Sketch.ConstraintPointType.None
        geom2_1.SplineDefiningPointIndex = 0
        Dim sketchGeometricConstraint1 As NXOpen.SketchGeometricConstraint = Nothing
        sketchGeometricConstraint1 = theSession.ActiveSketch.CreateCoincidentConstraint(geom1_1, geom2_1)

        theSession.ActiveSketch.Update()

    End Sub
    Sub veduongtrongsketch(ByVal point01 As NXOpen.Point, ByVal point02 As NXOpen.Point, ByRef line1 As NXOpen.Line)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------

        Dim startPoint1 As NXOpen.Point3d = New NXOpen.Point3d(point01.Coordinates.X, point01.Coordinates.Y, point01.Coordinates.Z)
        Dim endPoint1 As NXOpen.Point3d = New NXOpen.Point3d(point02.Coordinates.X, point02.Coordinates.Y, point02.Coordinates.Z)
        'Dim line1 As NXOpen.Line = Nothing
        line1 = workPart.Curves.CreateLine(startPoint1, endPoint1)

        theSession.ActiveSketch.AddGeometry(line1, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)

        Dim geom1_1 As NXOpen.Sketch.ConstraintGeometry = Nothing
        geom1_1.Geometry = line1
        geom1_1.PointType = NXOpen.Sketch.ConstraintPointType.StartVertex
        geom1_1.SplineDefiningPointIndex = 0
        Dim geom2_1 As NXOpen.Sketch.ConstraintGeometry = Nothing
        Dim point1 As NXOpen.Point = point01 'CType(theSession.ActiveSketch.FindObject("Point Point5"), NXOpen.Point)

        geom2_1.Geometry = point1
        geom2_1.PointType = NXOpen.Sketch.ConstraintPointType.None
        geom2_1.SplineDefiningPointIndex = 0
        Dim sketchGeometricConstraint1 As NXOpen.SketchGeometricConstraint = Nothing
        sketchGeometricConstraint1 = theSession.ActiveSketch.CreateCoincidentConstraint(geom1_1, geom2_1)

        Dim geom1_2 As NXOpen.Sketch.ConstraintGeometry = Nothing
        geom1_2.Geometry = line1
        geom1_2.PointType = NXOpen.Sketch.ConstraintPointType.EndVertex
        geom1_2.SplineDefiningPointIndex = 0
        Dim geom2_2 As NXOpen.Sketch.ConstraintGeometry = Nothing
        Dim point2 As NXOpen.Point = point02 'CType(theSession.ActiveSketch.FindObject("Point Point6"), NXOpen.Point)

        geom2_2.Geometry = point2
        geom2_2.PointType = NXOpen.Sketch.ConstraintPointType.None
        geom2_2.SplineDefiningPointIndex = 0
        Dim sketchGeometricConstraint2 As NXOpen.SketchGeometricConstraint = Nothing
        sketchGeometricConstraint2 = theSession.ActiveSketch.CreateCoincidentConstraint(geom1_2, geom2_2)

        theSession.ActiveSketch.Update()
    End Sub
    Sub chieuduong(ByVal arc1 As NXOpen.Curve, ByRef arc2 As NXOpen.Curve, ByVal ifeature1 As NXOpen.Features.Feature) 'ByRef arc2 As NXOpen.Ellipse
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Associative Curve->Project Curve...
        ' ---------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim sketchProjectBuilder1 As NXOpen.SketchProjectBuilder = Nothing
        sketchProjectBuilder1 = workPart.Sketches.CreateProjectBuilder(nullNXOpen_Features_Feature)

        sketchProjectBuilder1.Tolerance = 0.01

        sketchProjectBuilder1.Section.PrepareMappingData()

        sketchProjectBuilder1.Section.DistanceTolerance = 0.01

        sketchProjectBuilder1.Section.ChainingTolerance = 0.0095

        sketchProjectBuilder1.Section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.CurvesAndPoints)

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        'Dim curves1(0) As NXOpen.IBaseCurve
        Dim curves1(0) As NXOpen.Curve

        curves1(0) = arc1
        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        sketchProjectBuilder1.Section.AllowSelfIntersection(True)

        sketchProjectBuilder1.Section.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = TIMDIEMTHUOCDUONGTRON(arc1) 'New NXOpen.Point3d(799.03842119118872, 274.72679560016746, 1151.1098255475474)

        sketchProjectBuilder1.Section.AddToSection(rules1, arc1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        sketchProjectBuilder1.Section.CleanMappingData()

        sketchProjectBuilder1.ProjectAsDumbFixedCurves = False

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchProjectBuilder1.Commit()


        For Each iObj As NXOpen.NXObject In ifeature1.GetEntities
            If TypeOf (iObj) Is NXOpen.Curve Then 'Ellipse
                arc2 = iObj
                Exit For
                'ElseIf TypeOf (iObj) Is CartesianCoordinateSystem Then
                '    tmpCs = iObj
            End If
        Next

        sketchProjectBuilder1.Destroy()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub chieudiem(ByVal point1 As NXOpen.Point, ByRef ipoint As NXOpen.Point, ByVal ifeature As NXOpen.Features.Feature)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Associative Curve->Project Curve...
        ' ----------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim sketchProjectBuilder1 As NXOpen.SketchProjectBuilder = Nothing
        sketchProjectBuilder1 = workPart.Sketches.CreateProjectBuilder(nullNXOpen_Features_Feature)

        sketchProjectBuilder1.Tolerance = 0.01

        sketchProjectBuilder1.Section.PrepareMappingData()

        sketchProjectBuilder1.Section.DistanceTolerance = 0.01

        sketchProjectBuilder1.Section.ChainingTolerance = 0.0095

        sketchProjectBuilder1.Section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.CurvesAndPoints)

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim points1(0) As NXOpen.Point
        'Dim point1 As NXOpen.Point = CType(workPart.Points.FindObject("ENTITY 2 4 1"), NXOpen.Point)

        points1(0) = point1
        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleCurveDumbFromPoints(points1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        sketchProjectBuilder1.Section.AllowSelfIntersection(True)

        sketchProjectBuilder1.Section.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = New NXOpen.Point3d(point1.Coordinates.X, point1.Coordinates.Y, point1.Coordinates.Z)
        sketchProjectBuilder1.Section.AddToSection(rules1, point1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        sketchProjectBuilder1.Section.CleanMappingData()

        sketchProjectBuilder1.Section.CleanMappingData()

        sketchProjectBuilder1.ProjectAsDumbFixedCurves = False

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = sketchProjectBuilder1.Commit()

        For Each iObj As NXOpen.NXObject In ifeature.GetEntities
            If TypeOf (iObj) Is NXOpen.Point Then
                If System.Math.Abs(CType(iObj, NXOpen.Point).Coordinates.X - point1.Coordinates.X) < 0.04 Then
                    ipoint = CType(iObj, NXOpen.Point)
                    ipoint.SetVisibility(SmartObject.VisibilityOption.Invisible)
                    Exit For
                End If
            End If
        Next

        sketchProjectBuilder1.Destroy()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub Netdut(ByVal ellipse1 As NXOpen.NXObject)
        Dim displayModification1 As NXOpen.DisplayModification = Nothing
        displayModification1 = theSession.DisplayManager.NewDisplayModification()

        displayModification1.ApplyToAllFaces = True

        displayModification1.ApplyToOwningParts = True

        displayModification1.NewFont = NXOpen.DisplayableObject.ObjectFont.DottedDashed

        displayModification1.NewWidth = NXOpen.DisplayableObject.ObjectWidth.One

        Dim objects1(0) As NXOpen.DisplayableObject

        objects1(0) = ellipse1
        displayModification1.Apply(objects1)

        displayModification1.Dispose()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub

    '------------------------Wave------------------------------
    Sub wave_link(ByVal iCurve As NXOpen.Curve, ByRef iCurveOut As Curve)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Associative Copy->WAVE Geometry Linker...
        ' ----------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing
        Dim waveLinkBuilder1 As NXOpen.Features.WaveLinkBuilder = Nothing
        waveLinkBuilder1 = workPart.BaseFeatures.CreateWaveLinkBuilder(nullNXOpen_Features_Feature)
        Dim waveDatumBuilder1 As NXOpen.Features.WaveDatumBuilder = Nothing
        waveDatumBuilder1 = waveLinkBuilder1.WaveDatumBuilder
        Dim compositeCurveBuilder1 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder1 = waveLinkBuilder1.CompositeCurveBuilder
        Dim waveSketchBuilder1 As NXOpen.Features.WaveSketchBuilder = Nothing
        waveSketchBuilder1 = waveLinkBuilder1.WaveSketchBuilder
        Dim waveRoutingBuilder1 As NXOpen.Features.WaveRoutingBuilder = Nothing
        waveRoutingBuilder1 = waveLinkBuilder1.WaveRoutingBuilder
        Dim wavePointBuilder1 As NXOpen.Features.WavePointBuilder = Nothing
        wavePointBuilder1 = waveLinkBuilder1.WavePointBuilder
        Dim extractFaceBuilder1 As NXOpen.Features.ExtractFaceBuilder = Nothing
        extractFaceBuilder1 = waveLinkBuilder1.ExtractFaceBuilder
        Dim mirrorBodyBuilder1 As NXOpen.Features.MirrorBodyBuilder = Nothing
        mirrorBodyBuilder1 = waveLinkBuilder1.MirrorBodyBuilder
        Dim curveFitData1 As NXOpen.GeometricUtilities.CurveFitData = Nothing
        curveFitData1 = compositeCurveBuilder1.CurveFitData
        curveFitData1.Tolerance = 0.01
        curveFitData1.AngleTolerance = 0.5
        Dim section1 As NXOpen.Section = Nothing
        section1 = compositeCurveBuilder1.Section
        section1.SetAllowRefCrvs(False)
        extractFaceBuilder1.FaceOption = NXOpen.Features.ExtractFaceBuilder.FaceOptionType.FaceChain
        extractFaceBuilder1.FaceOption = NXOpen.Features.ExtractFaceBuilder.FaceOptionType.FaceChain
        extractFaceBuilder1.AngleTolerance = 45.0
        waveLinkBuilder1.Associative = False
        waveDatumBuilder1.DisplayScale = 2.0
        extractFaceBuilder1.ParentPart = NXOpen.Features.ExtractFaceBuilder.ParentPartType.OtherPart
        mirrorBodyBuilder1.ParentPartType = NXOpen.Features.MirrorBodyBuilder.ParentPart.OtherPart
        compositeCurveBuilder1.Section.DistanceTolerance = 0.01
        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095
        compositeCurveBuilder1.Section.AngleTolerance = 0.5
        compositeCurveBuilder1.Section.DistanceTolerance = 0.01
        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095
        compositeCurveBuilder1.Associative = False
        compositeCurveBuilder1.MakePositionIndependent = False
        compositeCurveBuilder1.FixAtCurrentTimestamp = False
        compositeCurveBuilder1.HideOriginal = False
        compositeCurveBuilder1.InheritDisplayProperties = False
        compositeCurveBuilder1.JoinOption = NXOpen.Features.CompositeCurveBuilder.JoinMethod.No
        compositeCurveBuilder1.Tolerance = 0.01
        Dim section2 As NXOpen.Section = Nothing
        section2 = compositeCurveBuilder1.Section
        Dim curveFitData2 As NXOpen.GeometricUtilities.CurveFitData = Nothing
        curveFitData2 = compositeCurveBuilder1.CurveFitData
        extractFaceBuilder1.InheritMaterial = True
        waveLinkBuilder1.InheritMaterial = True
        mirrorBodyBuilder1.InheritMaterial = True
        section2.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.CurvesAndPoints)
        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()
        selectionIntentRuleOptions1.SetSelectedFromInactive(False)
        Dim features1(0) As NXOpen.Features.Feature
        Dim line1 As NXOpen.Curve = CType(iCurve, NXOpen.Curve)
        Dim nullNXOpen_Curve As NXOpen.Curve = Nothing
        Dim iCurveTangentRule As NXOpen.CurveTangentRule = Nothing

        Try

            iCurveTangentRule = workPart.ScRuleFactory.CreateRuleCurveTangent(line1, Nothing, False, 0.0095, 0.5)
        Catch ex As Exception
            iCurveTangentRule = displayPart.ScRuleFactory.CreateRuleCurveTangent(line1, Nothing, False, 0.0095, 0.5)
        End Try

        selectionIntentRuleOptions1.Dispose()
        section2.AllowSelfIntersection(False)
        section2.AllowDegenerateCurves(False)
        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = iCurveTangentRule
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing
        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", True, iCurve)
        section2.AddToSection(rules1, line1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)
        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = waveLinkBuilder1.Commit()
        Dim iOutCurve As NXOpen.Features.CompositeCurve = CType(nXObject1, NXOpen.Features.CompositeCurve)
        iCurveOut = CType(iOutCurve.GetEntities(0), NXOpen.Curve)
        waveLinkBuilder1.Destroy()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub Wave_Body(ByVal iBody As NXOpen.Body, ByRef iBodyWave As Body)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim TheUFsession As UFSession = UFSession.GetUFSession()
        theSession = Session.GetSession()
        lw = theSession.ListingWindow
        lw.Open()

        Try
            Dim iComponent As Component = iBody.OwningComponent
            'lw.WriteLine(iComponent.GetStringAttribute("DB_PART_NAME").ToString)
            Dim part1 As NXOpen.Part = CType(theSession.Parts.FindObject(iComponent.GetStringAttribute("DB_PART_NO").ToString), NXOpen.Part)
            Dim partLoadStatus1 As NXOpen.PartLoadStatus = Nothing
            partLoadStatus1 = part1.LoadThisPartFully()
            partLoadStatus1.Dispose()
        Catch ex As Exception

        End Try
        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing
        Dim waveLinkBuilder1 As NXOpen.Features.WaveLinkBuilder = Nothing
        waveLinkBuilder1 = workPart.BaseFeatures.CreateWaveLinkBuilder(nullNXOpen_Features_Feature)
        Dim compositeCurveBuilder1 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder1 = waveLinkBuilder1.CompositeCurveBuilder
        Dim extractFaceBuilder1 As NXOpen.Features.ExtractFaceBuilder = Nothing
        extractFaceBuilder1 = waveLinkBuilder1.ExtractFaceBuilder
        Dim mirrorBodyBuilder1 As NXOpen.Features.MirrorBodyBuilder = Nothing
        mirrorBodyBuilder1 = waveLinkBuilder1.MirrorBodyBuilder
        Dim curveFitData1 As NXOpen.GeometricUtilities.CurveFitData = Nothing
        curveFitData1 = compositeCurveBuilder1.CurveFitData
        curveFitData1.Tolerance = 0.01
        curveFitData1.AngleTolerance = 0.5
        extractFaceBuilder1.FaceOption = NXOpen.Features.ExtractFaceBuilder.FaceOptionType.FaceChain
        waveLinkBuilder1.Type = NXOpen.Features.WaveLinkBuilder.Types.BodyLink
        extractFaceBuilder1.FaceOption = NXOpen.Features.ExtractFaceBuilder.FaceOptionType.FaceChain
        extractFaceBuilder1.AngleTolerance = 45.0
        waveLinkBuilder1.Associative = False
        extractFaceBuilder1.ParentPart = NXOpen.Features.ExtractFaceBuilder.ParentPartType.OtherPart
        mirrorBodyBuilder1.ParentPartType = NXOpen.Features.MirrorBodyBuilder.ParentPart.OtherPart
        compositeCurveBuilder1.Section.DistanceTolerance = 0.01
        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095
        compositeCurveBuilder1.Section.AngleTolerance = 0.5
        compositeCurveBuilder1.Section.DistanceTolerance = 0.01
        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095
        extractFaceBuilder1.Associative = False
        extractFaceBuilder1.MakePositionIndependent = False
        extractFaceBuilder1.FixAtCurrentTimestamp = False
        extractFaceBuilder1.HideOriginal = False
        extractFaceBuilder1.InheritDisplayProperties = False
        Dim scCollector1 As NXOpen.ScCollector = Nothing
        scCollector1 = extractFaceBuilder1.ExtractBodyCollector
        extractFaceBuilder1.CopyThreads = True
        extractFaceBuilder1.FeatureOption = NXOpen.Features.ExtractFaceBuilder.FeatureOptionType.OneFeatureForAllBodies
        'waveLinkBuilder1.Type = NXOpen.Features.WaveLinkBuilder.Types.FaceLink
        'extractFaceBuilder1.Associative = False
        'extractFaceBuilder1.MakePositionIndependent = False
        'extractFaceBuilder1.FixAtCurrentTimestamp = False
        'extractFaceBuilder1.HideOriginal = False
        'extractFaceBuilder1.DeleteHoles = False
        'extractFaceBuilder1.InheritDisplayProperties = False

        'Dim selectDisplayableObjectList1 As NXOpen.SelectDisplayableObjectList = Nothing
        'selectDisplayableObjectList1 = extractFaceBuilder1.ObjectToExtract
        'waveLinkBuilder1.Type = NXOpen.Features.WaveLinkBuilder.Types.BodyLink
        'extractFaceBuilder1.Associative = False
        'extractFaceBuilder1.MakePositionIndependent = False
        'extractFaceBuilder1.FixAtCurrentTimestamp = False
        'extractFaceBuilder1.HideOriginal = False
        'extractFaceBuilder1.InheritDisplayProperties = False
        'Dim scCollector2 As NXOpen.ScCollector = Nothing
        'scCollector2 = extractFaceBuilder1.ExtractBodyCollector
        'extractFaceBuilder1.CopyThreads = True
        'extractFaceBuilder1.FeatureOption = NXOpen.Features.ExtractFaceBuilder.FeatureOptionType.OneFeatureForAllBodies

        Dim bodies1(0) As NXOpen.Body
        'Dim component1 As NXOpen.Assemblies.Component = CType(displayPart.ComponentAssembly.RootComponent.FindObject("COMPONENT @DB/NML29774031/AA 1"), NXOpen.Assemblies.Component)
        'Dim component2 As NXOpen.Assemblies.Component = CType(component1.FindObject("COMPONENT @DB/NML29774082/AA 1"), NXOpen.Assemblies.Component)
        'Dim body1 As NXOpen.Body = CType(component2.FindObject("PROTO#.Bodies|UNPARAMETERIZED_FEATURE(1)"), NXOpen.Body)
        bodies1(0) = iBody
        'bodies1(0) = body1
        Dim bodyDumbRule1 As NXOpen.BodyDumbRule = Nothing
        bodyDumbRule1 = workPart.ScRuleFactory.CreateRuleBodyDumb(bodies1, True)
        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = bodyDumbRule1
        scCollector1.ReplaceRules(rules1, False)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = waveLinkBuilder1.Commit()
        waveLinkBuilder1.Destroy()
        theSession.CleanUpFacetedFacesAndEdges()

        'Dim iFeature As Features.Feature
        'iFeature = nXObject1
        ''iJoinCurve = JoinCurve(iCurveIn)
        'Dim i As long = 0
        'For Each iObj As NXOpen.NXObject In iFeature.GetEntities
        '    If TypeOf (iObj) Is NXOpen.Body Then
        '        iBody = iObj
        '        'iBody.SetVisibility(SmartObject.VisibilityOption.Invisible)
        '        Exit For
        '    End If
        'Next
        iBodyWave = CType(workPart.Bodies.FindObject(nXObject1.JournalIdentifier.ToString), NXOpen.Body)

    End Sub
    Sub Blank_obj(ByVal body1 As NXOpen.NXObject)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        Dim objects1(0) As NXOpen.DisplayableObject

        objects1(0) = body1
        theSession.DisplayManager.BlankObjects(objects1)

        displayPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
    End Sub
    Sub Del_body(ByVal Bodytodel As NXOpen.Body)
        Try
            theUFSession.Obj.DeleteObject(Bodytodel.Tag)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try
    End Sub
    '------------------------Tao mat cat------------------------------
    'cat theo Csys
    Sub aCreSecZ(ByVal iSelCoor As NXOpen.CartesianCoordinateSystem)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim iMatrix As Matrix3x3 = iSelCoor.Orientation.Element
        ' ----------------------------------------------
        ' メニュー: 表示(V)->断面(S)->断面を編集(C)...
        ' ----------------------------------------------
        Dim dynamicSectionBuilder1 As NXOpen.Display.DynamicSectionBuilder = Nothing
        dynamicSectionBuilder1 = displayPart.DynamicSections.CreateSectionBuilder(displayPart.ModelingViews.WorkView)

        dynamicSectionBuilder1.DeferCurveUpdate = False

        dynamicSectionBuilder1.SeriesSpacing = 71.0
        dynamicSectionBuilder1.DefaultPlaneAxis = NXOpen.Display.DynamicSectionTypes.Axis.Z

        dynamicSectionBuilder1.ShowClip = False
        dynamicSectionBuilder1.CsysType = NXOpen.Display.DynamicSectionTypes.CoordinateSystem.Absolute
        dynamicSectionBuilder1.DefaultPlaneAxis = NXOpen.Display.DynamicSectionTypes.Axis.Z
        Dim modelingView1 As NXOpen.ModelingView = Nothing
        modelingView1 = dynamicSectionBuilder1.View
        dynamicSectionBuilder1.Type = NXOpen.Display.DynamicSectionTypes.Type.Box

        dynamicSectionBuilder1.DeferCurveUpdate = False
        Dim originX As NXOpen.Point3d = New NXOpen.Point3d(iSelCoor.Origin.X, iSelCoor.Origin.Y, iSelCoor.Origin.Z)
        modelingView1.SetOrigin(originX)

        Dim rotMatrix1 As NXOpen.Matrix3x3 = Nothing
        rotMatrix1.Xx = iMatrix.Xx
        rotMatrix1.Xy = iMatrix.Xy
        rotMatrix1.Xz = iMatrix.Xz
        rotMatrix1.Yx = iMatrix.Yx
        rotMatrix1.Yy = iMatrix.Yy
        rotMatrix1.Yz = iMatrix.Yz
        rotMatrix1.Zx = iMatrix.Zx
        rotMatrix1.Zy = iMatrix.Zy
        rotMatrix1.Zz = iMatrix.Zz
        dynamicSectionBuilder1.SetPlane(originX, originX, rotMatrix1)

        Dim iDistance As Double = 80

        Dim iY1 As New Point3d(iSelCoor.Origin.X + iMatrix.Yx * iDistance, iSelCoor.Origin.Y + iMatrix.Yy * iDistance, iSelCoor.Origin.Z + iMatrix.Yz * iDistance)
        Dim iY2 As New Point3d(2 * iSelCoor.Origin.X - iY1.X, 2 * iSelCoor.Origin.Y - iY1.Y, 2 * iSelCoor.Origin.Z - iY1.Z)
        Dim iX1 As New Point3d(iSelCoor.Origin.X + iMatrix.Xx * iDistance, iSelCoor.Origin.Y + iMatrix.Xy * iDistance, iSelCoor.Origin.Z + iMatrix.Xz * iDistance)
        Dim iX2 As New Point3d(2 * iSelCoor.Origin.X - iX1.X, 2 * iSelCoor.Origin.Y - iX1.Y, 2 * iSelCoor.Origin.Z - iX1.Z)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.X, NXOpen.Display.DynamicSectionTypes.ActivePlane.Primary)

        dynamicSectionBuilder1.SetOrigin(iX1)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.X, NXOpen.Display.DynamicSectionTypes.ActivePlane.Secondary)

        dynamicSectionBuilder1.SetOrigin(iX2)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.Y, NXOpen.Display.DynamicSectionTypes.ActivePlane.Primary)

        dynamicSectionBuilder1.SetOrigin(iY1)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.Y, NXOpen.Display.DynamicSectionTypes.ActivePlane.Secondary)

        dynamicSectionBuilder1.SetOrigin(iY2)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.Z, NXOpen.Display.DynamicSectionTypes.ActivePlane.Primary)
        Dim origin5 As NXOpen.Point3d = New NXOpen.Point3d(iSelCoor.Origin.X, iSelCoor.Origin.Y, iSelCoor.Origin.Z)
        dynamicSectionBuilder1.SetOrigin(origin5)

        dynamicSectionBuilder1.SetActivePlane(NXOpen.Display.DynamicSectionTypes.Axis.Z, NXOpen.Display.DynamicSectionTypes.ActivePlane.Secondary)

        dynamicSectionBuilder1.SetOrigin(origin5)
        dynamicSectionBuilder1.SaveCurves(Nothing)
        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = dynamicSectionBuilder1.Commit()
        Dim objects2() As NXOpen.NXObject
        objects2 = dynamicSectionBuilder1.GetCommittedObjects()
        dynamicSectionBuilder1.Destroy()
        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    'Cat theo Datum
    Sub CreateSect(ByVal iplane1 As DatumPlane)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        Dim dynamicSectionBuilder1 As NXOpen.Display.DynamicSectionBuilder = Nothing
        dynamicSectionBuilder1 = displayPart.DynamicSections.CreateSectionBuilder(displayPart.ModelingViews.WorkView)

        dynamicSectionBuilder1.DeferCurveUpdate = True

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(iplane1.Origin.X, iplane1.Origin.Y, iplane1.Origin.Z)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(iplane1.Normal.X, iplane1.Normal.Y, iplane1.Normal.Z)
        Dim plane1 As NXOpen.Plane = Nothing
        plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        dynamicSectionBuilder1.SeriesSpacing = 5.0

        dynamicSectionBuilder1.DefaultPlaneAxis = NXOpen.Display.DynamicSectionTypes.Axis.Z


        plane1.SetMethod(NXOpen.PlaneTypes.MethodType.Coincident)

        dynamicSectionBuilder1.ShowClip = True

        dynamicSectionBuilder1.ShowClip = True

        dynamicSectionBuilder1.CsysType = NXOpen.Display.DynamicSectionTypes.CoordinateSystem.Absolute

        dynamicSectionBuilder1.DefaultPlaneAxis = NXOpen.Display.DynamicSectionTypes.Axis.Z
        dynamicSectionBuilder1.CurveColor = displayPart.Colors.Find("Blue")

        dynamicSectionBuilder1.DeferCurveUpdate = False

        plane1.SetMethod(NXOpen.PlaneTypes.MethodType.Coincident)

        Dim geom1(0) As NXOpen.NXObject
        Dim datumPlane1 As NXOpen.DatumPlane = iplane1

        geom1(0) = datumPlane1
        plane1.SetGeometry(geom1)

        plane1.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)

        'plane1.Evaluate()

        Dim nullNXOpen_Plane As NXOpen.Plane = Nothing

        dynamicSectionBuilder1.SetAssociativePlane(nullNXOpen_Plane)

        Dim axisorigin1 As NXOpen.Point3d = origin1 'New NXOpen.Point3d(plane1.Origin.X, plane1.Origin.Y, plane1.Origin.Z)
        Dim origin2 As NXOpen.Point3d = origin1 'New NXOpen.Point3d(plane1.Origin.X, plane1.Origin.Y, plane1.Origin.Z)
        Dim rotationmatrix1 As NXOpen.Matrix3x3 = Nothing
        rotationmatrix1.Xx = plane1.Matrix.Xx '0.0
        rotationmatrix1.Xy = plane1.Matrix.Xy
        rotationmatrix1.Xz = plane1.Matrix.Xz
        rotationmatrix1.Yx = plane1.Matrix.Yx
        rotationmatrix1.Yy = plane1.Matrix.Yy
        rotationmatrix1.Yz = plane1.Matrix.Yz
        rotationmatrix1.Zx = plane1.Matrix.Zx
        rotationmatrix1.Zy = plane1.Matrix.Zy
        rotationmatrix1.Zz = plane1.Matrix.Zz
        dynamicSectionBuilder1.SetPlane(axisorigin1, origin2, rotationmatrix1)
        dynamicSectionBuilder1.SaveCurves(Nothing)

        'Dim datumPlane2 As NXOpen.DatumPlane = Nothing
        'datumPlane2 = dynamicSectionBuilder1.CreateDatumPlane()


        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = dynamicSectionBuilder1.Commit()

        Dim objects1() As NXOpen.NXObject
        objects1 = dynamicSectionBuilder1.GetCommittedObjects()


        Dim objects2(0) As NXOpen.NXObject
        objects2(0) = plane1
        Dim nErrs1 As Integer = Nothing
        nErrs1 = theSession.UpdateManager.AddToDeleteList(objects2)

        dynamicSectionBuilder1.Destroy()


        theSession.CleanUpFacetedFacesAndEdges()

        ' ----------------------------------------------
        '   Menu: View->Section->Clip Section
        ' ----------------------------------------------
        displayPart.ModelingViews.WorkView.DisplaySectioningToggle = False

    End Sub


    '------------------------Offset------------------------------
    Sub offset_3d(ByVal ibody As NXOpen.Body, ByVal iSpace As Double, ByVal Flip As Boolean, ByRef iBodyOffset As NXOpen.Body)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Offset/Scale->Offset Surface...
        ' ----------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim offsetSurfaceBuilder1 As NXOpen.Features.OffsetSurfaceBuilder = Nothing
        offsetSurfaceBuilder1 = workPart.Features.CreateOffsetSurfaceBuilder(nullNXOpen_Features_Feature)

        Dim unit1 As NXOpen.Unit = Nothing
        unit1 = offsetSurfaceBuilder1.Radius.Units

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits(iSpace, unit1)

        offsetSurfaceBuilder1.OutputOption = NXOpen.Features.OffsetSurfaceBuilder.OutputOptionType.OneFeatureForAllFaces

        offsetSurfaceBuilder1.Tolerance = 0.01

        offsetSurfaceBuilder1.PartialOption = True

        offsetSurfaceBuilder1.MaximumExcludedObjects = 10

        offsetSurfaceBuilder1.RemoveProblemVerticesOption = True

        offsetSurfaceBuilder1.Radius.SetFormula("5")

        offsetSurfaceBuilder1.ApproxOption = True

        offsetSurfaceBuilder1.SetOrientationMethod(NXOpen.Features.OffsetSurfaceBuilder.OrientationMethodType.UseExistingNormals)

        Dim nullNXOpen_ScCollector As NXOpen.ScCollector = Nothing

        Dim faceSetOffset1 As NXOpen.GeometricUtilities.FaceSetOffset = Nothing
        faceSetOffset1 = workPart.FaceSetOffsets.CreateFaceSet(iSpace, nullNXOpen_ScCollector, False, 0)

        offsetSurfaceBuilder1.FaceSets.Append(faceSetOffset1)

        Dim scCollector1 As NXOpen.ScCollector = Nothing
        scCollector1 = workPart.ScCollectors.CreateCollector()

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim body1 As NXOpen.Body = ibody 'CType(workPart.Bodies.FindObject("LINKED_BODY(0)"), NXOpen.Body)

        Dim faceBodyRule1 As NXOpen.FaceBodyRule = Nothing
        faceBodyRule1 = workPart.ScRuleFactory.CreateRuleFaceBody(body1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = faceBodyRule1
        scCollector1.ReplaceRules(rules1, False)

        faceSetOffset1.FaceCollector = scCollector1

        If Flip = True Then
            faceSetOffset1.ItemFlipFlag = True
        End If

        offsetSurfaceBuilder1.PartialOption = True

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = offsetSurfaceBuilder1.Commit()

        workPart.Expressions.Delete(expression1)

        Dim expression2 As NXOpen.Expression = offsetSurfaceBuilder1.Radius

        Dim expression3 As NXOpen.Expression = faceSetOffset1.Offset

        offsetSurfaceBuilder1.Destroy()

        iBodyOffset = CType(workPart.Bodies.FindObject(nXObject1.JournalIdentifier.ToString), NXOpen.Body)

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub extrude(ByVal line1 As Curve, ByVal component1 As Component, ByRef saigaimen As NXOpen.Body)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Design Feature->Extrude...
        ' ----------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim extrudeBuilder1 As NXOpen.Features.ExtrudeBuilder = Nothing
        extrudeBuilder1 = workPart.Features.CreateExtrudeBuilder(nullNXOpen_Features_Feature)

        Dim section1 As NXOpen.Section = Nothing
        section1 = workPart.Sections.CreateSection(0.0095, 0.01, 0.5)

        extrudeBuilder1.Section = section1

        extrudeBuilder1.AllowSelfIntersectingSection(True)

        Dim unit1 As NXOpen.Unit = Nothing
        unit1 = extrudeBuilder1.Draft.FrontDraftAngle.Units

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("2.00", unit1)

        extrudeBuilder1.DistanceTolerance = 0.01

        extrudeBuilder1.BooleanOperation.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Create

        Dim targetBodies1(0) As NXOpen.Body
        Dim nullNXOpen_Body As NXOpen.Body = Nothing

        targetBodies1(0) = nullNXOpen_Body
        extrudeBuilder1.BooleanOperation.SetTargetBodies(targetBodies1)

        extrudeBuilder1.Limits.StartExtend.Value.SetFormula("0")

        extrudeBuilder1.Limits.EndExtend.Value.SetFormula(extrude_space)

        extrudeBuilder1.BooleanOperation.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Create

        Dim targetBodies2(0) As NXOpen.Body
        targetBodies2(0) = nullNXOpen_Body
        extrudeBuilder1.BooleanOperation.SetTargetBodies(targetBodies2)

        extrudeBuilder1.Draft.FrontDraftAngle.SetFormula("2")

        extrudeBuilder1.Draft.BackDraftAngle.SetFormula("2")

        extrudeBuilder1.Offset.StartOffset.SetFormula("0")

        extrudeBuilder1.Offset.EndOffset.SetFormula("5")

        Dim smartVolumeProfileBuilder1 As NXOpen.GeometricUtilities.SmartVolumeProfileBuilder = Nothing
        smartVolumeProfileBuilder1 = extrudeBuilder1.SmartVolumeProfile

        smartVolumeProfileBuilder1.OpenProfileSmartVolumeOption = False

        smartVolumeProfileBuilder1.CloseProfileRule = NXOpen.GeometricUtilities.SmartVolumeProfileBuilder.CloseProfileRuleType.Fci

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(0.0, 0.0, 0.0)
        Dim vector1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 1.0)
        Dim direction1 As NXOpen.Direction = Nothing
        direction1 = workPart.Directions.CreateDirection(origin1, vector1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        extrudeBuilder1.Direction = direction1

        section1.DistanceTolerance = 0.01

        section1.ChainingTolerance = 0.0095

        section1.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

        Dim compositeCurveBuilder1 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder1 = workPart.Features.CreateCompositeCurveBuilder(nullNXOpen_Features_Feature)

        Dim section2 As NXOpen.Section = Nothing
        section2 = compositeCurveBuilder1.Section

        compositeCurveBuilder1.Associative = True

        compositeCurveBuilder1.ParentPart = NXOpen.Features.CompositeCurveBuilder.PartType.OtherPart

        compositeCurveBuilder1.AllowSelfIntersection = True

        section2.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

        section2.SetAllowRefCrvs(False)

        compositeCurveBuilder1.FixAtCurrentTimestamp = True

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim curves1(0) As NXOpen.IBaseCurve
        'Dim line1 As NXOpen.Line = CType(displayPart.Lines.FindObject("ENTITY 3 1 1"), NXOpen.Line)

        curves1(0) = line1
        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", True, line1) 'New NXOpen.Point3d(3181.3880376183652, -989.99999999999818, -0.00000000000011368683772161603)
        section2.AddToSection(rules1, line1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        Dim feature1 As NXOpen.Features.Feature = Nothing
        feature1 = compositeCurveBuilder1.CommitCreateOnTheFly()

        Dim waveLinkRepository1 As NXOpen.GeometricUtilities.WaveLinkRepository = Nothing
        waveLinkRepository1 = workPart.CreateWavelinkRepository()

        waveLinkRepository1.SetNonFeatureApplication(False)

        waveLinkRepository1.SetBuilder(extrudeBuilder1)

        Dim compositeCurve1 As NXOpen.Features.CompositeCurve = CType(feature1, NXOpen.Features.CompositeCurve)

        waveLinkRepository1.SetLink(compositeCurve1)

        Dim compositeCurveBuilder2 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder2 = workPart.Features.CreateCompositeCurveBuilder(compositeCurve1)

        compositeCurveBuilder2.Associative = False

        Dim feature2 As NXOpen.Features.Feature = Nothing
        feature2 = compositeCurveBuilder2.CommitCreateOnTheFly()

        compositeCurveBuilder2.Destroy()

        compositeCurveBuilder1.Destroy()

        Dim selectionIntentRuleOptions2 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions2 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions2.SetSelectedFromInactive(False)

        Dim features1(0) As NXOpen.Features.Feature
        Dim compositeCurve2 As NXOpen.Features.CompositeCurve = CType(feature2, NXOpen.Features.CompositeCurve)

        features1(0) = compositeCurve2
        Dim nullNXOpen_DisplayableObject As NXOpen.DisplayableObject = Nothing

        Dim curveFeatureRule1 As NXOpen.CurveFeatureRule = Nothing
        curveFeatureRule1 = workPart.ScRuleFactory.CreateRuleCurveFeature(features1, nullNXOpen_DisplayableObject, selectionIntentRuleOptions2)

        selectionIntentRuleOptions2.Dispose()
        section1.AllowSelfIntersection(True)

        section1.AllowDegenerateCurves(False)

        Dim rules2(0) As NXOpen.SelectionIntentRule
        rules2(0) = curveFeatureRule1
        Dim line2 As NXOpen.Line = CType(compositeCurve2.FindObject("CURVE 1"), NXOpen.Line)

        Dim helpPoint2 As NXOpen.Point3d = helpPoint1 'New NXOpen.Point3d(3181.3880376183652, -989.99999999999818, -0.00000000000011368683772161603)
        section1.AddToSection(rules2, line2, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint2, NXOpen.Section.Mode.Create, False)

        extrudeBuilder1.ParentFeatureInternal = False

        Dim feature3 As NXOpen.Features.Feature = Nothing
        feature3 = extrudeBuilder1.CommitFeature()

        Dim expression2 As NXOpen.Expression = extrudeBuilder1.Limits.StartExtend.Value

        Dim expression3 As NXOpen.Expression = extrudeBuilder1.Limits.EndExtend.Value

        extrudeBuilder1.Destroy()

        workPart.Expressions.Delete(expression1)

        waveLinkRepository1.Destroy()
        saigaimen = CType(workPart.Bodies.FindObject(feature3.JournalIdentifier.ToString), NXOpen.Body)
        theSession.CleanUpFacetedFacesAndEdges()

    End Sub
    Sub Flip_boolean(ByVal iMinface As Body, ByRef Flip As Boolean)
        ''''''thuat toan: Tim CS co diem goc <>(0,0,0) dau tien, lay huong Z cua CS
        ''''''Tu CS nay, tim diem tren Body gan voi CS nhat
        ''''''Tu diem nay, tim mat tren SheetBody chua diem (Vi Code tim Vec to phap tuyen chi tim giua mat va 1 diem tren mat)
        ''''''Tu mat va Diem o tren tim vec to phap tuyen cua mat
        ''''''So sanh huong phap tuyen va huong Z cuqa CS
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim displayPart As Part = theSession.Parts.Display
        Dim cp1(2) As Double


        ''''Tim Diem goc cua CS va tim BOdy de Tim Diem tren Body gan voi CS nhat
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




        Dim pt2(2) As Double

        ''''Tim Vector phap tuyen cua mat
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim pt(2) As Double
        Dim u1(2) As Double
        Dim v1(2) As Double
        Dim u2(2) As Double
        Dim v2(2) As Double
        Dim norm(2) As Double
        Dim radii(1) As Double
        Dim param(1) As Double
        theUFSession.Modl.AskFaceParm(iMinface.Tag, pt2, param, pt)   '195011
        theUFSession.Modl.AskFaceProps(iMinface.Tag, param, pt, u1, v1, u2, v2, norm, radii)
        'lw.WriteLine("huong phap tuyen: X: " & norm(0).ToString & " Y: " & norm(1).ToString & " Z: " & norm(2).ToString)
        'lw.WriteLine("huong phap tuyen: X: " & Int(norm(0)).ToString & " Y: " & Int(norm(1)).ToString & " Z: " & Int(norm(2)).ToString)
        '''' So sanh Vecto phap tuyen cua mat va huong Z cua CS (Cung huong/khac huong) True: Cung huong, False: khac huong
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim faceType As Integer
        Dim facePt(2) As Double
        Dim faceDir(2) As Double
        Dim bbox(5) As Double
        Dim faceRadius As Double
        Dim faceRadData As Double
        Dim normDirection As Integer




        'lw.WriteLine(iAngleMesure(iZCS(0), iZCS(1), iZCS(2), norm(0), norm(1), norm(2)))
        If C007_iAngleMesure(0, 1, 0, norm(0), norm(1), norm(2)) < 0 Then
            Flip = False
        Else
            Flip = True
        End If







    End Sub
    Function C007_iAngleMesure(ByVal X1 As Double, Y1 As Double, Z1 As Double, X2 As Double, Y2 As Double, Z2 As Double) As Double
        C007_iAngleMesure = (X1 * X2 + Y1 * Y2 + Z1 * Z2) / (Math.Sqrt(X1 ^ 2 + Y1 ^ 2 + Z1 ^ 2) * Math.Sqrt(X2 ^ 2 + Y2 ^ 2 + Z2 ^ 2))
        'lw.WriteLine(Math.Abs(C007_iAngleMesure))
    End Function
    'Sub Offset_nonSketch_Thickness(ByVal icurve() As Curve, ByVal k_cach As Double, ByRef Out_Curve() As Curve)

    '    Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
    '    Dim workPart As NXOpen.Part = theSession.Parts.Work
    '    Dim idem, idem1 As Integer

    '    Dim displayPart As NXOpen.Part = theSession.Parts.Display

    '    ' ----------------------------------------------
    '    '   Menu: Insert->Derived Curve->Offset...
    '    ' ----------------------------------------------
    '    Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

    '    Dim offsetCurveBuilder1 As NXOpen.Features.OffsetCurveBuilder = Nothing
    '    offsetCurveBuilder1 = workPart.Features.CreateOffsetCurveBuilder(nullNXOpen_Features_Feature)

    '    Dim unit1 As NXOpen.Unit = Nothing
    '    unit1 = offsetCurveBuilder1.OffsetDistance.Units

    '    Dim expression1 As NXOpen.Expression = Nothing
    '    expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

    '    offsetCurveBuilder1.CurveFitData.Tolerance = 0.01

    '    offsetCurveBuilder1.CurveFitData.AngleTolerance = 0.5

    '    offsetCurveBuilder1.OffsetDistance.SetFormula(k_cach)

    '    offsetCurveBuilder1.DraftHeight.SetFormula("5")

    '    offsetCurveBuilder1.DraftAngle.SetFormula("0")

    '    offsetCurveBuilder1.LawControl.Value.SetFormula("5")

    '    offsetCurveBuilder1.LawControl.StartValue.SetFormula("5")

    '    offsetCurveBuilder1.LawControl.EndValue.SetFormula("5")

    '    offsetCurveBuilder1.Offset3dDistance.SetFormula("5")

    '    offsetCurveBuilder1.InputCurvesOptions.InputCurveOption = NXOpen.GeometricUtilities.CurveOptions.InputCurve.Retain

    '    offsetCurveBuilder1.TrimMethod = NXOpen.Features.OffsetCurveBuilder.TrimOption.ExtendTangents

    '    offsetCurveBuilder1.LawControl.AlongSpineData.SetFeatureSpine(offsetCurveBuilder1.CurvesToOffset)

    '    Dim expression2 As NXOpen.Expression = Nothing
    '    expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

    '    offsetCurveBuilder1.CurvesToOffset.DistanceTolerance = 0.01

    '    offsetCurveBuilder1.CurvesToOffset.ChainingTolerance = 0.0095

    '    offsetCurveBuilder1.LawControl.AlongSpineData.Spine.DistanceTolerance = 0.01

    '    offsetCurveBuilder1.LawControl.AlongSpineData.Spine.ChainingTolerance = 0.0095

    '    offsetCurveBuilder1.LawControl.LawCurve.DistanceTolerance = 0.01

    '    offsetCurveBuilder1.LawControl.LawCurve.ChainingTolerance = 0.0095

    '    offsetCurveBuilder1.CurvesToOffset.AngleTolerance = 0.5

    '    offsetCurveBuilder1.LawControl.AlongSpineData.Spine.AngleTolerance = 0.5

    '    offsetCurveBuilder1.LawControl.LawCurve.AngleTolerance = 0.5

    '    Dim offsetdirection1 As NXOpen.Vector3d = Nothing
    '    Dim startpoint1 As NXOpen.Point3d = Nothing
    '    Try
    '        ' The Profile Cannot be Empty
    '        offsetCurveBuilder1.ComputeOffsetDirection(offsetdirection1, startpoint1)
    '    Catch ex As NXException
    '        ex.AssertErrorCode(671415)
    '    End Try

    '    offsetCurveBuilder1.CurvesToOffset.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

    '    Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
    '    selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

    '    selectionIntentRuleOptions1.SetSelectedFromInactive(False)


    '    '--------------------
    '    Dim curves1(UBound(icurve)) As NXOpen.IBaseCurve
    '    For i As Integer = 0 To UBound(icurve)
    '        curves1(i) = icurve(i)
    '    Next
    '    '--------------------


    '    Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
    '    curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

    '    selectionIntentRuleOptions1.Dispose()
    '    offsetCurveBuilder1.CurvesToOffset.AllowSelfIntersection(True)

    '    offsetCurveBuilder1.CurvesToOffset.AllowDegenerateCurves(False)

    '    Dim rules1(0) As NXOpen.SelectionIntentRule
    '    rules1(0) = curveDumbRule1
    '    Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

    '    Dim helpPoint1 As NXOpen.Point3d = timdiemdauZnho(icurve(0)) 'New NXOpen.Point3d(1350.7306541812175, -804.414662369558, 759.566269867803)
    '    offsetCurveBuilder1.CurvesToOffset.AddToSection(rules1, icurve(0), nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

    '    Dim offsetdirection2 As NXOpen.Vector3d = Nothing
    '    Dim startpoint2 As NXOpen.Point3d = Nothing
    '    offsetCurveBuilder1.ComputeOffsetDirection(offsetdirection2, startpoint2)

    '    offsetCurveBuilder1.ReverseDirection = False

    '    Dim nXObject1 As NXOpen.NXObject = Nothing
    '    nXObject1 = offsetCurveBuilder1.Commit()

    '    Dim expression3 As NXOpen.Expression = offsetCurveBuilder1.OffsetDistance

    '    offsetCurveBuilder1.Destroy()

    '    workPart.MeasureManager.SetPartTransientModification()

    '    workPart.Expressions.Delete(expression1)

    '    workPart.MeasureManager.ClearPartTransientModification()

    '    workPart.MeasureManager.SetPartTransientModification()

    '    workPart.Expressions.Delete(expression2)

    '    workPart.MeasureManager.ClearPartTransientModification()

    '    theSession.CleanUpFacetedFacesAndEdges()

    '    Dim offset_feature As NXOpen.Features.Feature
    '    offset_feature = CType(nXObject1, Feature)
    '    For Each obj0 As NXOpen.NXObject In offset_feature.GetEntities
    '        ReDim Preserve Out_Curve(idem)
    '        Out_Curve(idem) = CType(obj0, Curve)
    '        idem = idem + 1
    '    Next

    '    If timdiemdauZnho(Out_Curve(0)).Y < timdiemdauZnho(icurve(0)).Y Then
    '        Dim offsetCurveBuilder2 As NXOpen.Features.OffsetCurveBuilder = Nothing
    '        offsetCurveBuilder2 = workPart.Features.CreateOffsetCurveBuilder(CType(nXObject1, Feature))
    '        offsetCurveBuilder2.ReverseDirection = True

    '        Dim nXObject2 As NXOpen.NXObject = Nothing
    '        nXObject2 = offsetCurveBuilder2.Commit()
    '        offsetCurveBuilder2.Destroy()

    '        offset_feature = CType(nXObject2, Feature)
    '        For Each obj0 As NXOpen.NXObject In offset_feature.GetEntities
    '            ReDim Preserve Out_Curve(idem1)
    '            Out_Curve(idem1) = CType(obj0, Curve)
    '            idem1 = idem1 + 1
    '        Next
    '    End If
    'End Sub
    Sub Offset_nonSketch(ByVal icurve() As Curve, ByVal k_cach As Double, ByRef Out_Curve() As Curve)
        Dim iicurve1(), iicurve2() As Curve
        Dim del_feature, fture1, fture2 As Feature
        Call offset_curve(icurve, k_cach, True, iicurve1, fture1)
        Call offset_curve(icurve, k_cach, False, iicurve2, fture2)

        If timdiemdau("Z", True, iicurve1(0)).Y < timdiemdau("Z", True, iicurve2(0)).Y Then
            del_feature = fture1
            For i As Integer = 0 To UBound(iicurve2)
                ReDim Preserve Out_Curve(i)
                Out_Curve(i) = iicurve2(i)
            Next
        Else
            del_feature = fture2
            For i As Integer = 0 To UBound(iicurve1)
                ReDim Preserve Out_Curve(i)
                Out_Curve(i) = iicurve1(i)
            Next
        End If
        Try
            theUFSession.Obj.DeleteObject(del_feature.Tag)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try
    End Sub
    Sub Offset_nonSketch_draft(ByVal icurve() As Curve, ByVal k_cach As Double, ByRef Out_Curve() As Curve, ByRef Fture_todel As Feature)
        Dim iicurve1(), iicurve2() As Curve
        Dim del_feature, fture1, fture2 As Feature
        Call offset_curve(icurve, k_cach, True, iicurve1, fture1)
        Call offset_curve(icurve, k_cach, False, iicurve2, fture2)

        If timdiemdau("Z", True, iicurve1(0)).Y > timdiemdau("Z", True, iicurve2(0)).Y Then
            del_feature = fture1
            Fture_todel = fture2
            For i As Integer = 0 To UBound(iicurve2)
                ReDim Preserve Out_Curve(i)
                Out_Curve(i) = iicurve2(i)
            Next
        Else
            del_feature = fture2
            Fture_todel = fture1
            For i As Integer = 0 To UBound(iicurve1)
                ReDim Preserve Out_Curve(i)
                Out_Curve(i) = iicurve1(i)
            Next
        End If
        Try
            theUFSession.Obj.DeleteObject(del_feature.Tag)
        Catch ex As Exception
            lw.WriteLine(ex.ToString)
        End Try
    End Sub
    Sub offset_curve(ByVal icurve() As Curve, ByVal k_cach As Double, ByVal flip As Boolean, ByRef Out_Curve() As Curve, ByRef offset_feature As NXOpen.Features.Feature)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim idem, idem1 As Integer

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Derived Curve->Offset...
        ' ----------------------------------------------
        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim offsetCurveBuilder1 As NXOpen.Features.OffsetCurveBuilder = Nothing
        offsetCurveBuilder1 = workPart.Features.CreateOffsetCurveBuilder(nullNXOpen_Features_Feature)

        Dim unit1 As NXOpen.Unit = Nothing
        unit1 = offsetCurveBuilder1.OffsetDistance.Units

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        offsetCurveBuilder1.CurveFitData.Tolerance = 0.01

        offsetCurveBuilder1.CurveFitData.AngleTolerance = 0.5

        offsetCurveBuilder1.OffsetDistance.SetFormula(k_cach)

        offsetCurveBuilder1.DraftHeight.SetFormula("5")

        offsetCurveBuilder1.DraftAngle.SetFormula("0")

        offsetCurveBuilder1.LawControl.Value.SetFormula("5")

        offsetCurveBuilder1.LawControl.StartValue.SetFormula("5")

        offsetCurveBuilder1.LawControl.EndValue.SetFormula("5")

        offsetCurveBuilder1.Offset3dDistance.SetFormula("5")

        offsetCurveBuilder1.InputCurvesOptions.InputCurveOption = NXOpen.GeometricUtilities.CurveOptions.InputCurve.Retain

        offsetCurveBuilder1.TrimMethod = NXOpen.Features.OffsetCurveBuilder.TrimOption.ExtendTangents

        offsetCurveBuilder1.LawControl.AlongSpineData.SetFeatureSpine(offsetCurveBuilder1.CurvesToOffset)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        offsetCurveBuilder1.CurvesToOffset.DistanceTolerance = 0.01

        offsetCurveBuilder1.CurvesToOffset.ChainingTolerance = 0.0095

        offsetCurveBuilder1.LawControl.AlongSpineData.Spine.DistanceTolerance = 0.01

        offsetCurveBuilder1.LawControl.AlongSpineData.Spine.ChainingTolerance = 0.0095

        offsetCurveBuilder1.LawControl.LawCurve.DistanceTolerance = 0.01

        offsetCurveBuilder1.LawControl.LawCurve.ChainingTolerance = 0.0095

        offsetCurveBuilder1.CurvesToOffset.AngleTolerance = 0.5

        offsetCurveBuilder1.LawControl.AlongSpineData.Spine.AngleTolerance = 0.5

        offsetCurveBuilder1.LawControl.LawCurve.AngleTolerance = 0.5

        Dim offsetdirection1 As NXOpen.Vector3d = Nothing
        Dim startpoint1 As NXOpen.Point3d = Nothing
        Try
            ' The Profile Cannot be Empty
            offsetCurveBuilder1.ComputeOffsetDirection(offsetdirection1, startpoint1)
        Catch ex As NXException
            ex.AssertErrorCode(671415)
        End Try

        offsetCurveBuilder1.CurvesToOffset.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)


        '--------------------
        Dim curves1(UBound(icurve)) As NXOpen.IBaseCurve
        For i As Integer = 0 To UBound(icurve)
            curves1(i) = icurve(i)
        Next
        '--------------------


        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        offsetCurveBuilder1.CurvesToOffset.AllowSelfIntersection(True)

        offsetCurveBuilder1.CurvesToOffset.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", True, icurve(0)) 'New NXOpen.Point3d(1350.7306541812175, -804.414662369558, 759.566269867803)
        offsetCurveBuilder1.CurvesToOffset.AddToSection(rules1, icurve(0), nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        Dim offsetdirection2 As NXOpen.Vector3d = Nothing
        Dim startpoint2 As NXOpen.Point3d = Nothing
        offsetCurveBuilder1.ComputeOffsetDirection(offsetdirection2, startpoint2)

        offsetCurveBuilder1.ReverseDirection = flip

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = offsetCurveBuilder1.Commit()

        Dim expression3 As NXOpen.Expression = offsetCurveBuilder1.OffsetDistance

        offsetCurveBuilder1.Destroy()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression1)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression2)

        workPart.MeasureManager.ClearPartTransientModification()

        theSession.CleanUpFacetedFacesAndEdges()

        'Dim offset_feature As NXOpen.Features.Feature
        offset_feature = CType(nXObject1, Feature)
        For Each obj0 As NXOpen.NXObject In offset_feature.GetEntities
            ReDim Preserve Out_Curve(idem)
            Out_Curve(idem) = CType(obj0, Curve)
            idem = idem + 1
        Next
    End Sub
    Sub Join_curve_Group(ByVal myGroup As NXOpen.Group, ByVal line1 As Curve, ByRef feature_curve As Feature)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Associative Copy->Extract Geometry...
        ' ----------------------------------------------

        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim wavePointBuilder1 As NXOpen.Features.WavePointBuilder = Nothing
        wavePointBuilder1 = workPart.Features.CreateWavePointBuilder(nullNXOpen_Features_Feature)

        Dim waveDatumBuilder1 As NXOpen.Features.WaveDatumBuilder = Nothing
        waveDatumBuilder1 = workPart.Features.CreateWaveDatumBuilder(nullNXOpen_Features_Feature)

        Dim compositeCurveBuilder1 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder1 = workPart.Features.CreateCompositeCurveBuilder(nullNXOpen_Features_Feature)

        Dim extractFaceBuilder1 As NXOpen.Features.ExtractFaceBuilder = Nothing
        extractFaceBuilder1 = workPart.Features.CreateExtractFaceBuilder(nullNXOpen_Features_Feature)

        Dim mirrorBodyBuilder1 As NXOpen.Features.MirrorBodyBuilder = Nothing
        mirrorBodyBuilder1 = workPart.Features.CreateMirrorBodyBuilder(nullNXOpen_Features_Feature)

        Dim waveSketchBuilder1 As NXOpen.Features.WaveSketchBuilder = Nothing
        waveSketchBuilder1 = workPart.Features.CreateWaveSketchBuilder(nullNXOpen_Features_Feature)

        compositeCurveBuilder1.CurveFitData.Tolerance = 0.01

        compositeCurveBuilder1.CurveFitData.AngleTolerance = 0.5

        compositeCurveBuilder1.Section.SetAllowRefCrvs(False)

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.FixAtCurrentTimestamp = True

        compositeCurveBuilder1.JoinOption = NXOpen.Features.CompositeCurveBuilder.JoinMethod.Genernal

        compositeCurveBuilder1.CurveFitData.CurveJoinMethod = NXOpen.GeometricUtilities.CurveFitData.Join.General

        waveDatumBuilder1.ParentPart = NXOpen.Features.WaveDatumBuilder.ParentPartType.WorkPart

        wavePointBuilder1.ParentPart = NXOpen.Features.WavePointBuilder.ParentPartType.WorkPart

        extractFaceBuilder1.ParentPart = NXOpen.Features.ExtractFaceBuilder.ParentPartType.WorkPart

        mirrorBodyBuilder1.ParentPartType = NXOpen.Features.MirrorBodyBuilder.ParentPart.WorkPart

        compositeCurveBuilder1.ParentPart = NXOpen.Features.CompositeCurveBuilder.PartType.WorkPart

        waveSketchBuilder1.ParentPart = NXOpen.Features.WaveSketchBuilder.ParentPartType.WorkPart

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.JoinOption = NXOpen.Features.CompositeCurveBuilder.JoinMethod.Genernal

        compositeCurveBuilder1.Section.DistanceTolerance = 0.01

        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095

        compositeCurveBuilder1.Section.AngleTolerance = 0.5

        compositeCurveBuilder1.Section.DistanceTolerance = 0.01

        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.FixAtCurrentTimestamp = True

        compositeCurveBuilder1.HideOriginal = False

        compositeCurveBuilder1.InheritDisplayProperties = False

        extractFaceBuilder1.InheritMaterial = True

        mirrorBodyBuilder1.InheritMaterial = True

        compositeCurveBuilder1.Section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.CurvesAndPoints)




        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim groups1(0) As NXOpen.Group
        Dim group1 As NXOpen.Group = myGroup 'CType(workPart.FindObject("ENTITY 15 1 1"), NXOpen.Group)

        groups1(0) = group1
        Dim curveGroupRule1 As NXOpen.CurveGroupRule = Nothing
        curveGroupRule1 = workPart.ScRuleFactory.CreateRuleCurveGroup(groups1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        compositeCurveBuilder1.Section.AllowSelfIntersection(False)

        compositeCurveBuilder1.Section.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveGroupRule1
        'Dim line1 As NXOpen.Line = CType(workPart.Lines.FindObject("ENTITY 3 12 1"), NXOpen.Line)

        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", False, line1) 'New NXOpen.Point3d(1323.2537684293329, -829.61244017211709, 823.79785703297773)
        compositeCurveBuilder1.Section.AddToSection(rules1, line1, nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression2 As NXOpen.Expression = Nothing
        expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression3 As NXOpen.Expression = Nothing
        expression3 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim expression4 As NXOpen.Expression = Nothing
        expression4 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = compositeCurveBuilder1.Commit()

        feature_curve = CType(nXObject1, Feature)
        'Dim idem As Integer
        'For Each icu As Curve In feature_curve.GetEntities
        '    idem = idem + 1
        'Next
        'lw.WriteLine(idem)
        compositeCurveBuilder1.Destroy()

        waveDatumBuilder1.Destroy()

        wavePointBuilder1.Destroy()

        extractFaceBuilder1.Destroy()

        mirrorBodyBuilder1.Destroy()

        waveSketchBuilder1.Destroy()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression1)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression2)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression3)

        workPart.MeasureManager.ClearPartTransientModification()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression4)

        workPart.MeasureManager.ClearPartTransientModification()

        theSession.CleanUpFacetedFacesAndEdges()

    End Sub
    Sub Joint_curve_Arr(ByVal Curve_Arr() As Curve, ByRef Curve_feat As Feature)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Insert->Associative Copy->Extract Geometry...
        ' ----------------------------------------------
        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

        Dim wavePointBuilder1 As NXOpen.Features.WavePointBuilder = Nothing
        wavePointBuilder1 = workPart.Features.CreateWavePointBuilder(nullNXOpen_Features_Feature)

        Dim waveDatumBuilder1 As NXOpen.Features.WaveDatumBuilder = Nothing
        waveDatumBuilder1 = workPart.Features.CreateWaveDatumBuilder(nullNXOpen_Features_Feature)

        Dim compositeCurveBuilder1 As NXOpen.Features.CompositeCurveBuilder = Nothing
        compositeCurveBuilder1 = workPart.Features.CreateCompositeCurveBuilder(nullNXOpen_Features_Feature)

        Dim extractFaceBuilder1 As NXOpen.Features.ExtractFaceBuilder = Nothing
        extractFaceBuilder1 = workPart.Features.CreateExtractFaceBuilder(nullNXOpen_Features_Feature)

        Dim mirrorBodyBuilder1 As NXOpen.Features.MirrorBodyBuilder = Nothing
        mirrorBodyBuilder1 = workPart.Features.CreateMirrorBodyBuilder(nullNXOpen_Features_Feature)

        Dim waveSketchBuilder1 As NXOpen.Features.WaveSketchBuilder = Nothing
        waveSketchBuilder1 = workPart.Features.CreateWaveSketchBuilder(nullNXOpen_Features_Feature)

        compositeCurveBuilder1.CurveFitData.Tolerance = 0.01

        compositeCurveBuilder1.CurveFitData.AngleTolerance = 0.5

        compositeCurveBuilder1.Section.SetAllowRefCrvs(False)

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.FixAtCurrentTimestamp = True

        compositeCurveBuilder1.JoinOption = NXOpen.Features.CompositeCurveBuilder.JoinMethod.Quintic

        compositeCurveBuilder1.CurveFitData.CurveJoinMethod = NXOpen.GeometricUtilities.CurveFitData.Join.Quintic

        waveDatumBuilder1.ParentPart = NXOpen.Features.WaveDatumBuilder.ParentPartType.WorkPart

        wavePointBuilder1.ParentPart = NXOpen.Features.WavePointBuilder.ParentPartType.WorkPart

        extractFaceBuilder1.ParentPart = NXOpen.Features.ExtractFaceBuilder.ParentPartType.WorkPart

        mirrorBodyBuilder1.ParentPartType = NXOpen.Features.MirrorBodyBuilder.ParentPart.WorkPart

        compositeCurveBuilder1.ParentPart = NXOpen.Features.CompositeCurveBuilder.PartType.WorkPart

        waveSketchBuilder1.ParentPart = NXOpen.Features.WaveSketchBuilder.ParentPartType.WorkPart

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.JoinOption = NXOpen.Features.CompositeCurveBuilder.JoinMethod.Quintic

        compositeCurveBuilder1.Section.DistanceTolerance = 0.01

        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095

        compositeCurveBuilder1.Section.AngleTolerance = 0.5

        compositeCurveBuilder1.Section.DistanceTolerance = 0.01

        compositeCurveBuilder1.Section.ChainingTolerance = 0.0095

        compositeCurveBuilder1.Associative = False

        compositeCurveBuilder1.FixAtCurrentTimestamp = True

        compositeCurveBuilder1.HideOriginal = False

        compositeCurveBuilder1.InheritDisplayProperties = False

        extractFaceBuilder1.InheritMaterial = True

        mirrorBodyBuilder1.InheritMaterial = True

        compositeCurveBuilder1.Section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.CurvesAndPoints)

        Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
        selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()

        selectionIntentRuleOptions1.SetSelectedFromInactive(False)

        Dim curves1(UBound(Curve_Arr)) As NXOpen.IBaseCurve
        For i As Integer = 0 To UBound(Curve_Arr)
            Dim curve_1 As Curve = CType(Curve_Arr(i), Curve)
            curves1(0) = curve_1
        Next

        Dim curveDumbRule1 As NXOpen.CurveDumbRule = Nothing
        curveDumbRule1 = workPart.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)

        selectionIntentRuleOptions1.Dispose()
        compositeCurveBuilder1.Section.AllowSelfIntersection(False)

        compositeCurveBuilder1.Section.AllowDegenerateCurves(False)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = curveDumbRule1
        Dim nullNXOpen_NXObject As NXOpen.NXObject = Nothing

        Dim helpPoint1 As NXOpen.Point3d = timdiemdau("Z", True, Curve_Arr(0))
        compositeCurveBuilder1.Section.AddToSection(rules1, Curve_Arr(0), nullNXOpen_NXObject, nullNXOpen_NXObject, helpPoint1, NXOpen.Section.Mode.Create, False)

        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)

        Dim expression1 As NXOpen.Expression = Nothing
        expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = compositeCurveBuilder1.Commit()

        Curve_feat = CType(nXObject1, Feature)

        compositeCurveBuilder1.Destroy()

        waveDatumBuilder1.Destroy()

        wavePointBuilder1.Destroy()

        extractFaceBuilder1.Destroy()

        mirrorBodyBuilder1.Destroy()

        waveSketchBuilder1.Destroy()

        workPart.MeasureManager.SetPartTransientModification()

        workPart.Expressions.Delete(expression1)

        workPart.MeasureManager.ClearPartTransientModification()

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    '------------------------Component------------------------------
    Sub Blank(ByVal newComp As Assemblies.Component)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim topLevelComp As Component = Nothing
        For Each aComp As Component In theSession.Parts.Display.ComponentAssembly.MapComponentsFromSubassembly(newComp)
            topLevelComp = aComp
        Next

        Dim objects1(0) As NXOpen.DisplayableObject
        objects1(0) = topLevelComp
        theSession.DisplayManager.BlankObjects(objects1)

        displayPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
    End Sub
    Sub Unblank(ByVal newComp As Assemblies.Component)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim topLevelComp As Component = Nothing
        For Each aComp As Component In theSession.Parts.Display.ComponentAssembly.MapComponentsFromSubassembly(newComp)
            topLevelComp = aComp
        Next

        Dim objects1(0) As NXOpen.DisplayableObject
        objects1(0) = topLevelComp
        theSession.DisplayManager.UnblankObjects(objects1)

        displayPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
    End Sub
    '------------------------View------------------------------
    Sub View(ByVal iview As String)
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim modelingView1 As NXOpen.ModelingView = Nothing
        modelingView1 = displayPart.ModelingViews.WorkView
        If iview = "FR" Then
            modelingView1.Orient(NXOpen.View.Canned.Left, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "RR" Then
            modelingView1.Orient(NXOpen.View.Canned.Right, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "Left" Then
            modelingView1.Orient(NXOpen.View.Canned.Front, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "Right" Then
            modelingView1.Orient(NXOpen.View.Canned.Back, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "Top" Then
            modelingView1.Orient(NXOpen.View.Canned.Top, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "Bottom" Then
            modelingView1.Orient(NXOpen.View.Canned.Bottom, NXOpen.View.ScaleAdjustment.Current)
            displayPart.ModelingViews.WorkView.Fit()
        ElseIf iview = "ISO" Then
            Dim matrix1 As NXOpen.Matrix3x3 = Nothing
            matrix1.Xx = 0.707106781
            matrix1.Xy = -0.707106781
            matrix1.Xz = 0.0
            matrix1.Yx = 0.40824829
            matrix1.Yy = 0.40824829
            matrix1.Yz = 0.816496581
            matrix1.Zx = -0.577350269
            matrix1.Zy = -0.577350269
            matrix1.Zz = 0.577350269
            modelingView1.Orient(matrix1)

            displayPart.ModelingViews.WorkView.Fit()
        End If
    End Sub
    Sub zentaizu(ByVal ImgeFolder As String, ByVal Pic_name As String)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        ' メニュー: レンダリングスタイル(D)->静的ワイヤフレーム(W)
        ' ----------------------------------------------
        Try
            workPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.StaticWireframe
        Catch ex As Exception
            displayPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.StaticWireframe
        End Try

        If Pic_name = "zentaizu" Then
            Dim numberHidden1 As Long = Nothing
            numberHidden1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_POINTS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
            Try
                workPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)
            Catch ex As Exception
                displayPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.StaticWireframe
            End Try
        End If

        Call ChupAnh(ImgeFolder & Pic_name)

    End Sub
    Sub ChupAnh(ByVal iSavePath As String)
        Dim theSession As Session = Session.GetSession()
        Dim theUfSession As UFSession = UFSession.GetUFSession()
        If IsNothing(theSession.Parts.BaseWork) Then 'active part required
            Return
        End If
        Dim workPart As Part = theSession.Parts.Work
        Dim displayPart As Part = theSession.Parts.Display

        Dim tempScreenshot As String
        Dim tempLocation As String = IO.Path.GetTempPath
        tempScreenshot = IO.Path.Combine(tempLocation, "NXscreenshot.png")

        Try
            workPart.Views.Refresh()
        Catch ex As Exception
            displayPart.Views.Refresh()
        End Try

        Try
            ExpScrShotToImg(tempScreenshot)
        Catch ex As Exception
            Return
        End Try

        CropScreenshot(tempScreenshot, 2, iSavePath)

    End Sub
    Sub ExpScrShotToImg(ByVal fileImageInfo As String)
        Dim theSession As Session = Session.GetSession()
        Dim theUfSession As UFSession = UFSession.GetUFSession()
        Dim wcsVisible As Boolean = theSession.Parts.BaseDisplay.WCS.Visibility
        Dim triadVisible As Long = theSession.Preferences.ScreenVisualization.TriadVisibility
        Dim dispModelViewNames As Boolean = theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames
        Dim dispModelViewimgPlusBorders As Boolean = theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders

        theSession.Parts.BaseDisplay.WCS.Visibility = False 'turn off the WCS
        theSession.Preferences.ScreenVisualization.TriadVisibility = 0  'turn off triad
        theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders = False
        theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames = False


        Try
            theUfSession.Disp.CreateImage(fileImageInfo, UFDisp.ImageFormat.Png, UFDisp.BackgroundColor.White)
        Catch ex As Exception
            MsgBox(ex.Message & ControlChars.NewLine & "'" & fileImageInfo & "' can not be created")
            Throw New Exception("Screenshot can not be created")
        Finally
            theSession.Parts.BaseDisplay.WCS.Visibility = wcsVisible
            theSession.Preferences.ScreenVisualization.TriadVisibility = triadVisible
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders = dispModelViewimgPlusBorders
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames = dispModelViewNames
        End Try

        theSession.Parts.BaseDisplay.WCS.Visibility = True  'turn off the WCS
        theSession.Preferences.ScreenVisualization.TriadVisibility = 1  'turn off triad
        theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders = True 'turn off view imgPlusBorder
        'theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames = False  'turn off view name
    End Sub
    Sub CropScreenshot(ByVal fileImageInfo As String, ByVal imgPlusBorderWidth As Long, ByVal iSavePath As String)
        Dim theSession As Session = Session.GetSession()
        Dim theUfSession As UFSession = UFSession.GetUFSession()
        Dim imgInput As New Bitmap(fileImageInfo)
        Dim imgPixelMinX As Long = imgInput.Width
        Dim imgPixelMinY As Long = imgInput.Height
        Dim imgPixelMaxX As Long = 0
        Dim imgPixelMaxY As Long = 0
        Dim bckgrndColor As Color = Color.White
        Dim boderColor As Color = Color.White
        Dim c As Color
        Dim i As Long = 0

        For y As Long = 0 To imgInput.Height - 1
            For x As Long = 0 To imgInput.Width - 1
                If imgInput.GetPixel(x, y).ToArgb <> bckgrndColor.ToArgb Then
                    If x < imgPixelMinX Then
                        imgPixelMinX = x
                    ElseIf x + 1 > imgPixelMaxX Then
                        imgPixelMaxX = x + 1
                    End If
                    If y < imgPixelMinY Then
                        imgPixelMinY = y
                    ElseIf y + 1 > imgPixelMaxY Then
                        imgPixelMaxY = y + 1
                    End If
                End If
            Next
        Next
        Dim rect As New Rectangle
        Dim croppedZone As Bitmap
        Try
            rect = New Rectangle(imgPixelMinX - 5, imgPixelMinY - 5, imgPixelMaxX - imgPixelMinX + 5, imgPixelMaxY - imgPixelMinY + 5)
            croppedZone = imgInput.Clone(rect, imgInput.PixelFormat)
        Catch
            rect = New Rectangle(imgPixelMinX, imgPixelMinY, imgPixelMaxX - imgPixelMinX, imgPixelMaxY - imgPixelMinY)
            croppedZone = imgInput.Clone(rect, imgInput.PixelFormat)
        End Try
        Dim imgPlusBorder As New Bitmap(croppedZone.Width + (2 * imgPlusBorderWidth), croppedZone.Height + (2 * imgPlusBorderWidth), croppedZone.PixelFormat)
        Dim gr As Graphics = Graphics.FromImage(imgPlusBorder)
        Using myBrush As Brush = New SolidBrush(boderColor)
            gr.FillRectangle(myBrush, 0, 0, imgPlusBorder.Width, imgPlusBorder.Height)
        End Using
        Dim xImage As Long = imgPlusBorder.Width - croppedZone.Width - imgPlusBorderWidth
        Dim yImage As Long = imgPlusBorder.Height - croppedZone.Height - imgPlusBorderWidth

        gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
        gr.DrawImage(croppedZone, New Drawing.Point(xImage, yImage))
        imgInput.Dispose()
        croppedZone.Dispose()
        gr.Dispose()

        Dim iTextName As String
        imgPlusBorder.Save(iSavePath, Imaging.ImageFormat.Png)

    End Sub
    Sub iResetFolder(ByVal iFolderPath As String)
        If Not Directory.Exists(iFolderPath) Then
            System.IO.Directory.CreateDirectory(iFolderPath)
            Return
        End If

        Dim files() As String
        files = Directory.GetFileSystemEntries(iFolderPath)
        For Each element As String In files
            If (Not Directory.Exists(element)) Then
                File.Delete(Path.Combine(iFolderPath, Path.GetFileName(element)))
            End If
        Next
    End Sub
    Sub dodayduong(ByVal proc As Boolean)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Preferences->Visualization...
        ' ----------------------------------------------

        Dim studioMaterialDirectory1 As String = Nothing
        studioMaterialDirectory1 = theSession.Preferences.AppearanceManagementPref.GetDirectoryOfStudioMaterial()

        Dim textureDir1 As String = Nothing
        textureDir1 = theSession.Preferences.AppearanceManagementPref.GetDirectoryOfTexture()

        Dim showSeeThrough1 As Boolean = Nothing
        showSeeThrough1 = theSession.Preferences.AppearanceManagementPref.DispSeeThru

        Dim showSceneLable1 As Boolean = Nothing
        showSceneLable1 = theSession.Preferences.AppearanceManagementPref.DispSceneLabels

        Dim excludeFromSelection1 As Boolean = Nothing
        excludeFromSelection1 = theSession.Preferences.AppearanceManagementPref.ExcludeFromSelection

        Dim designatorPrefix1 As String = Nothing
        designatorPrefix1 = workPart.Preferences.AppearanceMgmtPreference.PrefixAppearanceSchemes

        Dim designatorPrefix2 As String = Nothing
        designatorPrefix2 = workPart.Preferences.AppearanceMgmtPreference.PrefixAppearanceDesignators

        ' ----------------------------------------------
        '   Dialog Begin Visualization Preferences
        ' ----------------------------------------------
        workPart.Preferences.LineVisualization.ShowWidths = proc
        If proc = False Then
            Exit Sub
        End If
        Dim pixelwidths1(8) As Integer
        pixelwidths1(0) = 4
        pixelwidths1(1) = 4
        pixelwidths1(2) = 4
        pixelwidths1(3) = 4
        pixelwidths1(4) = 4
        pixelwidths1(5) = 4
        pixelwidths1(6) = 4
        pixelwidths1(7) = 4
        pixelwidths1(8) = 4
        workPart.Preferences.LineVisualization.SetPixelWidths(pixelwidths1)

        Dim studioMaterialDirectory2 As String = Nothing
        studioMaterialDirectory2 = theSession.Preferences.AppearanceManagementPref.GetDirectoryOfStudioMaterial()

        theSession.Preferences.AppearanceManagementPref.SetDirectoryOfStudioMaterial(studioMaterialDirectory1)

        Dim textureDir2 As String = Nothing
        textureDir2 = theSession.Preferences.AppearanceManagementPref.GetDirectoryOfTexture()

        workPart.Preferences.AppearanceMgmtPreference.PrefixAppearanceSchemes = designatorPrefix1

        workPart.Preferences.AppearanceMgmtPreference.PrefixAppearanceDesignators = designatorPrefix2

        theSession.Preferences.AppearanceManagementPref.DispSeeThru = False

        theSession.Preferences.AppearanceManagementPref.DispSceneLabels = False

        theSession.Preferences.AppearanceManagementPref.ExcludeFromSelection = False

        theSession.CleanUpFacetedFacesAndEdges()
    End Sub
    Sub Fit_to_PMI()

        If IsNothing(theSession.Parts.Work) Then
            'active part required
            Return
        End If

        Dim workPart As Part = theSession.Parts.Work
        'Dim currentScale As Double = displayPart.ModelingViews.WorkView.Scale

        'Dim iPMISize As Double = 3 / currentScale

        'If currentScale > 1 Then
        '    iPMISize = 5
        'End If
        Dim objects1(0) As DisplayableObject
        Dim i As Integer = 0
        For Each myPmi As Annotations.Pmi In workPart.PmiManager.Pmis
            Dim dispInst() As Annotations.Annotation = myPmi.GetDisplayInstances
            For Each thisDispInst As Annotations.Annotation In dispInst
                ReDim Preserve objects1(i)
                objects1(i) = thisDispInst
                i = i + 1
                Try

                    'Dim pmiLabel1 As Annotations.PmiLabel = CType(thisDispInst, Annotations.PmiLabel)
                    'ReDim Preserve objects1(i)
                    'objects1(i) = pmiLabel1
                    'i = i + 1
                    'Dim editSettingsBuilder1 As Annotations.EditSettingsBuilder
                    'editSettingsBuilder1 = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects1)
                    'Dim editsettingsbuilders1(0) As Drafting.BaseEditSettingsBuilder
                    'editsettingsbuilders1(0) = editSettingsBuilder1
                    'workPart.SettingsManager.ProcessForMultipleObjectsSettings(editsettingsbuilders1)

                    'editSettingsBuilder1.AnnotationStyle.LetteringStyle.GeneralTextAspectRatio = 0.6
                    'editSettingsBuilder1.AnnotationStyle.LetteringStyle.GeneralTextLineSpaceFactor = 0.5
                    'editSettingsBuilder1.AnnotationStyle.LetteringStyle.GeneralTextSize = iPMISize

                    '    Dim nXObject1 As NXObject
                    '    nXObject1 = editSettingsBuilder1.Commit()
                    '    editSettingsBuilder1.Destroy()
                Catch
                End Try
            Next
        Next
        Try
            workPart.ModelingViews.WorkView.Regenerate()
            workPart.ModelingViews.WorkView.UpdateDisplay()
        Catch ex As Exception
            displayPart.ModelingViews.WorkView.Regenerate()
            displayPart.ModelingViews.WorkView.UpdateDisplay()
        End Try

        Try
            workPart.ModelingViews.WorkView.FitToObjects(objects1)
            workPart.ModelingViews.WorkView.SetScale(0.5)
        Catch ex As Exception
            displayPart.ModelingViews.WorkView.FitToObjects(objects1)
            displayPart.ModelingViews.WorkView.SetScale(0.5)
        End Try
    End Sub
    Sub Xoay_Datum(ByVal point1 As NXOpen.Point, ByVal iPlane1 As NXOpen.DatumPlane)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        ' ----------------------------------------------
        '   Menu: Snap View
        ' ----------------------------------------------

        Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(point1.Coordinates.X, point1.Coordinates.Y, point1.Coordinates.Z)
        Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(Math.Abs(iPlane1.Normal.X), Math.Abs(iPlane1.Normal.Y), Math.Abs(iPlane1.Normal.Z))

        'Dim origin1 As NXOpen.Point3d = New NXOpen.Point3d(iPlane1.Origin.X, iPlane1.Origin.Y, iPlane1.Origin.Z)
        'Dim normal1 As NXOpen.Vector3d = New NXOpen.Vector3d(iPlane1.Normal.X, iPlane1.Normal.Y, iPlane1.Normal.Z)
        Dim plane1 As NXOpen.Plane = Nothing

        Try
            plane1 = workPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Catch ex As Exception
            plane1 = displayPart.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        End Try

        Dim axisorigin1 As NXOpen.Point3d = origin1 'New NXOpen.Point3d(plane1.Origin.X, plane1.Origin.Y, plane1.Origin.Z)
        Dim origin2 As NXOpen.Point3d = origin1 'New NXOpen.Point3d(plane1.Origin.X, plane1.Origin.Y, plane1.Origin.Z)

        Dim rotationmatrix1 As NXOpen.Matrix3x3 = Nothing
        rotationmatrix1.Yx = plane1.Matrix.Xx * -1
        rotationmatrix1.Yy = plane1.Matrix.Xy * -1
        rotationmatrix1.Yz = plane1.Matrix.Xz * -1
        rotationmatrix1.Xx = plane1.Matrix.Yx * -1
        rotationmatrix1.Xy = plane1.Matrix.Yy * -1
        rotationmatrix1.Xz = plane1.Matrix.Yz * -1
        rotationmatrix1.Zx = plane1.Matrix.Zx * -1
        rotationmatrix1.Zy = plane1.Matrix.Zy * -1
        rotationmatrix1.Zz = plane1.Matrix.Zz * -1

        Try
            workPart.ModelingViews.WorkView.Orient(rotationmatrix1)
        Catch ex As Exception
            displayPart.ModelingViews.WorkView.Orient(rotationmatrix1)
        End Try
    End Sub

    Sub demhieuqua()
        ''''''''''''''''''''''''''đếm hiệu quả''''''''''''''''''''''''''''''''''''
        Dim EIWToolLibrary = System.Reflection.Assembly.LoadFile("D:\nxcad\scratch\VehiclePlanMacro\VPG-Sysytem.dll")
        Dim VPGToolList = System.Activator.CreateInstance(EIWToolLibrary.GetExportedTypes(0))
        Try
            VPGToolList.VPG_System_Check()
        Catch ex As Exception
            MsgBox("VPG-System.dllがありません", vbExclamation)
            Exit Sub
        End Try

        VPGToolList.aCountAAA("@@@")

    End Sub
#End Region


End Class
