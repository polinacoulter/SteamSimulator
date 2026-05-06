Attribute VB_Name = "GraphModule"
Declare Function AG3DStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nDepth&, ByVal nXGap&, ByVal nZGap&) As Long
Declare Function AGAmp Lib "GSWAG32.DLL" (ByVal nPts&, ByVal nGroup&, fAmp#) As Long
Declare Function AGAmpError Lib "GSWAG32.DLL" (ByVal nPts&, ByVal nGroup&, fAmp#) As Long
Declare Function AGAux Lib "GSWAG32.DLL" (ByVal nSize&, nAux&) As Long
Declare Function AGCageStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nClrWall&, ByVal nClrSide&) As Long
Declare Function AGClose Lib "GSWAG32.DLL" () As Long
Declare Function AGClr Lib "GSWAG32.DLL" (ByVal nGroup&, nClr&) As Long
Declare Function AGCurveStyle Lib "GSWAG32.DLL" (ByVal nType&, ByVal nOrder&, ByVal nSTEPs&) As Long
Declare Function AGDataZ Lib "GSWAG32.DLL" (ByVal nPts&, fDataZ#) As Long
Declare Function AGDist Lib "GSWAG32.DLL" (ByVal nPts&, nDist#) As Long
Declare Function AGDistError Lib "GSWAG32.DLL" (ByVal nPts&, nDist#) As Long
Declare Function AGErrorBar Lib "GSWAG32.DLL" (ByVal nSel&, ByVal nSymStyle&, ByVal nClr&, ByVal nErrSrc&, ByVal fValue#) As Long
Declare Function AGFGColor Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nClr&) As Long
Declare Function AGFFT Lib "GSWAG32.DLL" (ByVal nPts&, fData#, ByVal nMode&) As Long
Declare Function AGFontStyle Lib "GSWAG32.DLL" (ByVal nUse&, ByVal nFamily&, ByVal nAttribs&, ByVal nSize&) As Long
Declare Function AGGraphBG Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nClr&) As Long
Declare Function AGGridStyle Lib "GSWAG32.DLL" (ByVal nSel&, ByVal nStyleMaj&, ByVal nStyleMin&) As Long
Declare Function AGInfo Lib "GSWAG32.DLL" (ByVal nIndex&) As Double
Declare Function AGLegendStyle Lib "GSWAG32.DLL" (ByVal nVert&, ByVal nHoriz&, ByVal nSize&, ByVal nClr&, ByVal nMode&) As Long
Declare Function AGOpen Lib "GSWAG32.DLL" () As Long
Declare Function AGPatt Lib "GSWAG32.DLL" (ByVal nGroup&, nPatt&) As Long
Declare Function AGRefresh3D Lib "GSWAG32.DLL" (ByVal nMode&) As Long
Declare Function AGReset Lib "GSWAG32.DLL" () As Long
Declare Function AGSetPerspective Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nRot&, ByVal nElev&, ByVal nEyePos&) As Long
Declare Function AGShow Lib "GSWAG32.DLL" (ByVal nGType&, ByVal nStyle&, ByVal nStats&) As Long
Declare Function AGSurfaceClr Lib "GSWAG32.DLL" (ByVal nClrMin&, ByVal nClrMax&, ByVal nClrSide&) As Long
Declare Function AGSym Lib "GSWAG32.DLL" (ByVal nGroup&, nSym&) As Long
Declare Function AGTimeGraph Lib "GSWAG32.DLL" (ByVal nPts&, ByVal nGroup&, ByVal fDataMax#, ByVal fDataMin#, ByVal nStyle&) As Long
Declare Function AGTimeUpdate Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nGroup&, fData#) As Long
Declare Function AGTitleBG Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nClr&) As Long
Declare Function AGTitleG Lib "GSWAG32.DLL" (ByVal szTitle$) As Long
Declare Function AGTitleX Lib "GSWAG32.DLL" (ByVal szTitle$) As Long
Declare Function AGTitleY Lib "GSWAG32.DLL" (ByVal szTitle$) As Long
Declare Function AGTitleYR Lib "GSWAG32.DLL" (ByVal szTitle$) As Long
Declare Function AGXAxisStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nTicks&, ByVal nLabeEvery&, ByVal fMax#, ByVal fMin#) As Long
Declare Function AGYAxisStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nTicks&, ByVal nLabeEvery&, ByVal fMax#, ByVal fMin#) As Long
Declare Function AGYRAxisStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nTicks&, ByVal nLabeEvery&, ByVal fMax#, ByVal fMin#) As Long
Declare Function AGZAxisStyle Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nTicks&, ByVal nLabeEvery&, ByVal fMax#, ByVal fMin#) As Long
Declare Function GSArc Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fRad#, ByVal fAng1#, ByVal fAng2#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSArea Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fHt#, ByVal nMode&, ByVal nGroup&) As Long
Declare Function GSArea3D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fHt#, ByVal fDepth#, ByVal fAng#, ByVal nMode&) As Long
Declare Function GSArrow Lib "GSWDLL32.DLL" (ByVal fxA#, ByVal fya#, ByVal fxB#, ByVal fyB#, ByVal fHeadLen#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSAxis Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fLen#, ByVal fTickLen#, ByVal nMajDivs&, ByVal nMinDivs&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSBar2D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fSpace#, ByVal fStackHt#, ByVal nMode&, ByVal nGroup&) As Long
Declare Function GSBar3D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fSpace#, ByVal fStackHt#, ByVal fDepth#, ByVal fAng#, ByVal nMode&, ByVal nGroup&) As Long
Declare Function GSBox2D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal nPatt&, ByVal nClr&) As Long
Declare Function GSBox3D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal fDepth#, ByVal fAng#, ByVal nPatt&, ByVal nClr1&, ByVal nClr2&) As Long
Declare Function GSBoxWhisker Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fSpace#, ByVal nMode&) As Long
Declare Function GSBubbleChart Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal nMode&) As Long
Declare Function GSCage3D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fxLen#, ByVal fyLen#, ByVal fzLen#, ByVal fAng#, ByVal fThick#, ByVal nxGrid&, ByVal nyGrid&, ByVal nzGrid&, ByVal nMode&, ByVal nClr1&, ByVal nClr2&) As Long
Declare Function GSCircle Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fRad#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSClearView Lib "GSWDLL32.DLL" (ByVal nMode&) As Long
Declare Function GSClipRead Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal nFormat&, ByVal nMode&) As Long
Declare Function GSClipWrite Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal nFormat&, ByVal nMode&) As Long
Declare Function GSClosePrn Lib "GSWDLL32.DLL" () As Long
Declare Function GSCloseServer Lib "GSWDLL32.DLL" () As Long
Declare Function GSCloseView Lib "GSWDLL32.DLL" (ByVal nWin&, ByVal nView&, ByVal nMode&) As Long
Declare Function GSCloseWin Lib "GSWDLL32.DLL" (ByVal nWin&) As Long
Declare Function GSCurveFit Lib "GSWDLL32.DLL" (ByVal nType&, ByVal nOrder&, ByVal nSTEPs&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSDataAmp Lib "GSWDLL32.DLL" (ByVal nPts&, ByVal nGroup&, fAmp#) As Long
Declare Function GSDataAmpErr Lib "GSWDLL32.DLL" (ByVal nP&, ByVal nG&, fAmpErr#) As Long
Declare Function GSDataAux Lib "GSWDLL32.DLL" (ByVal nPts&, nAux&) As Long
Declare Function GSDataClr Lib "GSWDLL32.DLL" (ByVal nPts&, nClr&) As Long
Declare Function GSDataDim Lib "GSWDLL32.DLL" (ByVal nPts&, ByVal nGroup&) As Long
Declare Function GSDataDist Lib "GSWDLL32.DLL" (ByVal nPts&, fDist#) As Long
Declare Function GSDataDistErr Lib "GSWDLL32.DLL" (ByVal nP&, fDistErr#) As Long
Declare Function GSDataGetAmp Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nGroup&) As Double
Declare Function GSDataGetAmpErr Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nGroup&, ByVal nMode&) As Double
Declare Function GSDataGetAux Lib "GSWDLL32.DLL" (ByVal nPt&) As Long
Declare Function GSDataGetClr Lib "GSWDLL32.DLL" (ByVal nPt&) As Long
Declare Function GSDataGetDist Lib "GSWDLL32.DLL" (ByVal nPt&) As Double
Declare Function GSDataGetDistErr Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nMode&) As Double
Declare Function GSDataGetPatt Lib "GSWDLL32.DLL" (ByVal nPt&) As Long
Declare Function GSDataGetSym Lib "GSWDLL32.DLL" (ByVal nPt&) As Long
Declare Function GSDataGetZ Lib "GSWDLL32.DLL" (ByVal nPt&) As Double
Declare Function GSDataPatt Lib "GSWDLL32.DLL" (ByVal nPts&, nPatt&) As Long
Declare Function GSDataRange Lib "GSWDLL32.DLL" (ByVal nFirst&, ByVal nLast&) As Long
Declare Function GSDataReset Lib "GSWDLL32.DLL" () As Long
Declare Function GSDataScale Lib "GSWDLL32.DLL" (ByVal fScale#) As Long
Declare Function GSDataStoreAmp Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nGroup&, ByVal fAmp#) As Long
Declare Function GSDataStoreAmpErr Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nGroup&, ByVal fErrP#, ByVal fErrM#) As Long
Declare Function GSDataStoreAux Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nAux&) As Long
Declare Function GSDataStoreClr Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nClr&) As Long
Declare Function GSDataStoreDist Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal fDist#) As Long
Declare Function GSDataStoreDistErr Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal fErrP#, ByVal fErrM#) As Long
Declare Function GSDataStorePatt Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nPatt&) As Long
Declare Function GSDataStoreSym Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal nSym&) As Long
Declare Function GSDataStoreZ Lib "GSWDLL32.DLL" (ByVal nPt&, ByVal fDataZ#) As Long
Declare Function GSDataSym Lib "GSWDLL32.DLL" (ByVal nPts&, nSym&) As Long
Declare Function GSDataTrans Lib "GSWDLL32.DLL" (ByVal nPts&, ByVal nGroup&, fA#, fD#, nPatt&, nSymbol&, nAux&, nClr&) As Long
Declare Function GSDataZ Lib "GSWDLL32.DLL" (ByVal nPts&, fDataZ#) As Long
Declare Function GSDefPatt Lib "GSWDLL32.DLL" (ByVal nBitmap&, wBitmap&) As Long
Declare Function GSEllipse Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal fxA#, ByVal fya#, ByVal fxB#, ByVal fyB#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSErrorBar Lib "GSWDLL32.DLL" (ByVal nSel&, ByVal nSymStyle&, ByVal nClr&, ByVal nErrSrc&, ByVal fValue#, ByVal fOff#) As Long
Declare Function GSFixPos Lib "GSWDLL32.DLL" (ByVal fx#, ByVal fy#) As Long
Declare Function GSGantt Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal nMode&, ByVal nGroup&) As Long
Declare Function GSGetACos Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetALog10 Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetALogE Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetASin Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetATan Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetAXExt Lib "GSWDLL32.DLL" (ByVal hWnd&) As Double
Declare Function GSGetAYExt Lib "GSWDLL32.DLL" (ByVal hWnd&) As Double
Declare Function GSGetBG Lib "GSWDLL32.DLL" () As Long
Declare Function GSGetCC Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetCos Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetCurveCoeff Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetCurX Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetCurY Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetE Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetLog10 Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetLogE Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetMax Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetMean Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetMF Lib "GSWDLL32.DLL" (ByVal nMode&) As Long
Declare Function GSGetMin Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetPi Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetPrnHt Lib "GSWDLL32.DLL" (ByVal nUnits&) As Double
Declare Function GSGetPrnWid Lib "GSWDLL32.DLL" (ByVal nUnits&) As Double
Declare Function GSGetRTextHt Lib "GSWDLL32.DLL" (ByVal nCset&, ByVal nMode&, ByVal szString$) As Double
Declare Function GSGetRTextWid Lib "GSWDLL32.DLL" (ByVal nCset&, ByVal nMode&, ByVal szString$) As Double
Declare Function GSGetSD Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetSFHt Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetSFWid Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetSin Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetSXExt Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetSYExt Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetTan Lib "GSWDLL32.DLL" (ByVal fVal#) As Double
Declare Function GSGetVer Lib "GSWDLL32.DLL" (ByVal nVer&) As Long
Declare Function GSGetVXExt Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetVYExt Lib "GSWDLL32.DLL" () As Double
Declare Function GSGetWXExt Lib "GSWDLL32.DLL" (ByVal nMode&, ByVal nUnits&) As Double
Declare Function GSGetWYExt Lib "GSWDLL32.DLL" (ByVal nMode&, ByVal nUnits&) As Double
Declare Function GSGrid Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fAxisLen#, ByVal fGridLen#, ByVal nDivs&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSHLC Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSHotGraph Lib "GSWDLL32.DLL" (ByVal nMode&) As Long
Declare Function GSLabelnPie Lib "GSWDLL32.DLL" (ByVal fxOff#, ByVal fRad#, ByVal fWid#, ByVal fHt#, ByVal nPrec&, ByVal nMode&, ByVal nCset&, ByVal nTMode&, ByVal nClr&) As Long
Declare Function GSLabelnX Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fWid#, ByVal fHt#, ByVal fBaseVal#, ByVal fSTEPVal#, ByVal nPrec&, ByVal nNLabs&, ByVal nCset&, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSLabelnY Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fWid#, ByVal fHt#, ByVal fBaseVal#, ByVal fSTEPVal#, ByVal nPrec&, ByVal nNLabs&, ByVal nCset&, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSLineAbs Lib "GSWDLL32.DLL" (ByVal fxA#, ByVal fya#, ByVal fxB#, ByVal fyB#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSLineFit Lib "GSWDLL32.DLL" (ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSLineRel Lib "GSWDLL32.DLL" (ByVal fxr#, ByVal fya#, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSLinLog Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fBase#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSLoadRFont Lib "GSWDLL32.DLL" (ByVal nFamily&, ByVal nAttrib&, ByVal nSize&, ByVal nPitch&) As Long
Declare Function GSLoadVFont Lib "GSWDLL32.DLL" (ByVal nFamily&, ByVal nAttrib&, ByVal nPitch&) As Long
Declare Function GSLogAxis Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fLen#, ByVal fTickLen#, ByVal nCycles&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSLogGrid Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fAxisLen#, ByVal fGridLen#, ByVal nCycles&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSLogLin Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fCycleHt#, ByVal fBaseVal#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSLogLog Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fCycleHt#, ByVal fBaseY#, ByVal fCycleWid#, ByVal fBaseX#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSMClrRgn Lib "GSWDLL32.DLL" (ByVal nRgn&) As Long
Declare Function GSMean Lib "GSWDLL32.DLL" (ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSMGetX Lib "GSWDLL32.DLL" () As Double
Declare Function GSMGetY Lib "GSWDLL32.DLL" () As Double
Declare Function GSMinMax Lib "GSWDLL32.DLL" (ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSMMotion Lib "GSWDLL32.DLL" () As Long
Declare Function GSMNotify Lib "GSWDLL32.DLL" (ByVal hWnd&, ByVal nMsg&, ByVal nMode&) As Long
Declare Function GSMovePos Lib "GSWDLL32.DLL" (ByVal fxr#, ByVal fya#, ByVal nMode&) As Long
Declare Function GSMPtrOff Lib "GSWDLL32.DLL" () As Long
Declare Function GSMPtrOn Lib "GSWDLL32.DLL" () As Long
Declare Function GSMPtrType Lib "GSWDLL32.DLL" (ByVal nType&) As Long
Declare Function GSMSetRgn Lib "GSWDLL32.DLL" (ByVal fxr#, ByVal fya#, ByVal nMode&, ByVal fAng#) As Long
Declare Function GSMStatus Lib "GSWDLL32.DLL" () As Long
Declare Function GSOffView Lib "GSWDLL32.DLL" (ByVal nWin&, ByVal nView&) As Long
Declare Function GSOnView Lib "GSWDLL32.DLL" (ByVal nWin&, ByVal nView&) As Long
Declare Function GSOpenChildWin Lib "GSWDLL32.DLL" (ByVal hWndParent&, ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal fyExt#, ByVal nStyle&, ByVal nMode&, ByVal szTitle$) As Long
Declare Function GSOpenPrn Lib "GSWDLL32.DLL" (ByVal szDevice$, ByVal szFile$, ByVal nMode&) As Long
Declare Function GSOpenServer Lib "GSWDLL32.DLL" (ByVal szKey$, ByVal szHost$) As Long
Declare Function GSOpenView Lib "GSWDLL32.DLL" (ByVal nWin&, ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal fyExt#) As Long
Declare Function GSOpenWin Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal fyExt#, ByVal nStyle&, ByVal nMode&, ByVal szTitle$) As Long
Declare Function GSPicRead Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal nFormat&, ByVal nMode&, ByVal szFile$) As Long
Declare Function GSPicWrite Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal nFormat&, ByVal nMode&, ByVal szFile$) As Long
Declare Function GSPie2D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fRad#, ByVal nMode&) As Long
Declare Function GSPie3D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fRad#, ByVal fDepth#, ByVal fAng#, ByVal nMode&) As Long
Declare Function GSPolar Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fAng#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSPolarAxes Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fRad#, ByVal nRadDivs&, ByVal nAngDivs&, ByVal nMode&, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSPolyFill Lib "GSWDLL32.DLL" (ByVal fxr#, ByVal fya#, ByVal nMode&, ByVal fAng#, ByVal nPatt&, ByVal nClr&) As Long
Declare Function GSPolyVec Lib "GSWDLL32.DLL" (ByVal fxr#, ByVal fya#, ByVal nMode&, ByVal fAng#, ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSPrnOut Lib "GSWDLL32.DLL" (ByVal nView&, ByVal nNCopies&, ByVal nMode&) As Long
Declare Function GSPrnSetup Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#, ByVal nUnits&, ByVal nMode&) As Long
Declare Function GSRText Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal nCset&, ByVal nTMode&, ByVal nClr&, ByVal szString$) As Long
Declare Function GSScatter Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSSD Lib "GSWDLL32.DLL" (ByVal nStyle&, ByVal nClr&) As Long
Declare Function GSSelectPalette Lib "GSWDLL32.DLL" (ByVal nMode&) As Long
Declare Function GSSetBG Lib "GSWDLL32.DLL" (ByVal nClr&) As Long
Declare Function GSSetPal Lib "GSWDLL32.DLL" (ByVal nClr&, ByVal nR&, ByVal nG&, ByVal nB&) As Long
Declare Function GSSetRFontFace Lib "GSWDLL32.DLL" (ByVal nFamily&, ByVal szFaceName$) As Long
Declare Function GSSetROP Lib "GSWDLL32.DLL" (ByVal nROP&) As Long
Declare Function GSSetVFontFace Lib "GSWDLL32.DLL" (ByVal nFamily&, ByVal szFaceName$) As Long
Declare Function GSShade Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal nPatt&, ByVal nClr&) As Long
Declare Function GSSizeSymbol Lib "GSWDLL32.DLL" (ByVal fDiam#) As Long
Declare Function GSStatsArr Lib "GSWDLL32.DLL" (ByVal nIndex&) As Long
Declare Function GSStatsWin Lib "GSWDLL32.DLL" (ByVal fxBL#, ByVal fyBL#, ByVal fxTR#, ByVal fyTR#) As Long
Declare Function GSSymbol Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal nSymbol&, ByVal nClr&) As Long
Declare Function GSTapeGraph Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fDepth#, ByVal fAng#, ByVal nMode&, ByVal nClr&) As Long
Declare Function GSTimeGraph Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal nPts&, ByVal nGroups&, ByVal nMode&) As Long
Declare Function GSTimeUpdate Lib "GSWDLL32.DLL" (ByVal nMode&, ByVal nGroup&, fData#) As Long
Declare Function GSUseView Lib "GSWDLL32.DLL" (ByVal nWin&, ByVal nView&) As Long
Declare Function GSVText Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal fAng#, ByVal nCset&, ByVal nMode&, ByVal nClr&, ByVal szString$) As Long
Declare Function GSWinHandle Lib "GSWDLL32.DLL" (ByVal nWindow&) As Long
Declare Function GSWinNotify Lib "GSWDLL32.DLL" (ByVal hWnd&, ByVal nWM&, ByVal nEvents) As Long
Declare Function GSWinPaint Lib "GSWDLL32.DLL" (ByVal nMode&) As Long
Declare Function GSXDataScale Lib "GSWDLL32.DLL" (ByVal fScale#) As Long
Declare Function GSXYGraph Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal nMode&, ByVal nClr&) As Long
Declare Function VBAGDataLabels Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nLabs&, ByVal szLabs$) As Long
Declare Function VBAGLabels Lib "GSWAG32.DLL" (ByVal nNLabs&, ByVal szLabs$) As Long
Declare Function VBAGLabelY Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nLabs&, ByVal szLabs$) As Long
Declare Function VBAGLabelZ Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nLabs&, ByVal szLabs$) As Long
Declare Function VBAGLegend Lib "GSWAG32.DLL" (ByVal nNLegs&, ByVal szLegs$) As Long
Declare Function VBGSDataLabels Lib "GSWAG32.DLL" (ByVal nMode&, ByVal nPrec&, ByVal nCset&, ByVal nTMode&, ByVal nClr&, ByVal fOff#, ByVal nLabs&, ByVal szLabs$) As Long
Declare Function VBGSLabelPie Lib "GSWAG32.DLL" (ByVal fxOff#, ByVal fRad#, ByVal fWid#, ByVal fHt#, ByVal nNLabs&, ByVal nMode&, ByVal nCset&, ByVal nTMode&, ByVal nClr&, ByVal szLabs$) As Long
Declare Function VBGSLabelX Lib "GSWAG32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fWid#, ByVal fHt#, ByVal nNLabs&, ByVal nCset&, ByVal nMode&, ByVal nClr&, ByVal szLabs$) As Long
Declare Function VBGSLabelY Lib "GSWAG32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fInc#, ByVal fWid#, ByVal fHt#, ByVal nNLabs&, ByVal nCset&, ByVal nMode&, ByVal nClr&, ByVal szLabs$) As Long
Declare Function VBGSLegend Lib "GSWAG32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal nNLeg&, ByVal nRows&, ByVal nMode&, ByVal nCset&, ByVal nTMode&, ByVal nTClr&, nBClr&, nBPatt&, ByVal szLegs$) As Long

Global Const BLACK = 0
Global Const BLUE = 1
Global Const GREEN = 2
Global Const CYAN = 3
Global Const RED = 4
Global Const MAGENTA = 5
Global Const BROWN = 6
Global Const WHITE = 7
Global Const GREY = 8
Global Const YELLOW = 14
Global Const LIGHT = 8
Global Const LSSOLID = 0
Global Const LSDASH = 1
Global Const LSDOT = 2
Global Const LSDASHDOT = 3
Global Const LSDASHDD = 4
Global Const LSNULL = 5
Global Const BRSOLID = 0
Global Const BRNULL = 1
Global Const BRHATCH = 2
Global Const BRHATCHMAX = 6
Global Const BRBITMAP = 16
Global Const BRBITMAPMAX = 16
Global Const BRTRANS = 64
Global Const CSSYSTEM = 0
Global Const CSUSER = 1
Global Const CSRASTER = 2
Global Const TXEXACT = 1
Global Const TXLEFT = 0
Global Const TXMID = 2
Global Const TXRIGHT = 4
Global Const TXBOTTOM = 0
Global Const TXBASELINE = 8
Global Const TXTOP = 16
Global Const TXUP90 = 32
Global Const TXDOWN90 = 64
Global Const TXTRANS = 256
Global Const TXFLIP = 512
Global Const TXSLIP = 1024
Global Const UNLOG = 0
Global Const UNMM = 1
Global Const UNINCH = 2
Global Const UNDEV = 3
Global Const NULLDATA = 256
Global Const GSSURFACENET = 1
Global Const GSSURFACEGRID = 2
Global Const GSSURFACESIDES = 4
Global Const GSPERSPECTIVE = 1
Global Const GSCAGESIDEFLIP = 2
Global Const GSCAGEBACKFLIP = 4
Global Const PALDEFAULT = 0
Global Const PALGREYSCALE = 1
Global Const PALPASTEL = 2
Global Const PALRGBCMY = 3
Global Const PALRAINBOW = 4
Global Const PALREDSCALE = 5
Global Const PALGREENSCALE = 6
Global Const PALBLUESCALE = 7
Global Const PALCYANSCALE = 8
Global Const PALMAGENTASCALE = 9
Global Const PALYELLOWSCALE = 10
Global Const PALUSER = 11
Global Const EBY = 0
Global Const EBX = 1
Global Const EBNOPLUS = 1
Global Const EBNOMINUS = 2
Global Const EBNOSTEM = 4
Global Const EBNOTICK = 8
Global Const EBFIXED = 0
Global Const EBPERCENT = 1
Global Const EBSTDDEV = 2
Global Const EBSTDERR = 3
Global Const EBMAXMIN = 4
Global Const ERRPLUS = 0
Global Const ERRMINUS = 1
Global Const BWPARAMETRIC = 1
Global Const BWVARX = 4
Global Const BWNONOTCH = 8
Global Const BWNOWHISKER = 16
Global Const BWNOSAMPLES = 32
Global Const TIGMEAN = 1
Global Const TIGSTDDEV = 2
Global Const TIGUPDATE = 0
Global Const TIGLOAD = 1
Global Const TIGHIDE = 2
Global Const TIGSHOW = 3
Global Const AGFFTREALIMAG = 0
Global Const AGFFTAMPPHASE = 1
Global Const AGFFTSQAMPPHASE = 2
Global Const AGFFTINTERLEAVE = 16
Global Const AGGRFRAME = 1
Global Const AGGRFILL = 2
Global Const AGGRDROPSHADOW = 4
Global Const AGGRRAISED = 8
Global Const AGGRLOWERED = 16
Global Const AGTTLG = 0
Global Const AGTTLX = 1
Global Const AGTTLYLEFT = 2
Global Const AGTTLYRIGHT = 3
Global Const AGTTLUP = 4
Global Const AGTTLDOWN = 8
Global Const AGTTLFRAME = 16
Global Const AGTTLFILL = 32
Global Const AGTTLDROPSHADOW = 64
Global Const AGTTLRAISED = 128
Global Const AGTTLLOWERED = 256
Global Const AGGRIDX = 0
Global Const AGGRIDY = 1
Global Const AGGRIDZ = 2
Global Const AGCAGETHIN = 1
Global Const AGLABYLEFT = 0
Global Const AGLABYRIGHT = 1
Global Const AGEBY = 0
Global Const AGEBX = 1
Global Const AGEBNOPLUS = 1
Global Const AGEBNOMINUS = 2
Global Const AGEBNOSTEM = 4
Global Const AGEBNOTICK = 8
Global Const AGEBFIXED = 0
Global Const AGEBPERCENT = 1
Global Const AGEBSTDDEV = 2
Global Const AGEBSTDERR = 3
Global Const AGEBMAXMIN = 4
Global Const AGTIGNOLABELS = 1
Global Const AGTIGLEGCLR = 2
Global Const AGTIGXGRID = 4
Global Const AGTIGYGRID = 8
Global Const AGTIGMEAN = 1
Global Const AGTIGSTDDEV = 2
Global Const AGTIGUPDATE = 0
Global Const AGTIGLOAD = 1
Global Const AGTIGHIDEDATA = 2
Global Const AGTIGSHOWDATA = 3
Global Const AGFGALL = 0
Global Const AGFGTITLEG = 1
Global Const AGFGTITLEYL = 2
Global Const AGFGTITLEYR = 3
Global Const AGFGTITLEX = 4
Global Const AGFGLABELS = 5
Global Const AGFGLEGEND = 6
Global Const AGFGAXIS = 7
Global Const AGFGGRID = 8
Global Const AGFGMEAN = 9
Global Const AGFGMINMAX = 10
Global Const AGFGSTDDEV = 11
Global Const AGFGBESTFIT = 12
Global Const AGFGCURVE = 13
Global Const AGFGDATALABELS = 14
Global Const AGDLTEXT = 0
Global Const AGDLDATA = 1
Global Const AGDLGROUPCLR = 4
Global Const BARVARXPOS = 16
Global Const BARVARZPOS = 64
Global Const G3GRIDX = 0
Global Const G3GRIDY = 1
Global Const G3GRIDZ = 2
Global Const AGPIE2D = 1
Global Const AGPIE3D = 2
Global Const AGBAR2D = 3
Global Const AGBAR3D = 4
Global Const AGGANTT = 5
Global Const AGLINE = 6
Global Const AGLOGLIN = 7
Global Const AGAREA = 8
Global Const AGSCATTER = 9
Global Const AGPOLAR = 10
Global Const AGHLC = 11
Global Const AGBUBBLE = 12
Global Const AGTAPE = 13
Global Const AGAREA3D = 14
Global Const AGLOGLOG = 15
Global Const AGLINLOG = 16
Global Const AGBOXWHISKER = 17
Global Const AGSURFACE = 128
Global Const AGTRUE3D = 128
Global Const AGCOMBO = 256
Global Const AGDUALYAXIS = 512
Global Const AGLINENOLABELS = &H1
Global Const AGLINELEGCLR = &H2
Global Const AGLINEXGRID = &H4
Global Const AGLINEYGRID = &H8
Global Const AGLINESYMBOLS = &H10
Global Const AGLINEMONO = &H20
Global Const AGLINEFORECLR = &H40
Global Const AGLINESOLID = &H4000
Global Const AGLINETHICK = &H2000
Global Const AGLINEPATT = &H1000
Global Const AGLINESTICK = &H800
Global Const AGSCATTNOLABELS = &H1
Global Const AGSCATTLEGCLR = &H2
Global Const AGSCATTXGRID = &H4
Global Const AGSCATTYGRID = &H8
Global Const AGSCATTSYMBOLS = &H10
Global Const AGSCATTMONO = &H20
Global Const AGSCATTFORECLR = &H40
Global Const AGSCATTTHICK = &H2000
Global Const AGSCATTPATT = &H1000
Global Const AGSCATTCURVE = &H400
Global Const AGSCATTSTICK = &H200
Global Const AGLOGLINNOLABS = &H1
Global Const AGLOGLINLEGCLR = &H2
Global Const AGLOGLINXGRID = &H4
Global Const AGLOGLINYGRID = &H8
Global Const AGLOGLINSYMBOLS = &H10
Global Const AGLOGLINMONO = &H20
Global Const AGLOGLINFORECLR = &H40
Global Const AGLOGLINSOLID = &H4000
Global Const AGLOGLINTHICK = &H2000
Global Const AGLOGLINPATT = &H1000
Global Const AGLOGLINSTICK = &H800
Global Const AGLINLOGNOLABS = &H1
Global Const AGLINLOGLEGCLR = &H2
Global Const AGLINLOGXGRID = &H4
Global Const AGLINLOGYGRID = &H8
Global Const AGLINLOGSYMBOLS = &H10
Global Const AGLINLOGMONO = &H20
Global Const AGLINLOGFORECLR = &H40
Global Const AGLINLOGSOLID = &H4000
Global Const AGLINLOGTHICK = &H2000
Global Const AGLINLOGPATT = &H1000
Global Const AGLINLOGSTICK = &H800
Global Const AGLOGLOGNOLABS = &H1
Global Const AGLOGLOGLEGCLR = &H2
Global Const AGLOGLOGXGRID = &H4
Global Const AGLOGLOGYGRID = &H8
Global Const AGLOGLOGSYMBOLS = &H10
Global Const AGLOGLOGMONO = &H20
Global Const AGLOGLOGFORECLR = &H40
Global Const AGLOGLOGSOLID = &H4000
Global Const AGLOGLOGTHICK = &H2000
Global Const AGLOGLOGPATT = &H1000
Global Const AGLOGLOGSTICK = &H800
Global Const AGHLCNOLABELS = &H1
Global Const AGHLCXGRID = &H4
Global Const AGHLCYGRID = &H8
Global Const AGHLCMONO = &H20
Global Const AGHLCFORECLR = &H40
Global Const AGHLCOPEN = &H80
Global Const AGHLCCANDLESTICK = &H100
Global Const AGHLCNOCLOSE = &H4000
Global Const AGHLCNOBARS = &H2000
Global Const AGHLCTHICK = &H1000
Global Const AGBARNOLABELS = &H1
Global Const AGBARLEGCLR = &H2
Global Const AGBARXGRID = &H4
Global Const AGBARYGRID = &H8
Global Const AGBARMONO = &H20
Global Const AGBARFORECLR = &H40
Global Const AGBARCLUSTER = &H4000
Global Const AGBARSTACK = &H2000
Global Const AGBARSTACKPC = &H1000
Global Const AGBARHORIZ = &H800
Global Const AGBARCLUSTZ = &H400
Global Const AGBARLASTFIRST = &H200
Global Const AGGANTTNOLABELS = &H1
Global Const AGGANTTXGRID = &H4
Global Const AGGANTTYGRID = &H8
Global Const AGGANTTMONO = &H20
Global Const AGGANTTFORECLR = &H40
Global Const AGGANTTSPACE = &H4000
Global Const AGGANTTLASTFIRST = &H2000
Global Const AGAREANOLABELS = &H1
Global Const AGAREALEGCLR = &H2
Global Const AGAREAXGRID = &H4
Global Const AGAREAYGRID = &H8
Global Const AGAREAMONO = &H20
Global Const AGAREAFORECLR = &H40
Global Const AGAREAABS = &H4000
Global Const AGAREAPC = &H2000
Global Const AGPIENOLABELS = &H1
Global Const AGPIELEGCLR = &H2
Global Const AGPIEMONO = &H20
Global Const AGPIEFORECLR = &H40
Global Const AGPIENOLINES = &H4000
Global Const AGPIESEGCLR = &H2000
Global Const AGPIEPERCENT = &H1000
Global Const AGPIEPCCHAR = &H800
Global Const AGPIESAMECLR = &H400
Global Const AGPOLARNOLABELS = &H1
Global Const AGPOLARLEGCLR = &H2
Global Const AGPOLARANGGRID = &H4
Global Const AGPOLARRADGRID = &H8
Global Const AGPOLARSYMBOL = &H10
Global Const AGPOLARMONO = &H20
Global Const AGPOLARFORECLR = &H40
Global Const AGPOLARLINE = &H4000
Global Const AGPOLARTHICK = &H2000
Global Const AGPOLARPATT = &H1000
Global Const AGPOLARSTICK = &H800
Global Const AGBUBNOLABELS = &H1
Global Const AGBUBLEGCLR = &H2
Global Const AGBUBXGRID = &H4
Global Const AGBUBYGRID = &H8
Global Const AGBUBMONO = &H20
Global Const AGBUBFORECLR = &H40
Global Const AGTAPENOLABELS = &H1
Global Const AGTAPELEGCLR = &H2
Global Const AGTAPEXGRID = &H4
Global Const AGTAPEYGRID = &H8
Global Const AGTAPEMONO = &H20
Global Const AGTAPEFORECLR = &H40
Global Const AGSRFCNOLABELS = &H1
Global Const AGSRFCXGRID = &H4
Global Const AGSRFCYGRID = &H8
Global Const AGSRFCMONO = &H20
Global Const AGSRFCNET = &H4000
Global Const AGSRFCBLACKLINES = &H2000
Global Const AGSRFCSIDEWALL = &H1000
Global Const AGBWPARAMETRIC = &H4000
Global Const AGBWNONOTCH = &H2000
Global Const AGBWNOWHISKER = &H1000
Global Const AGBWNOSAMPLES = &H800
Global Const AGBWNOLABELS = &H1
Global Const AGBWXGRID = &H4
Global Const AGBWYGRID = &H8
Global Const AGBWMONO = &H20
Global Const AGMEAN = &H1
Global Const AGMINMAX = &H2
Global Const AGSD = &H4
Global Const AGLINEFIT = &H8
Global Const AGCURVEFIT = &H10
Global Const AGFUSETITG = 0
Global Const AGFUSETITXY = 1
Global Const AGFUSELABS = 2
Global Const AGFUSELEG = 3
Global Const AGLEGBOTTOM = 1
Global Const AGLEGTOP = 2
Global Const AGLEGCENTRE = 0
Global Const AGLEGLEFT = 1
Global Const AGLEGRIGHT = 2
Global Const AGLEGFRAME = 1
Global Const AGLEGFILL = 2
Global Const AGLEGDROPSHADOW = 4
Global Const AGLEGRAISED = 8
Global Const AGLEGLOWERED = 16
Global Const AGVARORIGIN = &H1
Global Const AGNOLABELS = &H2
Global Const AGUTICKS = &H4
Global Const AGUMAX = &H8
Global Const AGUMIN = &H10
Global Const AGLABEVERY = &H20
Global Const AGNOTICKS = &H40
Global Const AGTICKEVERY = &H80
Global Const AGAXISLEFT = &H100
Global Const AGAXISTOP = &H100
Global Const AGAXISRIGHT = &H200
Global Const AGAXISBOTTOM = &H200
Global Const AGAXISMINORTICK = &H400
Global Const AGAXISVERTLABELS = &H800
Global Const AG3DISO = &H0
Global Const AG3DPERSPECTIVE = &H1
Global Const AG3DFLIPLR = &H2
Global Const AG3DFLIPFB = &H4
Global Const AG3DSETDEPTH = &H1
Global Const AG3DSETXGAP = &H2
Global Const AG3DSETZGAP = &H4
Global Const ARABS = 1
Global Const ARVARX = 2
Global Const ARPC = 4
Global Const BARSIMPLE = 0
Global Const BARSTACK = 1
Global Const BARCLUST = 2
Global Const BARSTACKPC = 3
Global Const BARCLUSTZ = 4
Global Const BARHORIZ = 8
Global Const BARVARPOS = 16
Global Const BARHALFTONE = 32
Global Const BARLASTFIRST = 64
Global Const GAVARY = 1
Global Const GASPACE = 2
Global Const GALASTFIRST = 4
Global Const HLCVARX = 1
Global Const HLCNOCLOSE = 2
Global Const HLCNOBARS = 4
Global Const HLCTHICK = 8
Global Const HLCOPEN = 16
Global Const HLCCANDLESTICK = 32
Global Const PCNOEXPL = 0
Global Const PCEXPL = 1
Global Const PCSAMECLR = 2
Global Const PCABSENT = 4
Global Const POGLINE = 1
Global Const POGSYMBOL = 2
Global Const POGSTICK = 4
Global Const POGVARANG = 8
Global Const POGTHICK = 16
Global Const POGPATT = 32
Global Const POGGROUPED = 64
Global Const XYGSIMPLE = 0
Global Const XYGLINE = 1
Global Const XYGSYMBOL = 2
Global Const XYGSTICK = 4
Global Const XYGVARX = 8
Global Const XYGTHICK = 16
Global Const XYGPATT = 32
Global Const XYGGROUPED = 64
Global Const XYGVARZ = 128
Global Const OWSTHICKFRAME = 1
Global Const OWSBORDER = 2
Global Const OWSHSCROLL = 4
Global Const OWSVSCROLL = 8
Global Const OWSMAXIMIZEBOX = 16
Global Const OWSMINIMIZEBOX = 32
Global Const OWSSETFOCUS = 64
Global Const OWSHIDDEN = 128
Global Const OWSSYSMENU = 256
Global Const OWSCAPTION = 512
Global Const OWSCLIPCHILDREN = 1024
Global Const OWSCLIPSIBLINGS = 2048
Global Const OWSTRANSPARENT = 4096
Global Const OWMFIXED = 0
Global Const OWMFITHORZ = 1
Global Const OWMFITVERT = 2
Global Const OWMFITBOTH = 3
Global Const OWMFITOPT = 4
Global Const OWMFLEXIBLE = 5
Global Const OWMCENTRED = 256
Global Const CVKEEP = 0
Global Const CVDISCARD = 1
Global Const ROREPLACE = 0
Global Const ROOR = 1
Global Const ROXOR = 2
Global Const RONOT = 3
Global Const GWWHOLE = 0
Global Const GWCLIPPED = 1
Global Const GVSERVER = 0
Global Const GVDLL = 1
Global Const WNPAINT = 1
Global Const WNSIZE = 2
Global Const WPAUTO = 0
Global Const WPMANUAL = 1
Global Const WPNONE = 2
Global Const WPPAINT = 3
Global Const WPBITMAP = 4
Global Const WPMETAFILE = 5
Global Const AXTICKOUT = 0
Global Const AXTICKIN = 1
Global Const AXTICKTHRU = 2
Global Const AXISX = 0
Global Const AXISY = 4
Global Const GRX = 0
Global Const GRY = 1
Global Const GRNOFIRST = 2
Global Const GRNOLAST = 4
Global Const PARADGRID = 1
Global Const PAANGGRID = 2
Global Const PATHICK = 4
Global Const CGGRIDX = 1
Global Const CGGRIDY = 2
Global Const CGGRIDZ = 4
Global Const CFPOLY = 0
Global Const CFLOG = 1
Global Const CFEXP1 = 2
Global Const CFEXP2 = 3
Global Const CFPOWER = 4
Global Const CFINV1 = 5
Global Const CFINV2 = 6
Global Const CFINV3 = 7
Global Const CFINV4 = 8
Global Const CFINV5 = 9
Global Const CFSPLINE = 10
Global Const CFMOVINGAVEMID = 11
Global Const CFMOVINGAVEEND = 12
Global Const PXBMP = 1
Global Const PXPMF = 2
Global Const PXWMF = 3
Global Const PXDIB = 4
Global Const PXPCX = 5
Global Const PXMONO = 256
Global Const PXCENTRE = 1
Global Const PXSTRETCH = 2
Global Const PXTILE = 3
Global Const CBBMP = 1
Global Const CBWMF = 2
Global Const CBDIB = 4
Global Const CBMONO = 256
Global Const CBCENTRE = 1
Global Const CBSTRETCH = 2
Global Const CBTILE = 3
Global Const LAUPDATE = 1
Global Const LATHICK = 4
Global Const LACONT = 16
Global Const MPPOLAR = 1
Global Const LRNOUPDATE = 1
Global Const LRPOLAR = 2
Global Const LRTHICK = 4
Global Const AARADIUS = 1
Global Const AAFILL = 2
Global Const AATHICK = 4
Global Const AWTHIN = 1
Global Const AWMEDIUM = 0
Global Const AWTHICK = 2
Global Const AWFILLED = 0
Global Const AWHOLLOW = 4
Global Const BXNOBOX = 128
Global Const BXSHADOW = 256
Global Const BXBORDER = 512
Global Const BXRAISED = 1024
Global Const BXLOWERED = 2048
Global Const CCFILL = 2
Global Const CCTHICK = 4
Global Const ELRADIUS = 1
Global Const ELFILL = 2
Global Const ELTHICK = 4
Global Const PFXYORG = 0
Global Const PFRAORG = 1
Global Const PFXYDATA = 0
Global Const PFRADATA = 2
Global Const PFMIRRORV = 4
Global Const PFMIRRORH = 8
Global Const PVXYORG = 0
Global Const PVRAORG = 1
Global Const PVXYDATA = 0
Global Const PVRADATA = 2
Global Const PVMIRRORV = 4
Global Const PVMIRRORH = 8
Global Const PVTHICK = 16
Global Const CLTRANSP = 0
Global Const CLOPAQUE = 1
Global Const PRNSETMODE = 1
Global Const PRNPORTRAIT = 2
Global Const PRNLANDSCAPE = 4
Global Const PRNSETMODECANCEL = 8
Global Const PRNMM = 1
Global Const PRNINCH = 2
Global Const PRNFIXED = 0
Global Const PRNFITHORZ = 1
Global Const PRNFITVERT = 2
Global Const PRNFITBOTH = 3
Global Const PRNFITOPT = 4
Global Const PRNCENTRED = 256
Global Const PRNWINDOW = 0
Global Const PRNVIEW = 1
Global Const PRNFF = 2
Global Const PRNFRAME = 4
Global Const PRNCANCEL = 8
Global Const FOROMAN = 1
Global Const FOSWISS = 2
Global Const FOMODERN = 3
Global Const FOSCRIPT = 4
Global Const FODECO = 5
Global Const FOITALIC = 16
Global Const FOBOLD = 32
Global Const FOULINE = 64
Global Const FOFIXED = 128
Global Const LPSEGCLR = 1
Global Const LPNOLINES = 2
Global Const LPPC = 4
Global Const LPPCCHAR = 8
Global Const LGBOX = 1
Global Const LGTXCLR = 2
Global Const LGBG = 4
Global Const LGLINE = 8
Global Const LGSYMBOL = 16
Global Const DLTEXT = 0
Global Const DLDATA = 0
Global Const DLGROUPCLR = 4
Global Const SYHOLLOWFILL = 256
Global Const MNTRACK = 1
Global Const MNPRESS = 2
Global Const MNHIT = 4
Global Const MNHITPT = 8
Global Const MNHITOVERLAY = 16
Global Const MCARROW = 0
Global Const MCIBEAM = 1
Global Const MCWAIT = 2
Global Const MCCROSS = 3
Global Const MCUPARROW = 4
Global Const MCSIZE = 5
Global Const MCICON = 6
Global Const MCSIZENWSE = 7
Global Const MCSIZENESW = 8
Global Const MCSIZEWE = 9
Global Const MCSIZENS = 10
Global Const MBLEFT = 1
Global Const MBMIDDLE = 2
Global Const MBRIGHT = 4
Global NumOfPlots As Integer






Sub sizeform(frmIn As Form)
frmIn.Height = 7230
frmIn.Width = 9615
Dim iTop  As Integer, iLeft As Integer
If frmIn.WindowState <> 0 Then Exit Sub
iTop = (Screen.Height - frmIn.Height) \ 2
iLeft = (Screen.Width - frmIn.Width) \ 2
frmIn.Move iLeft, iTop
End Sub
