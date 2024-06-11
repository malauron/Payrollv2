Attribute VB_Name = "modGenDtr"
Option Explicit

Dim rsDtrTmp            As ADODB.Recordset
Dim rsTitoTmp           As ADODB.Recordset

Public Sub Compute_Dtr(mPerCode As String, mPayFreqCode As String, mEmpNo As String, mStartDate As String, mEndDate As String)

    Dim mTin            As Date
    Dim mTout           As Date
    Dim mLasTout        As String
    Dim mAdvDate        As String
    
    Dim mT1in           As String
    Dim mT1out          As String
    Dim mT2in           As String
    Dim mT2out          As String
    Dim mST1in          As String
    Dim mST1out         As String
    Dim mST2in          As String
    Dim mST2out         As String
    Dim mWrkdate        As String

    Dim rsEmployee      As ADODB.Recordset
    Dim rsOT            As ADODB.Recordset
    Dim rsOTTmp         As ADODB.Recordset
    Dim rsDtrEmp        As ADODB.Recordset
    Dim rsHoliday       As ADODB.Recordset
    Dim rsParmtr        As ADODB.Recordset
    Dim rsDts           As ADODB.Recordset
    Dim rsOTChk         As ADODB.Recordset
    Dim rsOTCmpr        As ADODB.Recordset
    Dim rsCostCenter    As ADODB.Recordset
    
    Dim mOT_Start       As String
    Dim mOT_End         As String
    Dim mNiteStart      As String
    Dim mNiteEnd        As String
    Dim mTimeIN         As String
    Dim mTimeOUT        As String
    Dim mDTS_Start      As String
    Dim mDTS_End        As String
    
    Dim mDTSActIN       As Date
    Dim mDTSActOUT      As Date
    Dim mOTActIN        As Date
    Dim mOTActOUT       As Date
    Dim mNiteIN         As Date
    Dim mNiteOUT        As Date
    Dim mPrevOUT        As Date
    Dim mTimeFrom       As Date
    Dim mTimeTo         As Date
    
    Dim mLateAllow      As Double
    Dim mOTAllow        As Double
    Dim mWrkHrs         As Double
    Dim mNiteWrkHrs     As Double
    Dim mUTMin          As Double
    Dim mLackHrs        As Double
    Dim mTtlReghrs      As Double
    
'*******************************************************
    
    
    
    NetOpen rsEmployee, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee where payfreqcode = '" & mPayFreqCode & "' and employeecode = '" & mEmpNo & "' "
    
    If rsEmployee.RecordCount > 0 Then
'
'        fra1.Enabled = False
'        Me.MousePointer = vbHourglass
'        cmdGenerate.Enabled = False
        
'        pb1.Max = rsEmployee.RecordCount
'        pb1.Value = 0
        
        rsEmployee.MoveFirst
        
        NetOpen rsParmtr, "select * from parmtr"
        If rsParmtr.RecordCount > 0 Then
            mLateAllow = Format(rsParmtr!lateallowance / 60, "#,##0.00")
            mUTMin = Format(rsParmtr!utmin / 60, "#,##0.00")
        End If
        
'        ConMain.Execute "set autocommit = 0"
'        ConMain.BeginTrans
        
        Do While Not rsEmployee.EOF
            
            'pb1.Value = pb1.Value + 1
                                                    
            ConMain.Execute "delete from dtremp where employeecode = '" & mEmpNo & "' and " & _
                                    "workdate between '" & mStartDate & "'  and  " & _
                                    "'" & mEndDate & "' and updatable = 'Y'"
                                                    
            Create_TmpTito mEmpNo, mStartDate, mEndDate, mPerCode
            
            Load_Dtr mEmpNo, mStartDate, mEndDate
            
            mLasTout = ""
            
            If rsDtrTmp.RecordCount > 0 Then
                'Assigning TITO to Employee DTR
                
                With rsDtrTmp
'                    pb2.Max = .RecordCount
'                    pb2.Value = 0
                    .MoveFirst
                    Do While Not .EOF
                        'pb2.Value = pb2.Value + 1
                        
                        !dayswrk = 0
                        !wrkhrs = 0
                        !nitewrkhrs = 0
                        !absent = 0
                        !latehrs = 0
                        !uthrs = 0
                        
                        If rsTitoTmp.RecordCount > 0 Then
                            If !updatable <> 0 Then
                                'For shifts that only have two time slots
                                If Trim(!st1in) <> "" And Trim(!st1out) <> "" And Trim(!st2in) = "" And Trim(!st2out) = "" Then
                                    mTin = Format(CDate(!wrkdate) & " " & !st1in, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    If CDate(!st1in) > CDate(!st1out) Then
                                        mTout = Format(CDate(!wrkdate) + 1 & " " & !st1out, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    Else
                                        mTout = Format(CDate(!wrkdate) & " " & !st1out, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    End If
                                    mST1in = mTin
                                    mST1out = mTout
                                    If mLasTout = "" Then
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "'"
                                    Else
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    End If
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t1in) = "" Then !t1in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t1out) = "" Then !t1out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                End If
                                
                                'For shifts that have four time slots
                                If Trim(!st1in) <> "" And Trim(!st1out) <> "" And Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                    mTin = Format(CDate(!wrkdate) & " " & !st1in, "MM/DD/YYYY HH:NN:SS AM/PM")
                                    If CDate(!st1in) > CDate(!st1out) Then
                                        mAdvDate = CDate(!wrkdate) + 1
                                    Else
                                        mAdvDate = CDate(!wrkdate)
                                    End If
                                    mTout = Format(CDate(mAdvDate) & " " & !st1out, "MM/DD/YYYY HH:NN:S AM/PM")
                                    mST1in = mTin
                                    mST1out = mTout
                                    If mLasTout = "" Then
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "'"
                                    Else
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    End If
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t1in) = "" Then !t1in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t1out) = "" Then !t1out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                    If CDate(!st1out) > CDate(!st2in) Then
                                        mAdvDate = CDate(mAdvDate) + 1
                                    Else
                                        mAdvDate = CDate(mAdvDate)
                                    End If
                                    mTin = Format(CDate(mAdvDate) & " " & !st2in, "MM/DD/YYYY HH:NN:SS AM/PM")
                                    
                                    If CDate(!st2in) > CDate(!st2out) Then
                                        mAdvDate = CDate(mAdvDate) + 1
                                    Else
                                        mAdvDate = CDate(mAdvDate)
                                    End If
                                    
                                    mTout = Format(CDate(mAdvDate) & " " & !st2out, "MM/DD/YYYY HH:NN:S AM/PM")
                                    mST2in = mTin
                                    mST2out = mTout
                                    rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t2in) = "" Then !t2in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t2out) = "" Then !t2out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                    
                                End If
                            End If
                        End If
                    
                        'Do the computations
                        'Clear all Time and Date variables
                        
                        mT1in = ""
                        mT1out = ""
                        mT2in = ""
                        mT2out = ""
                        mST1in = ""
                        mST1out = ""
                        mST2in = ""
                        mST2out = ""
                        mWrkHrs = 0
                        mNiteWrkHrs = 0
                        
                        'Check if employee has a schedule for today
                        If Trim(!st1in) <> "" And Trim(!st1out) <> "" Then
                            'Set Workdate to be used for shifting schedule
                            mWrkdate = !wrkdate
                            'assgin shifting schedule variables with vaues
                            mST1in = mWrkdate & " " & !st1in
                            If CDate(mST1in) > CDate(mWrkdate & " " & !st1out) Then
                                mWrkdate = Format(CDate(!wrkdate) + 1, "MM/DD/YYYY")
                            End If
                            
                            mST1out = mWrkdate & " " & !st1out
                            If Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                If CDate(mST1out) > CDate(mWrkdate & " " & !st2in) Then
                                    mWrkdate = CDate(mWrkdate) + 1
                                End If
                                mST2in = mWrkdate & "  " & !st2in
                                If CDate(mST2in) > CDate(mWrkdate & " " & !st2out) Then
                                    mWrkdate = CDate(mWrkdate) + 1
                                End If
                                mST2out = mWrkdate & " " & !st2out
                            End If
                            'Set Workdate back to original date to be used for actual time logs
                            mWrkdate = !wrkdate
                            'set values for 1st Tito
                            If Trim(!t1in) <> "" And Trim(!t1out) <> "" Then
                                If CDate(mWrkdate & " " & !t1in) > CDate(mST1out) Then
                                    mT1in = CDate(mWrkdate) - 1 & " " & !t1in
                                Else
                                    mT1in = mWrkdate & " " & !t1in
                                End If
                                If CDate(mT1in) > CDate(mWrkdate & " " & !t1out) Then
                                    mWrkdate = CDate(mWrkdate) + 1
                                End If
                                mT1out = mWrkdate & " " & !t1out
                                If CDate(mT1out) < CDate(mST1in) And CDate(mT1out) < CDate(mST1out) Then
                                        mWrkdate = CDate(mWrkdate) + 1
                                        mT1out = mWrkdate & " " & !t1out
                                End If
                            End If
                            
                            'set values for 2nd tito
                            If Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                If Trim(!t2in) <> "" And Trim(!t2out) <> "" Then
                                    If Trim(!t1in) <> "" Then
                                        If CDate(mT1out) > CDate(mWrkdate & " " & !t2in) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2in = mWrkdate & " " & !t2in
                                        If CDate(mT2in) > CDate(mWrkdate & " " & !t2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2out = mWrkdate & " " & !t2out
                                    Else
                                        If CDate(mWrkdate & " " & !t2in) > CDate(mST2out) Then
                                            mT2in = CDate(mWrkdate) - 1 & " " & !t2in
                                        Else
                                            mT2in = mWrkdate & " " & !t2in
                                        End If
                                        If CDate(mT2in) > CDate(mWrkdate & " " & !t2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2out = mWrkdate & " " & !t2out
                                    End If
                                    If CDate(mT2out) < CDate(mST2in) And CDate(mT2out) < CDate(mST2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                            mT2out = mWrkdate & " " & !t2out
                                    End If
                                End If
                            End If
                            
                            'Compute number of hours work, lates and undertimes.
                            If mST2in <> "" Then 'For schedules with four(4) time slots
                                If !travel = 0 And !leave = 0 Then
                                    If mT1in <> "" And mT2in <> "" Then 'if all four(4) time slots were consumed.
                                        !dayswrk = 1
                                        'compute for late
                                        If CDate(mT1in) > CDate(mST1in) Then
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                'If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) '- mLateAllow
                                                'End If
                                            End If
                                        End If
                                        'check if late on 2nd time in
                                        If !brkhrsperday > 0 Then
                                            If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                                'If (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) > mLateAllow Then
                                                    If .Fields("latehrs") > 0 Then
                                                        .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) '- mLateAllow
                                                    Else
                                                        .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday '- mLateAllow
                                                    End If
                                                'End If
                                            End If
                                        End If
                                        'compute for undertime
                                        If CDate(mT2out) < CDate(mST2out) Then
                                            If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                            End If
                                        End If
                                        mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                        mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                        mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    ElseIf mT1in <> "" And mT2in = "" Then 'if only the first two (2) time slots were consumed.
                                    
                                        If CDate(mT1in) <= CDate(mST1in) Then
                                            mTin = mST1in
                                        Else
                                            mTin = mT1in
                                            'compute for late
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                'If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) '- mLateAllow
                                                'End If
                                            End If
                                        End If
                                        
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            mTout = mT1out
                                            mWrkHrs = DiffHrs(mTin, mTout)
                                            'compute for 1st undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                            If Trim(.Fields("holiday")) = "" Then
                                                .Fields("absent") = 0.5
                                            End If
                                            .Fields("dayswrk") = 0.5
                                        ElseIf CDate(mT1out) >= CDate(mST1out) And CDate(mT1out) < CDate(mST2in) Then
                                            mTout = mST1out
                                            mWrkHrs = DiffHrs(mTin, mTout)
                                            If Trim(.Fields("holiday")) = "" Then
                                                .Fields("absent") = 0.5
                                            End If
                                            .Fields("dayswrk") = 0.5
                                        ElseIf CDate(mT1out) >= CDate(mST2in) And CDate(mT1out) < CDate(mST2out) Then
                                            mTout = mT1out
                                            'compute for 2nd undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST2out))
                                            End If
                                            mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                            .Fields("dayswrk") = 1
                                        Else
                                            mTout = mST2out
                                            mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                            .Fields("dayswrk") = 1
                                        End If
                                        
                                        mNiteWrkHrs = DiffNiteHrs(mTin, mTout, mST1in, mST2out, !nitepremstart, !nitepremend)
                                    ElseIf mT1in = "" And mT2in <> "" Then ' if only the last two(2) time slots were consumed.
                                        If CDate(mT2in) <= CDate(mST2in) Then
                                            mTin = mST2in
                                        Else
                                            mTin = mT2in
                                            If DiffHrs(CDate(mST1in), CDate(mT2in)) > 0 Then
                                                'If DiffHrs(CDate(mST2in), CDate(mT2in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in)) '- mLateAllow
                                                'End If
                                            End If
                                        End If
                                        If CDate(mT2out) < CDate(mST2out) Then
                                            mTout = mT2out
                                             'compute for undertime
                                            If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                            End If
                                        Else
                                            mTout = mST2out
                                        End If
                                        If Trim(.Fields("holiday")) = "" Then
                                            .Fields("absent") = 0.5
                                        End If
                                        .Fields("dayswrk") = 0.5
                                        mNiteWrkHrs = DiffNiteHrs(mTin, mTout, mST2in, mST2out, !nitepremstart, !nitepremend)
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                    Else
                                        If Trim(.Fields("holiday")) = "" Then
                                            If .Fields("required") = "Y" Then
                                                .Fields("absent") = 1
                                            Else
                                                .Fields("absent") = 0
                                            End If
                                            .Fields("dayswrk") = 0
                                        End If
                                    End If
                                    
                                Else 'On travel or On leave
                                    'If on travel or on leave during the second shift, compute only the first shift.
                                    If !firsttravel = 0 And !firstleave = 0 Then
                                        If mT1in <> "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT1in) > CDate(mST1in) Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                    'If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) '- mLateAllow
                                                    'End If
                                                End If
                                            End If
                                            If CDate(mT1out) < CDate(mST1out) Then
                                                If CDate(mT2in) < CDate(mST1out) Then
                                                    'compute for late
                                                    If !brkhrsperday > 0 Then
                                                        If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                                            'If (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) > mLateAllow Then
                                                                If .Fields("latehrs") > 0 Then
                                                                    .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) '- mLateAllow
                                                                Else
                                                                    .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday '- mLateAllow
                                                                End If
                                                            'End If
                                                        End If
                                                    End If
                                                Else
                                                    'compute for undertime
                                                    If CDate(mT1out) < CDate(mST1out) Then
                                                        If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                            .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !secondtravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mST2in, mST1in, mT1out, !nitepremstart, !nitepremend)
                                                mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mST2in, mST2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in <> "" And mT2in = "" Then
                                            'compute for late
                                            If CDate(mT1in) > CDate(mST1in) Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                    'If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) '- mLateAllow
                                                    'End If
                                                End If
                                            End If
                                            'Compute for undertime
                                            If CDate(mT1out) < CDate(mST1out) Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !secondtravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in = "" And mT2in <> "" Then
                                            .Fields("absent") = 0.5
                                            .Fields("dayswrk") = 0.5
                                            If !secondtravel = 1 Then
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                            End If
                                        End If
                                        
                                    'If on travel or on leave during the first shift, compute only the second shift
                                    ElseIf !secondtravel = 0 And !secondleave = 0 Then
                                        If mT1in <> "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT1out) >= CDate(mST2in) Then
                                                If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                    'If DiffHrs(CDate(mT1out), CDate(mT2in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) '- mLateAllow
                                                    'End If
                                                End If
                                            Else
                                                If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                    If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                        'If DiffHrs(CDate(mT2in), CDate(mST1out)) > mLateAllow Then
                                                            .Fields("latehrs") = DiffHrs(CDate(mT2in), CDate(mST1out)) '- mLateAllow
                                                        'End If
                                                    End If
                                                End If
                                            End If
                                            'compute for undertime
                                            If CDate(mT2out) < CDate(mST2out) Then
                                                If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                    .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !firsttravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mST1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                                mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mT2in, mT2out, mST2out, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in <> "" And mT2in = "" Then
                                        
                                            If CDate(mT1in) <= CDate(mST2in) Then
                                                'compute for undertime
                                                If CDate(mT1out) > (mST2in) And CDate(mT1out) < CDate(mST2out) Then
                                                    If DiffHrs(CDate(mT1out), CDate(mST2out)) > 0 Then
                                                        .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST2out))
                                                    End If
                                                    mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                                    If !firsttravel = 1 Then
                                                        .Fields("dayswrk") = 1
                                                        mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                ElseIf CDate(mT1out) > CDate(mST2out) Then
                                                    mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                                    If !firsttravel = 1 Then
                                                        .Fields("dayswrk") = 1
                                                        mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                ElseIf CDate(mT1out) <= CDate(mST2in) Then
                                                    .Fields("absent") = 0.5
                                                    .Fields("daywrk") = 0.5
                                                    If !firsttravel = 1 Then
                                                        mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                End If
                                            End If
                                            If !firsttravel = 1 Then
                                                mNiteWrkHrs = DiffNiteHrs(mST1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                            Else
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in = "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT2in) > CDate(mST2in) Then
                                                If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                    'If DiffHrs(CDate(mST2in), CDate(mT2in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in)) '- mLateAllow
                                                    'End If
                                                End If
                                            End If
                                            
                                            'compute for undertime
                                            If CDate(mT2out) < CDate(mST2out) Then
                                                If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                    .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !firsttravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                .Fields("absent") = 0.5
                                            End If
                                        End If
                                    Else
                                        .Fields("dayswrk") = 1
                                        If !firsttravel = 1 Then
                                            mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                        End If
                                        If !secondtravel = 1 Then
                                            mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                        End If
                                    End If
                                End If
                            Else 'For schedules with only two(2) time slots
                                If !travel = 0 And !leave = 0 Then
                                    If mT1in <> "" Then 'Check if time slots were used.
                                        .Fields("dayswrk") = 1
                                        If CDate(mT1in) < CDate(mST1in) Then
                                            mTin = mST1in
                                        Else
                                            mTin = mT1in
                                            'compute for late
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                'If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) '- mLateAllow
                                                'End If
                                            End If
                                        End If
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            mTout = mT1out
                                            'compute for undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                        Else
                                            mTout = mST1out
                                        End If
                                        mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                        
                                        'mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    Else 'absent
                                        If Trim(.Fields("holiday")) = "" Then
                                            If .Fields("required") = "Y" Then
                                                .Fields("absent") = 1
                                            End If
                                            .Fields("dayswrk") = 0
                                        End If
                                    End If
                                End If
                            End If
                            .Fields("wrkhrs") = mWrkHrs - mNiteWrkHrs
                            .Fields("nitewrkhrs") = mNiteWrkHrs
                        End If
                        
                        If Trim(!hrsperday) = "" Then !hrsperday = 0
                        If Trim(!brkhrsperday) = "" Then !brkhrsperday = 0

                        ConMain.Execute "delete from dtremp where employeecode = '" & mEmpNo & "' and workdate ='" & Format(!wrkdate, "YYYY-MM-DD") & "'"
                        
                        ConMain.Execute "insert into dtremp(employeecode,payfreqcode,dayno,workdate,shiftcode, " & _
                                              "t1in,t1out,t2in,t2out,st1in,st1out,st2in,st2out, " & _
                                              "wrkhrs,nitewrkhrs,dayswrk,absent,latehrs,uthrs,dayoff,updatable,brkstart,brkend, " & _
                                              "nitepremstart,nitepremend,hrsperday,brkhrsperday,required, " & _
                                              "branchcode,divisioncode,costcentercode,sectioncode, holiday, " & _
                                              "firstleave,secondleave) values " & _
                                              "('" & mEmpNo & "','" & mPayFreqCode & "','" & !dayno & "','" & Format(!wrkdate, "YYYY-MM-DD") & "','" & !shiftcode & "', " & _
                                              "'" & Format(!t1in, "hh:nn") & "','" & Format(!t1out, "hh:nn") & "','" & Format(!t2in, "hh:nn") & "','" & Format(!t2out, "hh:nn") & "', " & _
                                              "'" & !st1in & "','" & !st1out & "','" & !st2in & "','" & !st2out & "', " & _
                                              !wrkhrs & "," & !nitewrkhrs & "," & !dayswrk & "," & !absent & "," & !latehrs & "," & !uthrs & ",'" & IIf(!dayoff <> 0, "Y", "N") & "','" & IIf(!updatable <> 0, "Y", "N") & "','" & !brkstart & "','" & !brkend & "', " & _
                                              "'" & !nitepremstart & "','" & !nitepremend & "'," & !hrsperday & "," & !brkhrsperday & ",'" & !Required & "', " & _
                                              "'" & rsEmployee!branchcode & "','" & rsEmployee!divisioncode & "','" & rsEmployee!costcentercode & "','" & rsEmployee!sectioncode & "','" & !Holiday & "', " & _
                                              "'" & !firstleave & "','" & !secondleave & "')"

                        .MoveNext
                        DoEvents
                        
                    Loop
                    
                    .MoveFirst
                    
                End With
                
            End If
            
            'Computes Overtime
            
            Set rsOTTmp = Nothing
            Set rsOTTmp = New ADODB.Recordset
            
            With rsOTTmp
                .Fields.Append "otlneno", adVarChar, 7
                .Fields.Append "otcode", adVarChar, 15
                .Fields.Append "employeecode", adVarChar, 15
                .Fields.Append "percode", adVarChar, 7
                .Fields.Append "dayoff", adInteger
                .Fields.Append "holiday", adVarChar, 10
                .Fields.Append "wrkdate", adDate
                .Fields.Append "day", adVarChar, 15
                .Fields.Append "actotstart", adVarChar, 8
                .Fields.Append "actotend", adVarChar, 8
                .Fields.Append "otstart", adVarChar, 5
                .Fields.Append "otend", adVarChar, 5
                .Fields.Append "otwrkhrs", adDouble, 18
                .Fields.Append "regwrkhrs", adDouble, 18
                .Fields.Append "nitewrkhrs", adDouble, 18
                .Open
            End With
            
            NetOpen rsOT, "select * from overtimelne where employeecode = '" & mEmpNo & "' and status = 'Approved' and percode = '" & mPerCode & "' and fnlz <> 'Y' order by wrkdate,otstart"
            
            With rsOTTmp
                If rsOT.RecordCount > 0 Then
                    rsOT.MoveFirst
                    
'                    pb2.Value = 0
'                    pb2.Max = rsOT.RecordCount
                    Do While Not rsOT.EOF
                    
                        Set rsOTCmpr = New ADODB.Recordset
                        Set rsOTCmpr = rsOTTmp.Clone
                        
'                        If Format(rsOT!wrkdate, "YYYY-MM-DD") = "2008-04-08" Then
'                            MsgBox ""
'                        End If
                        
                        'pb2.Value = pb2.Value + 1
                        
                        mOT_Start = ""
                        mOT_End = ""
                        mNiteStart = ""
                        mNiteEnd = ""
                        mTimeIN = ""
                        mTimeOUT = ""
                        
                        .AddNew
                        .Fields("otlneno") = rsOT!otlneno
                        .Fields("otcode") = rsOT!otcode
                        .Fields("employeecode") = rsOT!employeecode
                        .Fields("percode") = rsOT!percode
                        .Fields("wrkdate") = rsOT!wrkdate
                        .Fields("day") = WeekdayName(Weekday(rsOT!wrkdate))
                        .Fields("otstart") = rsOT!otstart
                        .Fields("otend") = rsOT!otend
                        
                        'Check if dayoff
                        NetOpen rsDtrEmp, "select * from dtremp where employeecode = '" & mEmpNo & "' and workdate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "' "
                        
                        If rsDtrEmp.RecordCount > 0 Then
                            .Fields("dayoff") = IIf(rsDtrEmp!dayoff = "Y", 1, 0)
                            'Check if it has a night premium hours
                            If Trim(rsDtrEmp!nitepremstart) <> "" Then
                                mNiteStart = Format(rsOT!wrkdate & " " & rsDtrEmp!nitepremstart, "MM/DD/YYYY hh:nn:ss")
                                If CDate(mNiteStart) > Format(rsOT!wrkdate & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss") Then
                                    mNiteEnd = Format(CDate(rsOT!wrkdate) + 1 & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss")
                                Else
                                    mNiteEnd = Format(CDate(rsOT!wrkdate) & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss")
                                End If
                            End If
                        End If
                        
                        'Check if Holiday
                        NetOpen rsHoliday, "select * from holiday where holidaydate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
                        If rsHoliday.RecordCount > 0 Then
                            If CInt(rsHoliday!regular) = 1 Then
                              .Fields("holiday") = "Legal"
                            Else
                              .Fields("holiday") = "Special"
                            End If
                        Else
                            .Fields("holiday") = ""
                        End If
                        
                        Create_TmpOTDTSTito mEmpNo, rsOT!wrkdate, mPerCode
                        
                        mOT_Start = Format(rsOT!wrkdate & " " & rsOT!otstart, "MM/DD/YYYY hh:nn:ss")
                        
                        'Check if Overtime out is on the next day
                        
                        If CDate(mOT_Start) > Format(rsOT!wrkdate & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss") Then
                            mOT_End = Format(CDate(rsOT!wrkdate) + 1 & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss")
                        Else
                            mOT_End = Format(rsOT!wrkdate & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss")
                        End If
                        
                        If rsTitoTmp.RecordCount > 0 Then
                            
                            rsTitoTmp.Filter = "tout > '" & mOT_Start & "' and tin < '" & mOT_End & "'"
                            
                            If Not rsTitoTmp.EOF Then
                            
                                rsTitoTmp.MoveFirst
                                
                                !otwrkhrs = 0
                                !regwrkhrs = 0
                                !nitewrkhrs = 0
                                
                                .Fields("actotstart") = Format(rsTitoTmp!tin, "hh:nn:ss")
                                
                                Do While Not rsTitoTmp.EOF
                                
                                    .Fields("actotend") = Format(rsTitoTmp!tout, "hh:nn:ss")
                                    
                                    If CDate(rsTitoTmp!tin) < CDate(mOT_Start) Then
                                        mOTActIN = mOT_Start
                                    Else
                                        mOTActIN = rsTitoTmp!tin
                                    End If
                                    
                                    If CDate(rsTitoTmp!tout) > CDate(mOT_End) Then
                                        mOTActOUT = mOT_End
                                    Else
                                        mOTActOUT = rsTitoTmp!tout
                                    End If
                                    
                                    !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                                              
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                            
                                        End If
                                    End If
                                    
                                    rsTitoTmp.MoveNext
                                    
                                Loop
                            End If
                        End If
                        
                        'Computes manual TITO using Employee's DTR
                        
                        'only if Employee's TITO is not available from TITO TABLE
                        
                        If Trim(!actotstart) = "" Then
                            
                            !otwrkhrs = 0
                            !regwrkhrs = 0
                            !nitewrkhrs = 0
                            
                            If rsDtrEmp.RecordCount > 0 Then
                            
                                If Trim(rsDtrEmp!t1in) <> "" Then
                                    
                                    .Fields("actotstart") = rsDtrEmp!t1in
                                    .Fields("actotend") = rsDtrEmp!t1out
                                    
                                    mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t1in, "MM/DD/YYYY hh:nn:ss")
                                    If CDate(mTimeIN) > Format(rsOT!wrkdate & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss") Then
                                        mTimeOUT = Format(CDate(rsOT!wrkdate) + 1 & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss")
                                    Else
                                        mTimeOUT = Format(rsOT!wrkdate & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeOUT) > CDate(mOT_Start) And CDate(mTimeIN) < CDate(mOT_End) Then
                                        If CDate(mTimeIN) < CDate(mOT_Start) Then
                                            mOTActIN = mOT_Start
                                        Else
                                            mOTActIN = mTimeIN
                                        End If
                                        
                                        If CDate(mTimeOUT) > CDate(mOT_End) Then
                                            mOTActOUT = mOT_End
                                        Else
                                            mOTActOUT = mTimeOUT
                                        End If
                                        !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                    End If
                                    
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                        End If
                                    End If
                                    
                                End If
                                                                                
                                If Trim(rsDtrEmp!t2in) <> "" Then
                                
                                    If Trim(rsDtrEmp!t1in) = "" Then
                                        .Fields("actotstart") = rsDtrEmp!t2in
                                    End If
                                    .Fields("actotend") = rsDtrEmp!t2out
                                    
                                    If Trim(mTimeOUT) <> "" Then
                                        If CDate(mTimeOUT) > Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss") Then
                                            mTimeIN = Format(CDate(Format(mTimeOUT, "MM/DD/YYYY")) + 1 & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                        Else
                                            mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                        End If
                                    Else
                                        mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeIN) > Format(rsOT!wrkdate & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss") Then
                                        mTimeOUT = Format(CDate(Format(mTimeIN, "MM/DD/YYYY")) + 1 & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss")
                                    Else
                                        mTimeOUT = Format(rsOT!wrkdate & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeOUT) > CDate(mOT_Start) And CDate(mTimeIN) < CDate(mOT_End) Then
                                        If CDate(mTimeIN) < CDate(mOT_Start) Then
                                            mOTActIN = mOT_Start
                                        Else
                                            mOTActIN = mTimeIN
                                        End If
                                        
                                        If CDate(mTimeOUT) > CDate(mOT_End) Then
                                            mOTActOUT = mOT_End
                                        Else
                                            mOTActOUT = mTimeOUT
                                        End If
                                        !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                    End If
                                    
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                        End If
                                        
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        If !otwrkhrs >= mOTAllow Then
                            !otwrkhrs = !otwrkhrs - !nitewrkhrs
                        Else
                            !otwrkhrs = 0
                            !nitewrkhrs = 0
                        End If
                        
                        mLackHrs = 0
                        mTtlReghrs = 0
                        
                        
'                        If Format(rsOT!wrkdate, "YYYY-MM-DD") = "2008-04-08" Then
'                            MsgBox ""
'                        End If
                        
                        ' Checks the no. of hours work by the employee if it reaches the required no of working hours per day.
                        ' If not deduct the no. of regular overtime hours ONLY if any.
                        
                        If !otwrkhrs > 0 Then
                            NetOpen rsOTChk, "select (hrsperday - wrkhrs) lackhrs,hrsperday,firstleave,secondleave from dtremp where hrsperday > wrkhrs and employeecode = '" & mEmpNo & "' and workdate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
                            If rsOTChk.RecordCount > 0 Then
                                
                                If rsOTChk!lackhrs >= mUTMin Then
                                
                                    mLackHrs = rsOTChk!lackhrs
                                    
                                    If rsOTChk!firstleave = 1 Then
                                        mLackHrs = mLackHrs - Format(rsOTChk!hrsperday / 2, "#,##0.00")
                                    End If
                                    
                                    If rsOTChk!secondleave = 1 Then
                                        mLackHrs = mLackHrs - Format(rsOTChk!hrsperday / 2, "#,##0.00")
                                    End If
                                    
                                    If mLackHrs > 0 Then
                                        
                                        If rsOTCmpr.RecordCount > 0 Then
                                            rsOTCmpr.MoveFirst
                                            Do While Not rsOTCmpr.EOF
                                                If rsOTCmpr!otlneno <> !otlneno And rsOTCmpr!wrkdate = rsOT!wrkdate Then
                                                    mTtlReghrs = mTtlReghrs + rsOTCmpr!regwrkhrs
                                                End If
                                                rsOTCmpr.MoveNext
                                            Loop
                                        End If
                                        
                                        If mLackHrs > mTtlReghrs Then
                                            mLackHrs = mLackHrs - mTtlReghrs
                                            If mLackHrs > !otwrkhrs Then
                                                !regwrkhrs = !otwrkhrs
                                                !otwrkhrs = 0
                                            Else
                                                !regwrkhrs = mLackHrs
                                                !otwrkhrs = !otwrkhrs - mLackHrs
                                            End If
                                        
                                        End If
                                    
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        .Update
                        
                        ConMain.Execute "update overtimelne set actotstart = '" & !actotstart & "', actotend = '" & !actotend & "',otwrkhrs = " & !otwrkhrs & ", regwrkhrs = " & !regwrkhrs & ",nitewrkhrs = " & !nitewrkhrs & ", holiday = '" & !Holiday & "',dayoff = '" & IIf(!dayoff <> 0, "Y", "N") & "' " & _
                                              "where otlneno = '" & !otlneno & "'"
                        
                        rsOT.MoveNext
                    Loop
                    
                End If
                                
            End With
            
            rsEmployee.MoveNext
            
        Loop
                
        'ConMain.CommitTrans
'        fra1.Enabled = True
'        cmdGenerate.Enabled = True
'        Me.MousePointer = vbDefault
'
'        MsgBox "Process completed!", vbInformation + vbOKOnly
        
'        pb1.Value = 0
'        pb2.Value = 0
'    Else
'        MsgBox "No record found.", vbExclamation + vbOKOnly
    End If
    
End Sub

Private Sub Create_TmpTito(mEmpNo As String, mStartDate As String, mEndDate As String, mPerCode As String)
    
    Dim rsTito      As ADODB.Recordset
    Dim mDateTmp    As String
    Dim isIn        As Boolean

    Set rsTitoTmp = Nothing
    Set rsTitoTmp = New ADODB.Recordset

    With rsTitoTmp
        .Fields.Append "wrkdate", adDate
        .Fields.Append "tin", adDate
        .Fields.Append "tout", adDate
        .Open
        .Sort = "wrkdate"
    End With

'    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
'                        "from tito where employeecode = '" & mEmpno & "' and " & _
'                        "datelog Between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "'  and " & _
'                        "'" & Format(CDate(tdbPayrollPeriod.Columns("wrkdateto").Text) + 1, "YYYY-MM-DD") & "' order by complog"
                        
                        
    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
                        "from tito where employeecode = '" & mEmpNo & "' and " & _
                        "datelog Between '" & mStartDate & "'  and " & _
                        "'" & Format(CDate(mEndDate) + 1, "YYYY-MM-DD") & "' " & _
                        "Union All " & _
                        "select employeecode,complog,datelog,timelog,logstat " & _
                        "from gplne where employeecode = '" & mEmpNo & "' and " & _
                        "(datelog Between '" & mStartDate & "' and " & _
                        "'" & Format(CDate(mEndDate) + 1, "YYYY-MM-DD") & "') and  " & _
                        "percode = '" & mPerCode & "' and status = 'Approved' " & _
                        "order by complog"
                        
    With rsTito
        If .RecordCount > 0 Then

            .MoveFirst
'            pb2.Max = .RecordCount
'            pb2.Value = 0
            If !logstat = "Out" Then
                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
            End If
            isIn = False
            
            Do While Not .EOF
                
                'pb2.Value = pb2.Value + 1
                
                If !logstat = "In" Then
                
                    If Not isIn Then
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    Else
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    End If
                    
                Else
                    If isIn Then
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    Else
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = Format(mDateTmp, "MM/DD/YYYY")
                        rsTitoTmp.Fields("tin") = mDateTmp
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    End If
                End If
                .MoveNext
                DoEvents
            Loop

        End If

    End With
    
End Sub


Private Sub Load_Dtr(mEmpNo As String, mStartDate As String, mEndDate As String)
 
    Dim mDate       As Date
    
    Dim rsEmpDtr    As ADODB.Recordset
    Dim rsEmpShift  As ADODB.Recordset
    Dim rsEmpShift2 As ADODB.Recordset
    Dim rsHoliday   As ADODB.Recordset
    Dim rsOBT       As ADODB.Recordset
    Dim rsLeave     As ADODB.Recordset
    
    Set rsDtrTmp = Nothing
    Set rsDtrTmp = New ADODB.Recordset
    
    With rsDtrTmp
  
        .Fields.Append "updatable", adInteger
        .Fields.Append "wrkdate", adDate
        .Fields.Append "dayno", adInteger
        .Fields.Append "day", adVarChar, 15
        .Fields.Append "dayoff", adInteger
        .Fields.Append "dayswrk", adDouble, 18
        .Fields.Append "holiday", adVarChar, 10
        .Fields.Append "Travel", adInteger
        .Fields.Append "Leave", adInteger
        .Fields.Append "t1in", adVarChar, 15, adFldIsNullable
        .Fields.Append "t1out", adVarChar, 15, adFldIsNullable
        .Fields.Append "t2in", adVarChar, 15, adFldIsNullable
        .Fields.Append "t2out", adVarChar, 15, adFldIsNullable
        .Fields.Append "st1in", adVarChar, 5
        .Fields.Append "st1out", adVarChar, 5
        .Fields.Append "st2in", adVarChar, 5
        .Fields.Append "st2out", adVarChar, 5
        .Fields.Append "brkstart", adVarChar, 5
        .Fields.Append "brkend", adVarChar, 5
        .Fields.Append "nitepremstart", adVarChar, 5
        .Fields.Append "nitepremend", adVarChar, 5
        .Fields.Append "shiftcode", adVarChar, 7
        .Fields.Append "shiftdetail", adVarChar, 50
        .Fields.Append "wrkhrs", adDouble
        .Fields.Append "nitewrkhrs", adDouble
        .Fields.Append "absent", adDouble
        .Fields.Append "latehrs", adDouble
        .Fields.Append "uthrs", adDouble
        .Fields.Append "hrsperday", adDouble, 18
        .Fields.Append "brkhrsperday", adDouble, 18
        .Fields.Append "firsttravel", adVarChar, 1
        .Fields.Append "secondtravel", adVarChar, 1
        .Fields.Append "firstleave", adVarChar, 1
        .Fields.Append "secondleave", adVarChar, 1
        .Fields.Append "required", adVarChar, 1
        .Open

        mDate = mStartDate
'        pb2.Max = DateDiff("d", mdate, Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")) + 1
'        pb2.Value = 0
        
        Do While mDate <= CDate(Format(mEndDate, "MM/DD/YYYY"))
            
            'pb2.Value = pb2.Value + 1
            .AddNew
            .Fields("wrkdate") = mDate
            .Fields("dayno") = Weekday(mDate)
            .Fields("day") = WeekdayName(Weekday(mDate))
            
            NetOpen rsEmpShift, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from empshift x1 left outer join shift x2 on " & _
                                  "x1.shiftcode = x2.shiftcode where x1.shiftcode <> '' and x1.employeecode = '" & mEmpNo & "' and x1.dayno = '" & Weekday(mDate) & "'"
            
            If rsEmpShift.RecordCount > 0 Then
        
                .Fields("updatable") = 1
                .Fields("t1in") = ""
                .Fields("t1out") = ""
                .Fields("t2in") = ""
                .Fields("t2out") = ""
                .Fields("st1in") = rsEmpShift!t1in
                .Fields("st1out") = rsEmpShift!t1out
                .Fields("st2in") = rsEmpShift!t2in
                .Fields("st2out") = rsEmpShift!t2out
                .Fields("shiftcode") = rsEmpShift!shiftcode
                .Fields("shiftdetail") = rsEmpShift!t1in & "   " & rsEmpShift!t1out & "       " & rsEmpShift!t2in & "   " & rsEmpShift!t2out
                .Fields("brkstart") = rsEmpShift!brkstart
                .Fields("brkend") = rsEmpShift!brkend
                .Fields("nitepremstart") = rsEmpShift!nitepremstart
                .Fields("nitepremend") = rsEmpShift!nitepremend
                .Fields("hrsperday") = rsEmpShift!hrsperday
                .Fields("brkhrsperday") = rsEmpShift!brkhrsperday
                .Fields("dayoff") = 0
                .Fields("required") = rsEmpShift!Required
                
                
                NetOpen rsEmpDtr, "select * from dtremp where employeecode = '" & mEmpNo & "' and " & _
                                    "workdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
                                      
                If rsEmpDtr.RecordCount > 0 Then
                    NetOpen rsEmpShift2, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from  shift x2 where x2.shiftcode = '" & rsEmpDtr!shiftcode & "'"
                    If rsEmpShift2.RecordCount > 0 Then
                        If rsEmpDtr!updatable = "N" Then
                            .Fields("updatable") = 0
                            .Fields("t1in") = rsEmpDtr!t1in
                            .Fields("t1out") = rsEmpDtr!t1out
                            .Fields("t2in") = rsEmpDtr!t2in
                            .Fields("t2out") = rsEmpDtr!t2out
                            
                            .Fields("st1in") = rsEmpShift2!t1in
                            .Fields("st1out") = rsEmpShift2!t1out
                            .Fields("st2in") = rsEmpShift2!t2in
                            .Fields("st2out") = rsEmpShift2!t2out
                            
                        End If
                        
                        .Fields("shiftcode") = rsEmpShift2!shiftcode
                        .Fields("shiftdetail") = rsEmpShift2!t1in & "   " & rsEmpShift2!t1out & "       " & rsEmpShift2!t2in & "   " & rsEmpShift2!t2out
                        .Fields("brkstart") = rsEmpShift2!brkstart
                        .Fields("brkend") = rsEmpShift2!brkend
                        .Fields("nitepremstart") = rsEmpShift2!nitepremstart
                        .Fields("nitepremend") = rsEmpShift2!nitepremend
                        .Fields("hrsperday") = rsEmpShift2!hrsperday
                        .Fields("brkhrsperday") = rsEmpShift2!brkhrsperday
                        .Fields("required") = rsEmpShift2!Required
                    Else
                        .Fields("st1in") = ""
                        .Fields("st1out") = ""
                        .Fields("st2in") = ""
                        .Fields("st2out") = ""
                        .Fields("shiftcode") = ""
                        .Fields("shiftdetail") = ""
                        .Fields("brkstart") = ""
                        .Fields("brkend") = ""
                        .Fields("nitepremstart") = ""
                        .Fields("nitepremend") = ""
                        .Fields("hrsperday") = 0
                        .Fields("brkhrsperday") = 0
                        .Fields("dayoff") = 1
                        .Fields("required") = "N"
                    End If
                End If
            Else
            
                NetOpen rsEmpDtr, "select * from dtremp where employeecode = '" & mEmpNo & "' and " & _
                                    "workdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
                                      
                If rsEmpDtr.RecordCount > 0 Then
                    NetOpen rsEmpShift2, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from  shift x2 where x2.shiftcode = '" & rsEmpDtr!shiftcode & "'"
                    If rsEmpShift2.RecordCount > 0 Then
                        If rsEmpDtr!updatable = "N" Then
                            .Fields("updatable") = 0
                            .Fields("t1in") = rsEmpDtr!t1in
                            .Fields("t1out") = rsEmpDtr!t1out
                            .Fields("t2in") = rsEmpDtr!t2in
                            .Fields("t2out") = rsEmpDtr!t2out
                            
                            .Fields("st1in") = rsEmpShift2!t1in
                            .Fields("st1out") = rsEmpShift2!t1out
                            .Fields("st2in") = rsEmpShift2!t2in
                            .Fields("st2out") = rsEmpShift2!t2out
                            
                        End If
                        
                        .Fields("shiftcode") = rsEmpShift2!shiftcode
                        .Fields("shiftdetail") = rsEmpShift2!t1in & "   " & rsEmpShift2!t1out & "       " & rsEmpShift2!t2in & "   " & rsEmpShift2!t2out
                        .Fields("brkstart") = rsEmpShift2!brkstart
                        .Fields("brkend") = rsEmpShift2!brkend
                        .Fields("nitepremstart") = rsEmpShift2!nitepremstart
                        .Fields("nitepremend") = rsEmpShift2!nitepremend
                        .Fields("hrsperday") = rsEmpShift2!hrsperday
                        .Fields("brkhrsperday") = rsEmpShift2!brkhrsperday
                        .Fields("required") = rsEmpShift2!Required
                        .Fields("dayoff") = IIf(rsEmpDtr!dayoff = "Y", 1, 0)
                    Else
                    
                        .Fields("updatable") = IIf(rsEmpDtr!updatable = "N", 0, 1)
                        .Fields("t1in") = rsEmpDtr!t1in
                        .Fields("t1out") = rsEmpDtr!t1out
                        .Fields("t2in") = rsEmpDtr!t2in
                        .Fields("t2out") = rsEmpDtr!t2out
                        .Fields("st1in") = rsEmpDtr!st1in
                        .Fields("st1out") = rsEmpDtr!st1out
                        .Fields("st2in") = rsEmpDtr!st2in
                        .Fields("st2out") = rsEmpDtr!st2out
                        .Fields("shiftcode") = ""
                        .Fields("shiftdetail") = IIf(Trim(rsEmpDtr!st1in) <> "", rsEmpDtr!st1in, "") & "   " & IIf(Trim(rsEmpDtr!st1out) <> "", rsEmpDtr!st1out, "") & "       " & IIf(Trim(rsEmpDtr!st2in) <> "", rsEmpDtr!st2in, "") & "   " & IIf(Trim(rsEmpDtr!st2out) <> "", rsEmpDtr!st2out, "")
                        .Fields("brkstart") = rsEmpDtr!brkstart
                        .Fields("brkend") = rsEmpDtr!brkend
                        .Fields("nitepremstart") = rsEmpDtr!nitepremstart
                        .Fields("nitepremend") = rsEmpDtr!nitepremend
                        .Fields("hrsperday") = rsEmpDtr!hrsperday
                        .Fields("brkhrsperday") = rsEmpDtr!brkhrsperday
                        .Fields("required") = IIf(rsEmpDtr!Required & "" = "Y", "Y", "N")
                        .Fields("dayoff") = IIf(rsEmpDtr!dayoff = "Y", 1, 0)
                    End If
                Else
                    .Fields("updatable") = 1
                    .Fields("dayoff") = 1
                    .Fields("required") = "N"
                End If
                  
            End If
              
            NetOpen rsHoliday, "select x1.* from holiday x1 " & _
                    "left outer join holidaybranchinclude x2 on x1.holidaydate = x2.holidaydate " & _
                    "where x1.holidaydate = '" & Format(mDate, "YYYY-MM-DD") & "' and x2.branchcode = '" & mEmpNo & "'"
            
            If rsHoliday.RecordCount > 0 Then
                If CInt(rsHoliday!regular) = 1 Then
                  .Fields("holiday") = "Legal"
                Else
                  .Fields("holiday") = "Special"
                End If
            Else
                .Fields("holiday") = ""
            End If
            
            NetOpen rsOBT, "select x1.* from obtlne x1 left outer join obthdr x2 on x1.obtnum = x2.obtnum " & _
                             "where x2.employeecode = '" & mEmpNo & "' and x1.obtdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
            
            If rsOBT.RecordCount > 0 Then
                .Fields("travel") = 1
                .Fields("firsttravel") = rsOBT!firstshift
                .Fields("secondtravel") = rsOBT!secondshift
            Else
                .Fields("travel") = 0
                .Fields("firsttravel") = 0
                .Fields("secondtravel") = 0
            End If
            
            NetOpen rsLeave, "select x1.* from lvlne x1 " & _
                    "left outer join lvhdr x2 on x1.lvnum = x2.lvnum " & _
                    "where x2.employeecode = '" & mEmpNo & "' and x1.lvdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
            
            If rsLeave.RecordCount > 0 Then
                .Fields("leave") = 1
                .Fields("firstleave") = rsLeave!firstshift
                .Fields("secondleave") = rsLeave!secondshift
            Else
                .Fields("leave") = 0
                .Fields("firstleave") = 0
                .Fields("secondleave") = 0
            End If
            
            .Update
            
            mDate = mDate + 1
            DoEvents
        Loop
      
    End With

End Sub

Private Function DiffHrs(mHrs1 As Date, mHrs2 As Date) As Double
    DiffHrs = Format(Round(DateDiff("N", mHrs1, mHrs2) / 60, 2), "#,##0.00")
End Function


Private Function DiffNiteHrs(ByVal mActIn As Variant, ByVal mActOut As Variant, ByVal mSTin As String, ByVal mSTout As String, ByVal mNiteStart As String, ByVal mNiteEnd As String) As Double
    
    Dim mNiteIN     As Date
    Dim mNiteOUT    As Date
    Dim mDumIn      As Date
    Dim mDumOut     As Date
    
    DiffNiteHrs = 0
    
    If Trim(mNiteStart) <> "" Then
        If CDate(mActIn) < CDate(mSTin) Then
            mDumIn = CDate(mSTin)
        Else
            mDumIn = CDate(mActIn)
        End If
        If CDate(mActOut) > CDate(mSTout) Then
            If mDumIn > CDate(mSTout) Then
                mDumOut = mDumIn
            Else
                mDumOut = CDate(mSTout)
            End If
        Else
            mDumOut = CDate(mActOut)
        End If
        If mDumOut > CDate(mNiteStart) And mDumIn < CDate(mNiteEnd) Then
            If mDumIn < CDate(mNiteStart) Then
                mNiteIN = mNiteStart
            Else
                mNiteIN = mDumIn
            End If
            If mDumOut > CDate(mNiteEnd) Then
                mNiteOUT = mNiteEnd
            Else
                mNiteOUT = mDumOut
            End If
            DiffNiteHrs = DiffHrs(mNiteIN, mNiteOUT)
        End If
    End If
    
End Function

Private Sub Create_TmpOTDTSTito(mEmpNo As String, mDate As Date, mPerCode As String)
    
    Dim rsTito      As ADODB.Recordset
    Dim mDateTmp    As String
    Dim isIn        As Boolean

    Set rsTitoTmp = Nothing
    Set rsTitoTmp = New ADODB.Recordset

    With rsTitoTmp
        .Fields.Append "wrkdate", adDate
        .Fields.Append "tin", adDate
        .Fields.Append "tout", adDate
        .Open
        .Sort = "wrkdate"
    End With


'    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
'                        "from tito where employeecode = '" & mEmpno & "' and " & _
'                        "datelog Between '" & Format(mdate - 1, "YYYY-MM-DD") & "'  and " & _
'                        "'" & Format(mdate + 1, "YYYY-MM-DD") & "' order by complog"
    
    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
                        "from tito where employeecode = '" & mEmpNo & "' and " & _
                        "datelog Between '" & Format(mDate - 1, "YYYY-MM-DD") & "'  and " & _
                        "'" & Format(mDate + 1, "YYYY-MM-DD") & "' " & _
                        "Union All " & _
                        "select employeecode,complog,datelog,timelog,logstat " & _
                        "from gplne where employeecode = '" & mEmpNo & "' and " & _
                        "(datelog Between '" & Format(mDate - 1, "YYYY-MM-DD") & "' and " & _
                        "'" & Format(mDate + 1, "YYYY-MM-DD") & "') and  " & _
                        "percode = '" & mPerCode & "' and status = 'Approved' " & _
                        "order by complog"
    
    With rsTito
        If .RecordCount > 0 Then

            .MoveFirst
            
            If !logstat = "Out" Then
                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
            End If
            isIn = False
            
            Do While Not .EOF
                If !logstat = "In" Then
                
                    If Not isIn Then
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    Else
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    End If
                    
                Else
                
                    If isIn Then
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    Else
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = Format(mDateTmp, "MM/DD/YYYY")
                        rsTitoTmp.Fields("tin") = mDateTmp
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    End If
                    
                End If
                .MoveNext
                DoEvents
            Loop
            rsTitoTmp.MoveLast
            If Trim(rsTitoTmp!tout) = "" Then
                rsTitoTmp.Delete
            End If

        End If

    End With
    
End Sub


