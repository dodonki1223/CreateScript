'**************************************************************************************
'* �v���O������ �F �X�^�[�g�A�b�v�v���O�����ꊇ�I���X�N���v�g                         *
'* �����T�v     �F �X�^�[�g�A�b�v�ŋN��������v���O�������P���I�����邩�ǂ������[ *
'*                 �U�[�ɑΘb�Ŋm�F�i���b�Z�[�W�{�b�N�X�̂͂��A�������j�B�u�͂��v��   *
'*                 �I�����ꂽ�v���O���������ׂďI�����܂��B                           *
'*                 ��Clibor�̓N���b�v�{�[�h�̗������ۑ�����Ȃ��̂Œ��ӁI�I           *
'*                   T-Clock�Ɋւ��Ă͒ʏ�̂����ł͂��܂��I���ł��Ȃ������̂Ō�  *
'*                   �ŏI���������L�q����                                             *
'* ����         �F Stickies�����܂��I���ł����ɗ����鎖������̂ŏC������K�v����     *
'* �ݒ�         �F                                                                    *
'**************************************************************************************

Main()

'***********************************************************************
'* ������   �F ���C������                                              *
'* ����     �F �Ȃ�                                                    *
'* �������e �F ���C������                                              *
'* �߂�l   �F �Ȃ�                                                    *
'***********************************************************************
Sub Main()

    '*****************************************
    '* �������s�����[�U�[�ɑΘb              *
    '*****************************************
    '���b�Z�[�W�̕\��
    Dim mContinueProcessingResult : mContinueProcessingResult = MsgBox("�X�^�[�g�A�b�v�v���O�����ꊇ�I�����������s���܂��B" & vbCrLf & "��낵���ł����H", vbOKCancel, "�X�^�[�g�A�b�v�v���O�����ꊇ�I������")

    '�L�����Z���������ꂽ���͏������I��
    If mContinueProcessingResult = vbCancel Then Wscript.Quit()

    '*****************************************
    '* �I���v���O�����̑I�菈��              *
    '*****************************************
    '�I������v���O�����i�[Dictionary
    Dim mExitExes : Set mExitExes = WScript.CreateObject("Scripting.Dictionary")

    '�I������v���O�����̒ǉ�����
    Set mExitExes = AddExitExe(mExitExes)

    'T-Clock�̏I���� ���ʏ�̂����ł̓v���Z�X���I���ł��Ȃ��̂�TClock�ɂ��Ă͌ʂőΉ�����
    Dim mTClockExitResult
    Dim mIsRunTClock : mIsRunTClock = IsRunProgram("Clock64.exe")
    If mIsRunTClock = True Then

        mTClockExitResult  = GetSelectedUserResultForExitProgram("T-Clock���I�����܂����H", "T-Clock�I����")

    End If

    '*****************************************
    '* �v���O�����̈ꊇ�I������              *
    '*****************************************
    'Dictionary�Ɋi�[����Ă���v���O���������J��Ԃ�
    For Each mExeName In mExitExes.Keys

        '�v���O�����̏I������
        For Each Process in GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process where Name='" & mExitExes(mExeName) & "'")

            WScript.Sleep 1000
            Process.terminate

        Next

    Next

    '*****************************************
    '* T-Clock�̏I�������i���ʂőΉ�����j *
    '*****************************************
    If mTClockExitResult = vbYes Then

        '���s�h���C�u���擾����
        Dim mObjFso : set mObjFso = WScript.CreateObject("Scripting.FileSystemObject") 'FileSystemObject
        Dim mRunDrive : mRunDrive = mObjFso.GetDriveName(WScript.ScriptFullName)

        'T-Clock�̃t���p�X���擾�i�I���̃R�}���h���C���������܂߂āj
        Dim mTClockPath : mTClockPath = mRunDrive & "\Tools\T-Clock\Clock64.exe /exit"

        'T-Clock�̏I�������s
        Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell")
        mObjShell.Run mTClockPath

    End If

End Sub

'***********************************************************************
'* ������   �F �I���Ώۃv���O�����̒ǉ�����                            *
'* ����     �F pExitExes �I������v���O�����i�[Dictionary              *
'* �������e �F �I���Ώۃv���O������Dictionary�ɒǉ�����                *
'*             ���b�Z�[�W�{�b�N�X�łP���m�F���Ă���                *
'* �߂�l   �F pExitExes                                               *
'***********************************************************************
Function AddExitExe(ByVal pExitExes)

    'AkabeiMonitor�̏I����
    Dim mIsRunAkabeiMonitor : mIsRunAkabeiMonitor = IsRunProgram("akamoni.exe")
    If mIsRunAkabeiMonitor = True Then

        Dim mAkabeiMonitorExitResult : mAkabeiMonitorExitResult = GetSelectedUserResultForExitProgram("AkabeiMonitor���I�����܂����H", "AkabeiMonitor�I����")
        If mAkabeiMonitorExitResult = vbYes Then

            pExitExes.Add "AkabeiMonitor", "akamoni.exe"

        End If

    End If

    'AutoHotKeyTool�̏I����
    Dim mIsRunAutoHotKeyTool : mIsRunAutoHotKeyTool = IsRunProgram("AutoHotKeyTool.exe")
    If mIsRunAutoHotKeyTool = True Then

        Dim mAutoHotKeyToolExitResult : mAutoHotKeyToolExitResult  = GetSelectedUserResultForExitProgram("AutoHotKeyTool���I�����܂����H", "AutoHotKeyTool�I����")
        If mAutoHotKeyToolExitResult = vbYes Then

            pExitExes.Add "AutoHotKeyTool", "AutoHotKeyTool.exe"

        End If

    End If

    'BijinTokeiGadget�̏I����
    Dim mIsRunBijinTokeiGadget : mIsRunBijinTokeiGadget = IsRunProgram("BijinTokeiGadget.exe")
    If mIsRunBijinTokeiGadget = True Then

        Dim mBijinTokeiGadgetExitResult : mBijinTokeiGadgetExitResult  = GetSelectedUserResultForExitProgram("BijinTokeiGadget���I�����܂����H", "BijinTokeiGadget�I����")
        If mBijinTokeiGadgetExitResult = vbYes Then

            pExitExes.Add "BijinTokeiGadget", "BijinTokeiGadget.exe"

        End If

    End If


    'BijoLinuxGadget�̏I����
    Dim mIsRunBijoLinuxGadget : mIsRunBijoLinuxGadget = IsRunProgram("BijoLinuxGadget.exe")
    If mIsRunBijoLinuxGadget = True Then

        Dim mBijoLinuxGadgetExitResult : mBijoLinuxGadgetExitResult = GetSelectedUserResultForExitProgram("BijoLinuxGadget���I�����܂����H", "BijoLinuxGadget�I����")
        If mBijoLinuxGadgetExitResult = vbYes Then

            pExitExes.Add "BijoLinuxGadget", "BijoLinuxGadget.exe"

        End If

    End If

    'Clibor�̏I����
    Dim mIsRunClibor : mIsRunClibor = IsRunProgram("Clibor.exe")
    If mIsRunClibor = True Then

        Dim mCliborExitResult : mCliborExitResult  = GetSelectedUserResultForExitProgram("Clibor���I�����܂����H", "Clibor�I����")
        If mCliborExitResult = vbYes Then

            pExitExes.Add "Clibor", "Clibor.exe"

        End If

    End If

    'GoogleDrive�̏I����
    Dim mIsRunGoogleDrive : mIsRunGoogleDrive = IsRunProgram("GoogleDriveFS.exe")
    If mIsRunGoogleDrive = True Then

        Dim mGoogleDriveExitResult : mGoogleDriveExitResult = GetSelectedUserResultForExitProgram("GoogleDrive���I�����܂����H", "GoogleDrive�I����")
        If mGoogleDriveExitResult = vbYes Then

            pExitExes.Add "GoogleDrive", "GoogleDriveFS.exe"

        End If

    End If

    'MouseGestureL�̏I���ہi�Ȃ����P��ڂ͎��s����A�Q��ڈȍ~�ɐ�������
    Dim mIsRunMouseGestureL : mIsRunMouseGestureL = IsRunProgram("MouseGestureL.exe")
    If mIsRunMouseGestureL = True Then

        Dim mMouseGestureLExitResult : mMouseGestureLExitResult  = GetSelectedUserResultForExitProgram("MouseGestureL���I�����܂����H", "MouseGestureL�I����")
        If mMouseGestureLExitResult = vbYes Then

            pExitExes.Add "MouseGestureL", "MouseGestureL.exe"

        End If

    End If

    'Orchis�̏I����
    Dim mIsRunOrchis : mIsRunOrchis = IsRunProgram("orchis.exe")
    If mIsRunOrchis = True Then

        Dim mOrchisExitResult : mOrchisExitResult  = GetSelectedUserResultForExitProgram("Orchis���I�����܂����H", "Orchis�I����")
        If mOrchisExitResult = vbYes Then

            pExitExes.Add "OrchisService", "ocobsv.exe"
            pExitExes.Add "Orchis"       , "orchis.exe"

        End If

    End If

    'Slack�̏I����
    Dim mIsRunSlack : mIsRunSlack = IsRunProgram("slack.exe")
    If mIsRunSlack = True Then

        Dim mSlackExitResult : mSlackExitResult = GetSelectedUserResultForExitProgram("Slack���I�����܂����H", "Slack�I����")
        If mSlackExitResult = vbYes Then

            pExitExes.Add "Slack", "slack.exe"

        End If

    End If

    'WheelAccele�̏I���ہi�Ȃ����P��ڂ͎��s����A�Q��ڈȍ~�ɐ�������
    Dim mIsRunWheelAccele : mIsRunWheelAccele = IsRunProgram("WheelAccele.exe")
    If mIsRunWheelAccele = True Then

        Dim mWheelAcceleExitResult : mWheelAcceleExitResult  = GetSelectedUserResultForExitProgram("WheelAccele���I�����܂����H", "WheelAccele�I����")
        If mWheelAcceleExitResult  = vbYes Then

            pExitExes.Add "WheelAccele" , "WheelAccele.exe"

        End If

    End If

    'X-Finder�̏I����
    Dim mIsRunXFinder : mIsRunXFinder = IsRunProgram("XF.exe")
    If mIsRunXFinder = True Then

        Dim mXFinderExitResult : mXFinderExitResult = GetSelectedUserResultForExitProgram("X-Finder���I�����܂����H", "X-Finder�I����")
        If mXFinderExitResult = vbYes Then

            pExitExes.Add "X-Finder32" , "XF.exe"
            pExitExes.Add "X-Finder64" , "xf64.exe"

        End If

    End If

    Set AddExitExe = pExitExes

End Function

'***********************************************************************
'* ������   �F �v���O�����I����                                      *
'* ����     �F pMsgBoxTitle  ���b�Z�[�W�{�b�N�X�̃^�C�g��              *
'*             pMsgBoxDetail ���b�Z�[�W�{�b�N�X�̓��e                  *
'* �������e �F ���b�Z�[�W�{�b�N�X��\�����[�U�[�Ƀv���O�������I������  *
'*             ���ǂ����Θb�����ʂ�Ԃ�                                *
'* �߂�l   �F ���b�Z�[�W�{�b�N�X�̌��� / vbYes�AvbNo                  *
'***********************************************************************
Function GetSelectedUserResultForExitProgram(ByVal pMsgBoxTitle,ByVal pMsgBoxDetail)

    '���[�U�[�Ƀv���O�����̏I����
    Dim mMsgBoxResult : mMsgBoxResult = MsgBox(pMsgBoxTitle, vbYesNo, pMsgBoxDetail)

    '���[�U�[���I���������ʂ��Z�b�g
    GetSelectedUserResultForExitProgram = mMsgBoxResult

End Function

'***********************************************************************
'* ������   �F �v���O�����N����Ԃ��擾                                *
'* ����     �F pProgramExe  �Ώۃv���O�����i������.exe�`���j           *
'* �������e �F �Ώۃv���O�������N�������ǂ������擾����                *
'* �߂�l   �F �Ώۃv���O�������N�����L�� / True�AFalse                *
'***********************************************************************
Function IsRunProgram(ByVal pProgramExe)

    '�N�����L���A�f�t�H���g�l�ݒ�
    IsRunProgram = False

    '�Ώۃv���O�������擾
    Set mPrograms = GetObject("winmgmts:").ExecQuery("Select * from Win32_Process where Name='" & pProgramExe & "'")

    '�Ώۃv���O�������擾�o������N�����Ƃ���i�P���ł���������j
    For Each Program in mPrograms

        IsRunProgram = True
        Exit For

    Next

End Function
