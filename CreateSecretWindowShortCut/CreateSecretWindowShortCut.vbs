'**************************************************************************************
'* �v���O������ �F �V�[�N���b�g�E�B���h�E�iChrome�j�Ŏ��s����V���[�g�J�b�g���쐬     *
'* �����T�v     �F Chrome�œ����URL���V�[�N���b�g�E�B���h�E�Ŏ��s����V���[�g�J�b�g  *
'*                 ���쐬����                                                         *
'*                 ��Chrome�̃p�X�̓��W�X�g������擾���邽�߁A���ɂ���Ă�Chrome�� *
'*                   �p�X���Ⴄ�\�������邽�ߒ��ӂ��邱��                           *
'* ����         �F ����̃u���E�U��Chrome�ł��邱�ƑO��ō쐬���Ă���̂Ŋ���̃u���E *
'*                 �U��Chrome�Ŗ����ꍇ�͐������������̂��쐬����܂�                 *
'* �ݒ�         �F                                                                    *
'**************************************************************************************

Option Explicit

'*****************************************
'* �ϐ�                                  *
'*****************************************
Dim mObjShell    : Set mObjShell  = WScript.CreateObject("WScript.Shell")
Dim mFileInfo    : Set mFileInfo  = WScript.CreateObject("Scripting.Dictionary")

'*****************************************
'* �萔                                  *
'*****************************************
'MsgBox�InputBox�ɕ\�����镶����ݒ�
Dim cMsgTitle                  : cMsgTitle                  = "�V�[�N���b�g�E�B���h�E�Ŏ��s�ulnk�v�t�@�C���쐬"
Dim cMsgInputFileName          : cMsgInputFileName          = "�t�@�C��������͂��Ă��������B" & VbCrLf & "���t�@�C�����́u������.lnk�v�������̕������w�肵�Ă��������B�t�@�C���̓f�X�N�g�b�v�ɍ쐬����܂��B"
Dim cMsgInputURL               : cMsgInputURL               = "URL����͂��Ă��������B"
Dim cMsgIncorrectFileNameError : cMsgIncorrectFileNameError = "�t�@�C�����������͂܂��̓t�@�C�����Ƃ��Đ���������܂���" & VbCrLf & "�������I�����܂�..."
Dim cMsgIncorrectURLError      : cMsgIncorrectURLError      = "URL�������͂܂���URL�Ƃ��Đ���������܂���" & VbCrLf & "�������I�����܂�..."

'�t�@�C���̏o�͐�p�X
Dim cDesktopPath : cDesktopPath = mObjShell.SpecialFolders("Desktop") & "\"

'����̃u���E�U�̃p�X ������̃u���E�U��Chrome�ł��邱��
Dim cDefaultBrowserExePath : cDefaultBrowserExePath = GetDefaultBrowserPath()


Main()

'***********************************************************************
'* ������   �F ���C������                                              *
'* ����     �F �Ȃ�                                                    *
'* �������e �F ���C������                                              *
'* �߂�l   �F �Ȃ�                                                    *
'***********************************************************************
Sub Main()

    '------------------------------------
    ' �t�@�C�������擾
    '------------------------------------
    '���[�U�[�ɑΘb���t�@�C�������擾
    Dim mFileName : mFileName = InputBox(cMsgInputFileName, cMsgTitle)

    '�t�@�C���������������ǂ������擾
    Dim mIsCorrectFileName : mIsCorrectFileName = IsCorrectRegExpMatch("(\\|\/|\:|\*|\?|\""|\<|\>|\|)", mFileName)

    '�����͂܂��̓t�@�C�������������Ȃ��ꍇ�͏������I��
    If mIsCorrectFileName  = True Or mFileName = "" Then

        MsgBox cMsgIncorrectFileNameError, vbOKOnly, cMsgTitle
        Wscript.Quit()

    End If

    '�t�@�C�����̖����Ɂu.lnk�v��ǉ�
    mFileName = mFileName & ".lnk"

    '------------------------------------
    ' URL���擾
    '------------------------------------
    '���[�U�[�ɑΘb�iURL�j
    Dim mURL : mURL = InputBox(cMsgInputURL, cMsgTitle)

    'URL�����������ǂ������擾
    Dim mIsCorrectURL : mIsCorrectURL = IsCorrectRegExpMatch("^(https*|ftp)://[-_!~';:@&=,%#/a-zA-Z0-9\$\*\+\?\.\(\)]+$", mURL)

    '�����͂܂���URL���������Ȃ��ꍇ�͏������I��
    If mIsCorrectURL = False Or mURL = "" Then

        MsgBox cMsgIncorrectURLError, vbOKOnly, cMsgTitle
        Wscript.Quit()

    End If

    '------------------------------------
    ' �V�[�N���b�g�E�B���h�E�p�̃R�}���h���C���������쐬
    '------------------------------------
    '--incognito + "�\��������URL"
    Dim mCommandLineArg : mCommandLineArg = "--incognito " & """" & mURL & """"

    '------------------------------------
    ' �V���[�g�J�b�g�쐬����
    '------------------------------------
    '�t�@�C������ݒ�('�L�[�F�t�@�C�����A�쐬���F�t�@�C���p�X|�o�͐�t�H���_)
                  '�t�@�C����  '�V���[�g�J�b�g��               '�t�@�C���̏o�͐�         '�R�}���h���C������    '�A�C�R��
    mFileInfo.Add mFileName ,   cDefaultBrowserExePath & "|" &  cDesktopPath      & "|" & mCommandLineArg & "|" & cDefaultBrowserExePath

    '�V���[�g�J�b�g�쐬����
    Dim key
    For Each key In mFileInfo.Keys

        '�쐬����؂蕪����i�@�V���[�g�J�b�g�p�X�A�A�o�̓t�H���_�A�B�R�}���h���C�������A�C�A�C�R���j
        Dim mAryFileInfo : mAryFileInfo = Split(mFileInfo(key),"|")

        'lnk�t�@�C���̃t�@�C��������V���[�g�J�b�g���쐬����t���p�X���擾
        Dim mShortCutFileName     : mShortCutFileName     = key
        Dim mShortCutFileFullPath : mShortCutFileFullPath = mAryFileInfo(1) & mShortCutFileName

        '�V���[�g�J�b�g���쐬�t�@�C���̕ۑ���t�H���_�����������ꍇ�̓t�H���_���쐬����
        CreateNotExistFolder(mShortCutFileFullPath)

        '�V���[�g�J�b�g�I�u�W�F�N�g���쐬���o�͐�p�X�A�R�}���h���C�������A�A�C�R�����w��
        Dim mShortCut : Set mShortCut = mObjShell.CreateShortcut(mShortCutFileFullPath) '�V���[�g�J�b�g�I�u�W�F�N�g���쐬
        mShortCut.TargetPath                                    = mAryFileInfo(0)       '�V���[�g�J�b�g��       ��Chrome�̃p�X���Z�b�g
        If UBound(mAryFileInfo) > 1 Then mShortCut.Arguments    = mAryFileInfo(2)       '�R�}���h���C�������ݒ�
        If UBound(mAryFileInfo) > 2 Then mShortCut.IconLocation = mAryFileInfo(3)       '�A�C�R������ݒ�
    
        '�V���[�g�J�b�g���쐬
        mShortCut.Save
    
    Next

    '�I�u�W�F�N�g�̔j��
    Set mObjShell  = Nothing
    Set mFileInfo  = Nothing

End Sub

'***********************************************************************
'* ������   �F ����̃u���E�U�̃p�X���擾                              *
'* ����     �F �Ȃ�                                                    *
'* �������e �F ���W�X�g���������̃u���E�U�̃p�X��������擾���Ԃ�    *
'* �߂�l   �F ����̃u���E�U�̃p�X                                    *
'***********************************************************************
Function GetDefaultBrowserPath()

    '����̃u���E�U�ŊJ��Exe�̃p�X�����W�X�g������T���L�[
    '��OS�ɂ���Ă��Ⴄ���߃��W�X�g���G�f�B�^�Łuchrome.exe�v�Ō�������Ƃ�����ۂ����̂�������
    Dim cRegRunHttpKey : cRegRunHttpKey = "HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command\"

    '�uWScript.Shell�v�̃I�u�W�F�N�g���쐬
    Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell")

    '���W�X�g���ɐݒ肳��Ă���l���擾����
    '���u"C:\Program Files\Google\Chrome\Application\chrome.exe"  --single-argument %1�v�̌`���Ŏ擾�����
    Dim mDefaultBrowserValue   : mDefaultBrowserValue = mObjShell.RegRead(cRegRunHttpKey)

    '���W�X�g���ɐݒ肳��Ă���l����Exe�̃p�X���擾����
    '���u"�u���E�U�p�X" -- "%1"�v�̌`������u"�u���E�U�p�X"�v�݂̂𔲂��o��
    Dim mDefaultBrowserExePath : mDefaultBrowserExePath = Left(mDefaultBrowserValue, InStr(mDefaultBrowserValue, ".exe") + 4)

    '�쐬�����I�u�W�F�N�g��j��
    Set mObjShell = Nothing

    '�Ԃ�l��ݒ�i�_�u���N�I�[�e�[�V�������폜�j
    GetDefaultBrowserPath = Replace(mDefaultBrowserExePath, """", "")

End Function

'***********************************************************************
'* ������   �F ���K�\���Ɉ�v���邩                                    *
'* ����     �F pPattern ���K�\���p�^�[��                               *
'*             pString  �Ώە�����                                     *
'* �������e �F �p�^�[���Ɉ�v�i���K�\���Ń`�F�b�N�j���邩�ǂ���        *
'* �߂�l   �F �p�^�[���Ɉ�v�FTrue�A�p�^�[���ɕs��v�FFalse           *
'***********************************************************************
Function IsCorrectRegExpMatch(pPattern, pString)

    '----------------------------------
    ' ���K�\���I�u�W�F�N�g���쐬
    '----------------------------------
    Dim mRegExp : Set mRegExp = New RegExp
    mRegExp.Pattern    = pPattern '���K�\���̃p�^�[����ݒ�
    mRegExp.IgnoreCase = True     '�啶���E����������ʂ��Ȃ��悤�ɐݒ�
    mRegExp.Global     = True     '�������S�̂���������悤�ɐݒ�

    '----------------------------------
    ' ���K�\���Ɉ�v���邩���ʂ��擾
    '----------------------------------
    '���K�\���p�^�[���Ɉ�v������
    If mRegExp.test(pString) Then ' �������e�X�g���܂��B

        IsCorrectRegExpMatch = True

    Else

        IsCorrectRegExpMatch = False

    End If

    '----------------------------------
    ' �쐬�����I�u�W�F�N�g��j��
    '----------------------------------
    Set mRegExp = Nothing

End Function

'***********************************************************************
'* ������   �F �t�H���_�쐬����                                        *
'* ����     �F pPath �Ώۃp�X                                          *
'* �������e �F �Ώۃp�X�ɑ��݂��Ȃ��p�X��������쐬����                *
'*             ���t�H���_�̍쐬�͍ċA�I�ɍs���܂�                      *
'*               �h���C�u �� �h���C�u\�K�w�P���h���C�u\�K�w�P\�K�w�Q\  *
'*               �� �h���C�u\�K�w�P\�K�w�Q\�Ώۃt�H���_                *
'* �߂�l   �F �Ȃ�                                                    *
'***********************************************************************
Function CreateNotExistFolder(pPath)

    Dim mObjFso       : Set mObjFso   = WScript.CreateObject("Scripting.FileSystemObject")
    Dim mDriveName    : mDriveName    = Left(mObjFso.GetDriveName(pPath),2)                '�h���C�u�����擾
    Dim mParentFolder : mParentFolder = mObjFso.GetParentFolderName(pPath)                 '�e�t�H���_�[�����擾

    '�Ώۂ̃h���C�u�����݂��鎞
    If mObjFso.DriveExists(mDriveName) Then

        'Drive�I�u�W�F�N�g���쐬
        Dim mObjDrive : Set mObjDrive = mObjFso.GetDrive(mDriveName) 

    Else

        Exit Function

    End If

    '�h���C�u�̏������ł��Ă��鎞
    If mObjDrive.IsReady Then

        '�g���q�����񂪎擾�o�����ꍇ(�t�@�C���̎�)
        If Len(mObjFso.GetExtensionName(pPath)) > 0 Then 

            '�e�t�H���_�[�����݂��Ȃ����A�Ώۃp�X����e�t�H���_�[�쐬����i�ċA�I�j
            If Not(mObjFso.FolderExists(mParentFolder)) Then CreateNotExistFolder(mParentFolder)

        Else

            '�Ώۃt�H���_�[�����݂��Ȃ���
            If Not(mObjFso.FolderExists(pPath)) Then

                '�e�t�H���_�[���쐬��A�Ώۃt�H���_�[���쐬�i�ċA�I�j
                CreateNotExistFolder(mParentFolder)
                mObjFso.CreateFolder(pPath)

            End If

        End If

    End If

    '�I�u�W�F�N�g�̔j��
    Set mObjFso   = Nothing
    Set mObjDrive = Nothing

End Function
