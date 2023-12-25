'--------------------------------------
' �萔�錾
'--------------------------------------
Const PORTABLE_APPS_DIRECTORY = "\Tools\PortableApps\PortableApps\"

'--------------------------------------
' �ϐ��錾�E�C���X�^���X�쐬
'--------------------------------------
Dim objAppli : Set objAppli   = WScript.CreateObject("Shell.Application")          'WScript.Application�I�u�W�F�N�g
Dim objFso   : set objFso     = WScript.CreateObject("Scripting.FileSystemObject") 'FileSystemObject
Dim objShell : Set objShell   = WScript.CreateObject("WScript.Shell")              'WScript.Shell�I�u�W�F�N�g
Dim symLinks : Set symLinks   = WScript.CreateObject("Scripting.Dictionary")       '�V���{���b�N�����N�i�[Dictionary

'--------------------------------------
' �Ǘ��Ҍ����Ŏ��s������
'--------------------------------------
' 2��ڈȍ~�� runas �Ƃ����R�}���h���C��������n���Ď��s����
if Wscript.Arguments.Count = 0 then
    objAppli.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
    Wscript.Quit
end if

Main()

'***********************************************************************
'* ������   �F ���C������                                              *
'* ����     �F �Ȃ�                                                    *
'* �������e �F ���C������                                              *
'* �߂�l   �F �Ȃ�                                                    *
'***********************************************************************
Sub Main()

    '--------------------------------------
    ' ���s�h���C�u���擾
    '--------------------------------------
    Dim runDrive : runDrive = objFSo.GetDriveName(WScript.ScriptFullName)

    '--------------------------------------
    ' �V���{���b�N�����N�쐬����
    '--------------------------------------
    Call AddSymLinks(symLinks ,runDrive)

    For Each key In symLinks.Keys

        '�쐬����؂蕪����i�@PortableApps�Ǘ����̃t�H���_���A�A�C���X�g�[����t�H���_�j
        Dim arySymLinks : arySymLinks = Split(symLinks(key), "|")

        'SymLink�̍쐬��ƃ����N��̃f�B���N�g�����擾����
        Dim symLinkPath       : symLinkPath = runDrive & PORTABLE_APPS_DIRECTORY & arySymLinks(0)
        Dim symLinkTargetPath : symLinkTargetPath = arySymLinks(1)

        '�쐬���SymLink�����ɂ���ƃG���[�ɂȂ邽�ߎ��O�ɍ폜���ċ����I�ɍč쐬������
        if objFso.FolderExists(symLinkPath) then
            objShell.Run "cmd /c rmdir " & symLinkPath, 0, false
        end if

        '�V���{���b�N�����N�쐬�̃R�}���h�����s���Ă���
        'USB�Ȃǂ̃f�t�H���g�̃t�@�C���V�X�e����Fat�n���ƃV���{���b�N�����N�̍쐬���ł��Ȃ�����NTFS�ɂ��炩���߃t�H�[�}�b�g����K�v������
        objShell.Run "cmd /c mklink /d " & symLinkPath &  " " & symLinkTargetPath, 0, false

    Next

    '--------------------------------------
    ' �I�u�W�F�N�g�j������
    '--------------------------------------
    Set objShell = Nothing
    Set objAppli = Nothing
    Set objFso   = Nothing

End Sub


'***********************************************************************
'* ������   �F �V���{���b�N�����N�쐬�p�̃f�B���N�g���ǉ�����          *
'* ����     �F pSymLinks        �쐬�f�B���N�g���i�[Dictionary         *
'*             pRunDrive        ���s�h���C�u�p�X                       *
'* �������e �F �V���{���b�N�����N�쐬�p�̃f�B���N�g������Dictionary��  *
'*             �ǉ�����                                                *
'* �߂�l   �F pRunSymLinks                                            *
'***********************************************************************
Function AddSymLinks(ByRef pSymLinks,ByVal pRunDrive)

    '--------------------------------------
    ' �V���{���b�N�����N��ݒ肵�Ă���
    ' ���L�[�F�A�v�����A���ځF�A�v���p�X
    '--------------------------------------
    pSymLinks.Add "7-ZipPortable"             , "7-ZipPortable"             & "|" & pRunDrive & "\Tools\7-ZipPortable"
    pSymLinks.Add "CDExPortable"              , "CDExPortable"              & "|" & pRunDrive & "\Tools\CDExPortable"
    pSymLinks.Add "CPU-ZPortable"             , "CPU-ZPortable"             & "|" & pRunDrive & "\Tools\CPU-ZPortable"
    pSymLinks.Add "CrystalDiskInfoPortable"   , "CrystalDiskInfoPortable"   & "|" & pRunDrive & "\Tools\CrystalDiskInfoPortable"
    pSymLinks.Add "CrystalDiskMarkPortable"   , "CrystalDiskMarkPortable"   & "|" & pRunDrive & "\Tools\CrystalDiskMarkPortable"
    pSymLinks.Add "FastCopyPortable"          , "FastCopyPortable"          & "|" & pRunDrive & "\Tools\FastCopyPortable"
    pSymLinks.Add "GIMPPortable"              , "GIMPPortable"              & "|" & pRunDrive & "\Tools\GIMPPortable"
    pSymLinks.Add "GoogleChromePortable"      , "GoogleChromePortable"      & "|" & pRunDrive & "\Tools\GoogleChromePortable"
    pSymLinks.Add "GPU-ZPortable"             , "GPU-ZPortable"             & "|" & pRunDrive & "\Tools\GPU-ZPortable"
    pSymLinks.Add "IObitUninstallerPortable"  , "IObitUninstallerPortable"  & "|" & pRunDrive & "\Tools\IObitUninstallerPortable"
    pSymLinks.Add "IObitUnlockerPortable"     , "IObitUnlockerPortable"     & "|" & pRunDrive & "\Tools\IObitUnlockerPortable"
    pSymLinks.Add "PDFTKBuilderPortable"      , "PDFTKBuilderPortable"      & "|" & pRunDrive & "\Tools\PDFTKBuilderPortable"
    pSymLinks.Add "PDF-XChangeViewerPortable" , "PDF-XChangeViewerPortable" & "|" & pRunDrive & "\Tools\PDF-XChangeViewerPortable"
    pSymLinks.Add "ProcessExplorerPortable"   , "ProcessExplorerPortable"   & "|" & pRunDrive & "\Tools\ProcessExplorerPortable"
    pSymLinks.Add "ProcessMonitorPortable"    , "ProcessMonitorPortable"    & "|" & pRunDrive & "\Tools\ProcessMonitorPortable"
    pSymLinks.Add "SystemExplorerPortable"    , "SystemExplorerPortable"    & "|" & pRunDrive & "\Tools\SystemExplorerPortable"
    pSymLinks.Add "TeamViewerPortable"        , "TeamViewerPortable"        & "|" & pRunDrive & "\Tools\TeamViewerPortable"
    pSymLinks.Add "VLCPortable"               , "VLCPortable"               & "|" & pRunDrive & "\Tools\VLCPortable"
    pSymLinks.Add "WinMergePortable"          , "WinMergePortable"          & "|" & pRunDrive & "\Tools\WinMergePortable"
    pSymLinks.Add "wxMP3gainPortable"         , "wxMP3gainPortable"         & "|" & pRunDrive & "\Tools\wxMP3gainPortable"
    pSymLinks.Add "XnViewPortable"            , "XnViewPortable"            & "|" & pRunDrive & "\Tools\XnViewPortable"

End Function
