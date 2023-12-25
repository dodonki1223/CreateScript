'--------------------------------------
' �萔�錾
'--------------------------------------
Const AUTO_HOT_KEY_TOOL_TOOLS_DIRECTORY = "\Tools\AutoHotKey\Tools\"

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
        Dim symLinkPath       : symLinkPath = runDrive & AUTO_HOT_KEY_TOOL_TOOLS_DIRECTORY & arySymLinks(0)
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
    'Tools �t�H���_���̃V���{���b�N�����N�̍쐬
    pSymLinks.Add "Explorers"         , "Explorers"               & "|" & pRunDrive & "\Tools\Explorers"
    pSymLinks.Add "FolderFileList"    , "FolderFileList"          & "|" & pRunDrive & "\Tools\FolderFileList"
    pSymLinks.Add "ImageForClipboard" , "ImageForClipboard"       & "|" & pRunDrive & "\Tools\ImageForClipboard"
    pSymLinks.Add "MyAutoHotKeySpy"    , "MyAutoHotKeySpy"        & "|" & pRunDrive & "\Tools\MyAutoHotKeySpy"
    pSymLinks.Add "ShutDownDialog"    , "ShutDownDialog"          & "|" & pRunDrive & "\Tools\ShutDownDialog"

    'SelfMadeMenu�t�H���_���̃V���{���b�N�����N�̍쐬
    pSymLinks.Add "Components"        , "SelfMadeMenu\Components" & "|" & pRunDrive & "\Tools\AutoHotKey\Components"
    pSymLinks.Add "Icon"              , "SelfMadeMenu\Icon"       & "|" & pRunDrive & "\Tools\AutoHotKey\Icon"
    pSymLinks.Add "Tools"             , "SelfMadeMenu\Tools"      & "|" & pRunDrive & "\Tools\AutoHotKey\Tools"

End Function
