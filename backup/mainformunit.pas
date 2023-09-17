unit MainFormUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Lclintf,
  ExtDlgs, ExtCtrls, Math, Registry;

const
  AppAuthor = 'Stephanus Bagus Saputra';
  AppWebsite = 'http://www.sbskomputer.net';
  AppCopyright = 'Copryright 2020 - Stephanus Bagus Saputra';

type

  { TMainForm }

  TMainForm = class(TForm)
    BtnBrowseFolder: TButton;
    BtnRemoveFolderLib: TButton;
    BtnGenerateGuid: TButton;
    BtnBrowseIconFolder: TButton;
    BtnCreateFolderLib: TButton;
    BtnClearLogHistory: TButton;
    BtnSlideBarForm: TButton;
    LblLibPosition: TLabel;
    LblColon5: TLabel;
    LblColon6: TLabel;
    LblCopyright: TLabel;
    LblColon2: TLabel;
    LblColon3: TLabel;
    LblColon4: TLabel;
    LblStatus: TLabel;
    LblIconFolder: TLabel;
    LblInfoGuid: TLabel;
    LblColon1: TLabel;
    LblFolderLibrary: TLabel;
    DlgBrowseLibrary: TSelectDirectoryDialog;
    DlgBrowseIconFolder: TOpenPictureDialog;
    LblDescriptionFolder: TLabel;
    MeoLogging: TMemo;
    RBtnMapComputer: TRadioButton;
    RBtnMapNetwork: TRadioButton;
    RGrpMapType: TRadioGroup;
    RBtnRandomSort: TRadioButton;
    RBtnUpSort: TRadioButton;
    RBtnDownSort: TRadioButton;
    TxtNameFolder: TEdit;
    LblNameFolder: TLabel;
    TxtDescriptionFolder: TEdit;
    TxtResultGUID: TEdit;
    TxtTargetFolder: TEdit;
    TxtIconFolder: TEdit;
    procedure BtnBrowseFolderClick(Sender: TObject);
    procedure BtnBrowseIconFolderClick(Sender: TObject);
    procedure BtnClearLogHistoryClick(Sender: TObject);
    procedure BtnCreateFolderLibClick(Sender: TObject);
    procedure BtnGenerateGuidClick(Sender: TObject);
    procedure BtnRemoveFolderLibClick(Sender: TObject);
    procedure BtnSlideBarFormClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LblCopyrightClick(Sender: TObject);
  private

  public

  end;

var
  MainForm: TMainForm;
  UidGenerated: TGuid;
  UidResult: HResult;

implementation

{$R *.lfm}

{ TMainForm }

function InsertEditRegistryCLSID(const Path: String): Boolean;
var
  Registry: TRegistry;
  OpenKeyResult : Boolean;

  NameFolder: string='';
  InfoTip: string='';
  IconLibrary: string='';
  TargetFolderPath: string='';
begin
  NameFolder:= Trim(MainForm.TxtNameFolder.Text);
  InfoTip:= Trim(MainForm.TxtDescriptionFolder.Text);
  IconLibrary:= Trim(MainForm.TxtIconFolder.Text);
  TargetFolderPath:= Trim(MainForm.TxtTargetFolder.Text);

  try
    Registry:= TRegistry.Create;
    Registry.RootKey:= HKEY_LOCAL_MACHINE;
    Registry.Access := KEY_WRITE;
    OpenKeyResult := Registry.OpenKey(Path, true);

    if not OpenKeyResult then begin
      Result:= false;
    end else begin
      Registry.WriteString('', NameFolder);
      Registry.WriteExpandString('Infotip', InfoTip);
      if (MainForm.RBtnMapNetwork.Checked) then begin
        Registry.WriteInteger('DescriptionID', 9);
      end else begin
        Registry.WriteInteger('DescriptionID', 3);
      end;
      Registry.WriteInteger('System.IsPinnedToNameSpaceTree', 1);
      if (MainForm.RBtnUpSort.Checked) then begin
        Registry.WriteInteger('SortOrderIndex', 48);
      end else if (MainForm.RBtnDownSort.Checked) then begin
        Registry.WriteInteger('SortOrderIndex', 153);
      end else begin
        Registry.WriteInteger('SortOrderIndex', randomrange(0,153));
        Registry.DeleteValue('SortOrderIndex');
      end;

      OpenKeyResult := Registry.OpenKey(Path + '\DefaultIcon', true);
      if OpenKeyResult then begin
        Registry.WriteExpandString('', IconLibrary);
      end;

      OpenKeyResult := Registry.OpenKey(Path + '\InProcServer32', true);
      if OpenKeyResult then begin
        Registry.WriteExpandString('', '%systemroot%\system32\shell32.dll');
        Registry.WriteString('ThreadingModel', 'Both');
      end;

      OpenKeyResult := Registry.OpenKey(Path + '\Instance', true);
      if OpenKeyResult then begin
        Registry.WriteString('CLSID', '{0E5AAE11-A475-4c5b-AB00-C66DE400274E}');
      end;

      OpenKeyResult := Registry.OpenKey(Path + '\Instance\InitPropertyBag', true);
      if OpenKeyResult then begin
        Registry.WriteInteger('Attributes', 17);
        Registry.WriteExpandString('TargetFolderPath', TargetFolderPath);
      end;

      OpenKeyResult := Registry.OpenKey(Path + '\ShellFolder', true);
      if OpenKeyResult then begin
        Registry.WriteInteger('Attributes', 4034920525);
        Registry.WriteInteger('FolderValueFlags', 41);
        if (MainForm.RBtnUpSort.Checked) then begin
          Registry.WriteInteger('SortOrderIndex', 0);
        end else if (MainForm.RBtnDownSort.Checked) then begin
          Registry.WriteInteger('SortOrderIndex', 153);
        end else begin
          Registry.WriteInteger('SortOrderIndex', randomrange(0,153));
        end;
      end;
    end;
  finally
    Registry.Free;
  end;

  Result := true;
end;

function InsertEditRegistryNameSpace(const Path: string): Boolean;
var
  Registry: TRegistry;
  OpenKeyResult : Boolean;
begin
  try
    Registry:= TRegistry.Create;
    Registry.RootKey:= HKEY_LOCAL_MACHINE;
    Registry.Access := KEY_WRITE;
    OpenKeyResult := Registry.OpenKey(Path, true);
  finally
    Registry.Free
  end;

  Result:= OpenKeyResult;
end;

function RemoveRegistryCLSID(const Path: string; const GUIDKey: string): Boolean;
var
  Registry: TRegistry;
  OpenKeyResult: Boolean;
  DeleteSubKeyResult: Boolean;

  LastPathOpened: string='';
  CurrentKeyName: string='';
begin
  try
    Registry:= TRegistry.Create;
    Registry.RootKey:= HKEY_LOCAL_MACHINE;
    Registry.Access := KEY_ALL_ACCESS;

    LastPathOpened := Path + '\' + GUIDKey;
    OpenKeyResult := Registry.OpenKey(LastPathOpened, false);
    if OpenKeyResult then begin
      DeleteSubKeyResult:= Registry.DeleteKey('DefaultIcon');
      DeleteSubKeyResult:= Registry.DeleteKey('InProcServer32');
      CurrentKeyName := 'Instance';
      OpenKeyResult := Registry.OpenKey(LastPathOpened + '\' + CurrentKeyName, false);
      if OpenKeyResult then begin
        DeleteSubKeyResult:= Registry.DeleteKey('InitPropertyBag');

        OpenKeyResult := Registry.OpenKey(LastPathOpened, false);
        if OpenKeyResult then begin
           DeleteSubKeyResult:= Registry.DeleteKey(CurrentKeyName);
           DeleteSubKeyResult:= Registry.DeleteKey('ShellFolder');
        end;
      end;
    end;

    OpenKeyResult := Registry.OpenKey(Path, false);
    if OpenKeyResult then begin
       DeleteSubKeyResult:= Registry.DeleteKey(GUIDKey);
    end;

    Registry.CloseKey;
    FreeAndNil(Registry);
  except
    on E: Exception do begin
       ShowMessage(E.Message);
    end;
  end;

  Result:= DeleteSubKeyResult;
end;

function RemoveRegistryNamespace(const Path: string; const GUIDKey: string): Boolean;
var
  Registry: TRegistry;
  OpenKeyResult: Boolean;
  DeleteSubKeyResult: Boolean;
begin
  try
    Registry:= TRegistry.Create;
    Registry.RootKey:= HKEY_LOCAL_MACHINE;
    Registry.Access := KEY_ALL_ACCESS;

    OpenKeyResult := Registry.OpenKey(Path, false);
    if OpenKeyResult then begin
       DeleteSubKeyResult:= Registry.DeleteKey(GUIDKey);
    end;

    Registry.CloseKey;
    FreeAndNil(Registry);
  except
    on E: Exception do begin
       ShowMessage(E.Message);
    end;
  end;

  Result:= DeleteSubKeyResult;
end;

function CreateFolderLibrary(const GUIDKey: string): Boolean;
var
  CLSIDResult: Boolean;
  CLSID64Result: Boolean;
  NamespaceResult: Boolean;
  Namespace64Result: Boolean;
begin
  CLSIDResult:= InsertEditRegistryCLSID('\SOFTWARE\Classes\CLSID\' + GUIDKey);
  CLSID64Result:= InsertEditRegistryCLSID('\SOFTWARE\WOW6432Node\Classes\CLSID\' + GUIDKey);
  NamespaceResult:= InsertEditRegistryNameSpace('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\'+ GUIDKey);
  Namespace64Result:= InsertEditRegistryNameSpace('\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\'+ GUIDKey);

  if (CLSIDResult and CLSID64Result) and (NamespaceResult and Namespace64Result) then begin
    Result:= true;
  end else begin
    Result:= false;
  end;
end;

function RemoveFolderLibrary(const GUIDKey: string): Boolean;
var
  CLSIDResult: Boolean;
  CLSID64Result: Boolean;
  NamespaceResult: Boolean;
  Namespace64Result: Boolean;
begin
  CLSIDResult:= RemoveRegistryCLSID('\SOFTWARE\Classes\CLSID', GUIDKey);
  CLSID64Result:= RemoveRegistryCLSID('\SOFTWARE\WOW6432Node\Classes\CLSID', GUIDKey);
  NamespaceResult:= RemoveRegistryNameSpace('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace', GUIDKey);
  Namespace64Result:= RemoveRegistryNameSpace('\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace', GUIDKey);

  if (CLSIDResult and CLSID64Result) and (NamespaceResult and Namespace64Result) then begin
    Result:= true;
  end else begin
    Result:= false;
  end;
end;

function CreateMicrosoftLibrary(const NameLib: string; const Path: string; const Icon: string): Boolean;
var
  LogWritter: TStringList;
  MSLibFile: string='';
begin
  LogWritter:= TStringList.Create;
  MSLibFile:= GetEnvironmentVariable('APPDATA') + '\Microsoft\Windows\Libraries\' + NameLib + '.library-ms';

  LogWritter.Add('<?xml version="1.0" encoding="UTF-8"?>');
  LogWritter.Add('<libraryDescription xmlns="http://schemas.microsoft.com/windows/2009/library">');
  LogWritter.Add('<isLibraryPinned>true</isLibraryPinned>');
  LogWritter.Add('<iconReference>' + Icon + '</iconReference>');
  LogWritter.Add('<templateInfo>');
  LogWritter.Add('<folderType>{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}</folderType>');
  LogWritter.Add('</templateInfo>');
  LogWritter.Add('<propertyStore>');
  LogWritter.Add('<property name="HasModifiedLocations" type="boolean"><![CDATA[true]]></property>');
  LogWritter.Add('</propertyStore>');
  LogWritter.Add('<searchConnectorDescriptionList>');
  LogWritter.Add('<searchConnectorDescription>');
  LogWritter.Add('<isDefaultSaveLocation>true</isDefaultSaveLocation>');
  LogWritter.Add('<isDefaultNonOwnerSaveLocation>true</isDefaultNonOwnerSaveLocation>');
  LogWritter.Add('<isSupported>true</isSupported>');
  LogWritter.Add('<simpleLocation>');
  LogWritter.Add('<url>' + Path + '</url>');
  LogWritter.Add('</simpleLocation>');
  LogWritter.Add('</searchConnectorDescription>');
  LogWritter.Add('</searchConnectorDescriptionList>');
  LogWritter.Add('</libraryDescription>');

  LogWritter.SaveToFile(MSLibFile);

  Result:= FileExists(MSLibFile);
end;

function RemoveMicrosoftLibrary(const NameLib: string): Boolean;
var
  MSLibFile: string='';
begin
  MSLibFile:= GetEnvironmentVariable('APPDATA') + '\Microsoft\Windows\Libraries\' + NameLib + '.library-ms';
  Result:= DeleteFile(MSLibFile);
end;

procedure LoadLogHistory();
var
  LogWritter: TStringList;
  LogFile: string='log.txt';
  i: integer=0;
begin
  MainForm.MeoLogging.Lines.Clear;

  try
    LogWritter:= TStringList.Create;

    if FileExists(LogFile) then begin
       LogWritter.LoadFromFile(LogFile);
       for i:=0 to LogWritter.Count-1 do begin
         MainForm.MeoLogging.Lines.Add(LogWritter[i]);
       end;
    end;
  finally
    LogWritter.Free;
  end;
end;

procedure CreateLogHistory(const msg: string);
var
  LogWritter: TStringList;
  LogFile: string='log.txt';
begin
  try
    LogWritter:= TStringList.Create;

    if FileExists(LogFile) then begin
       LogWritter.LoadFromFile(LogFile);
    end;

    LogWritter.Add(msg);
    LogWritter.SaveToFile(LogFile);
  finally
    LogWritter.Free;
  end;
end;

procedure ClearLogHistory();
var
  LogWritter: TStringList;
  LogFile: string='log.txt';
begin
  try
    LogWritter:= TStringList.Create;
    LogWritter.Add('');
    LogWritter.SaveToFile(LogFile);
  finally
    LogWritter.Free;
  end;

end;

procedure TMainForm.FormShow(Sender: TObject);
begin
  MainForm.LblCopyright.Caption:= AppCopyright;
  MainForm.BtnGenerateGuid.Click();
  LoadLogHistory();
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
   MainForm.LblCopyright.Caption:= AppCopyright;
   MainForm.BorderIcons:= [biSystemMenu,biMinimize];
   MainForm.MeoLogging.ReadOnly:= true;
   MainForm.MeoLogging.Lines.Clear;
end;

procedure TMainForm.LblCopyrightClick(Sender: TObject);
begin
  OpenURL(AppWebsite);
end;

procedure TMainForm.BtnBrowseFolderClick(Sender: TObject);
begin
  if DlgBrowseLibrary.Execute then begin
    TxtTargetFolder.Text := Trim(DlgBrowseLibrary.FileName);
    MainForm.TxtNameFolder.Text:= ExtractFileName(Trim(MainForm.TxtTargetFolder.Text));
  end;
end;

procedure TMainForm.BtnBrowseIconFolderClick(Sender: TObject);
begin
  if DlgBrowseIconFolder.Execute then begin
    TxtIconFolder.Text := DlgBrowseIconFolder.FileName;
  end;
end;

procedure TMainForm.BtnClearLogHistoryClick(Sender: TObject);
begin
  ClearLogHistory();
  LoadLogHistory();
end;

procedure TMainForm.BtnGenerateGuidClick(Sender: TObject);
begin
  UidResult := CreateGUID(UidGenerated);
  if UidResult = S_OK then begin
    TxtResultGUID.Text := GuidToString(UidGenerated);
  end;
end;

procedure TMainForm.BtnRemoveFolderLibClick(Sender: TObject);
var
  GuidVal: string='';
  NameLibrary: string='';
  MsgResult: string='';
begin
  GuidVal := Trim(MainForm.TxtResultGUID.Text);
  NameLibrary := Trim(MainForm.TxtNameFolder.Text);

  if (GuidVal = '') then begin
    ShowMessage('GUID Cannot be empty!!');
  end else if (NameLibrary = '') then begin
    ShowMessage('Name of Folder Cannot be empty!!');
  end else begin
      if RemoveFolderLibrary(GuidVal) then begin
        CreateLogHistory('Removed Folder GUID: ' + GuidVal + '; Name: ' + NameLibrary);
        RemoveMicrosoftLibrary(NameLibrary);
        MsgResult:= 'Folder has been Removed!!';
        MainForm.LblStatus.Font.Color:= clGreen;
      end else begin
        MsgResult:= 'Failed to removed folder!!';
        MainForm.LblStatus.Font.Color:= clRed;
      end;
      MainForm.LblStatus.Caption := MsgResult;
      ShowMessage(MsgResult);
      LoadLogHistory();
  end;
end;

procedure TMainForm.BtnSlideBarFormClick(Sender: TObject);
begin
  if MainForm.Width <= 460 then begin
    MainForm.Width:= 643;
    MainForm.BtnSlideBarForm.Caption := '<';
  end else begin
    MainForm.Width:= 455;
    MainForm.BtnSlideBarForm.Caption := '>';
  end;
end;

procedure TMainForm.BtnCreateFolderLibClick(Sender: TObject);
var
  GuidVal: string='';
  NameFolder: string='';
  InfoTip: string='';
  IconLibrary: string='';
  TargetFolderPath: string='';

  MsgResult: string='';
begin
  GuidVal:= Trim(MainForm.TxtResultGUID.Text);
  NameFolder:= Trim(MainForm.TxtNameFolder.Text);
  InfoTip:= Trim(MainForm.TxtDescriptionFolder.Text);
  IconLibrary:= Trim(MainForm.TxtIconFolder.Text);
  TargetFolderPath:= Trim(MainForm.TxtTargetFolder.Text);

  if (GuidVal = '') then begin
    ShowMessage('GUID Cannot be empty!!');
  end else if (NameFolder = '') then begin
    ShowMessage('Name of Folder Cannot be empty!!');
  end else if (TargetFolderPath = '') then begin
    ShowMessage('Target Folder Library Cannot be empty!!');
  end else if (IconLibrary = '') then begin
    ShowMessage('Icon Folder Library Cannot be empty!!');
  end else begin
    if (InfoTip = '') then begin
      MainForm.TxtDescriptionFolder.Text:= NameFolder;
      InfoTip:= Trim(MainForm.TxtDescriptionFolder.Text);
    end;

    if CreateFolderLibrary(GuidVal) then begin
      if (MainForm.RBtnMapNetwork.Checked) then begin
        CreateLogHistory('Created/Modifed Folder on Map Network by GUID: ' + GuidVal + '; Name: ' + NameFolder + '; Path: ' + TargetFolderPath);
      end else begin
        CreateMicrosoftLibrary(NameFolder, TargetFolderPath, IconLibrary);
        CreateLogHistory('Created/Modifed Folder by GUID: ' + GuidVal + '; Name: ' + NameFolder + '; Path: ' + TargetFolderPath);
      end;
      MsgResult:= 'Folder Added Successfull!!';
      MainForm.LblStatus.Font.Color:= clGreen;
    end else begin
      MsgResult:= 'Please run program as Administrator!!';
      MainForm.LblStatus.Font.Color:= clRed;
    end;
    MainForm.LblStatus.Caption := MsgResult;
    ShowMessage(MsgResult);
    LoadLogHistory();
  end;
end;

end.

