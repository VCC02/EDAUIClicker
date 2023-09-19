{
    Copyright (C) 2023 VCC
    creation date: Nov 2022
    initial release date: 24 Nov 2022

    author: VCC
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the
    Software is furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
}


unit EDAUIClickerMainForm;

{$H+}
{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

interface

uses
  Windows, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, Buttons, ExtDlgs, StdCtrls, ClickerIniFiles;

type

  { TfrmEDAUIClickerMain }

  TfrmEDAUIClickerMain = class(TForm)
    bitbtnShowEDAProjectsForm: TBitBtn;
    bitbtnShowRemoteScreenShotForm: TBitBtn;
    bitbtnShowTemplateCallTree: TBitBtn;
    imgAppIcon: TImage;
    lblBitness: TLabel;
    OpenDialog1: TOpenDialog;
    OpenDialog2: TOpenDialog;
    OpenPictureDialog1: TOpenPictureDialog;
    SaveDialog1: TSaveDialog;
    SaveDialog2: TSaveDialog;
    tmrStartup: TTimer;
    procedure btnShowEDAProjectsFormClick(Sender: TObject);
    procedure btnDisplayTemplateCallTreeClick(Sender: TObject);
    procedure btnShowRemoteScreenShotFormClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure rdgrpFileSystemClick(Sender: TObject);
    procedure tmrStartupTimer(Sender: TObject);
  private
    { Private declarations }
    FUsingMultiSelect: Boolean;

    procedure LoadSettings;
    procedure SaveSettings;

    function HandleOnGetConnectionAddress: string;
    function HandleOnGetSelectedCompFromRemoteWin: THandle;

    procedure HandleOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);

    function HandleOnFileExists(const FileName: string): Boolean;
    function HandleOnTClkIniReadonlyFileCreate(AFileName: string): TClkIniReadonlyFile;
    procedure HandleOnSetTemplateOpenDialogInitialDir(AInitialDir: string);
    function HandleOnOpenDialogExecute(AFilter: string): Boolean;
    function HandleOnGetOpenDialogFileName: string;
    procedure HandleOnSetTemplateSaveDialogInitialDir(AInitialDir: string);
    function HandleOnTemplateSaveDialogExecute: Boolean;
    function HandleOnGetTemplateSaveDialogFileName: string;

    procedure HandleOnSetOpenDialogMultiSelect;

    procedure HandleOnSetEDAClickerFileOpenDialogInitialDir(AInitialDir: string);
    function HandleOnEDAClickerFileOpenDialogExecute: Boolean;
    function HandleOnGetEDAClickerFileOpenDialogFileName: string;
    procedure HandleOnSetEDAClickerFileSaveDialogInitialDir(AInitialDir: string);
    function HandleOnEDAClickerFileSaveDialogExecute: Boolean;
    function HandleOnGetEDAClickerFileSaveDialogFileName: string;
    procedure HandleOnGetListOfFilesFromDir(ADir, AFileExtension: string; AListOfFiles: TStringList);
    procedure HandleOnSaveTextToFile(AStringList: TStringList; const AFileName: string);
    procedure HandleOnSetStopAllActionsOnDemand(Value: Boolean);
  public
    { Public declarations }
    procedure SetHandles;
  end;

var
  frmEDAUIClickerMain: TfrmEDAUIClickerMain;


implementation

{$IFnDEF FPC}
  {$R *.dfm}
{$ELSE}
  {$R *.frm}
{$ENDIF}


uses
  EDAProjectsClickerForm, ClickerTemplateCallTreeForm, ClickerRemoteScreenForm,
  ClickerActionsClient, IniFiles;


procedure TfrmEDAUIClickerMain.SetHandles;
begin
  frmEDAProjectsClickerForm.OnFileExists := HandleOnFileExists;
  frmEDAProjectsClickerForm.OnTClkIniReadonlyFileCreate := HandleOnTClkIniReadonlyFileCreate;
  frmEDAProjectsClickerForm.OnSetEDAClickerFileOpenDialogInitialDir := HandleOnSetEDAClickerFileOpenDialogInitialDir;
  frmEDAProjectsClickerForm.OnEDAClickerFileOpenDialogExecute := HandleOnEDAClickerFileOpenDialogExecute;
  frmEDAProjectsClickerForm.OnGetEDAClickerFileOpenDialogFileName := HandleOnGetEDAClickerFileOpenDialogFileName;
  frmEDAProjectsClickerForm.OnSetEDAClickerFileSaveDialogInitialDir := HandleOnSetEDAClickerFileSaveDialogInitialDir;
  frmEDAProjectsClickerForm.OnEDAClickerFileSaveDialogExecute := HandleOnEDAClickerFileSaveDialogExecute;
  frmEDAProjectsClickerForm.OnGetEDAClickerFileSaveDialogFileName := HandleOnGetEDAClickerFileSaveDialogFileName;
  frmEDAProjectsClickerForm.OnGetListOfFilesFromDir := HandleOnGetListOfFilesFromDir;
  frmEDAProjectsClickerForm.OnSaveTextToFile := HandleOnSaveTextToFile;
  frmEDAProjectsClickerForm.OnSetStopAllActionsOnDemand := HandleOnSetStopAllActionsOnDemand;
  frmEDAProjectsClickerForm.OnGetConnectionAddress := HandleOnGetConnectionAddress;
  frmEDAProjectsClickerForm.OnLoadMissingFileContent := HandleOnLoadMissingFileContent;

  frmClickerRemoteScreen.OnGetConnectionAddress := HandleOnGetConnectionAddress;

  frmClickerTemplateCallTree.OnSetOpenDialogMultiSelect := HandleOnSetOpenDialogMultiSelect;
  frmClickerTemplateCallTree.OnFileExists := HandleOnFileExists;
  frmClickerTemplateCallTree.OnTClkIniReadonlyFileCreate := HandleOnTClkIniReadonlyFileCreate;
  frmClickerTemplateCallTree.OnOpenDialogExecute := HandleOnOpenDialogExecute;
  frmClickerTemplateCallTree.OnGetOpenDialogFileName := HandleOnGetOpenDialogFileName;
end;


procedure TfrmEDAUIClickerMain.LoadSettings;
var
  Ini: TMemIniFile;
begin
  Ini := TMemIniFile.Create(ExtractFilePath(ParamStr(0)) + 'EDAUIClicker.ini');
  try
    Left := Ini.ReadInteger('MainWindow', 'Left', Left);
    Top := Ini.ReadInteger('MainWindow', 'Top', Top);
    Width := Ini.ReadInteger('MainWindow', 'Width', Width);
    Height := Ini.ReadInteger('MainWindow', 'Height', Height);

    if frmEDAProjectsClickerForm = nil then
      MessageBox(Handle, 'frmEDAProjectsClickerForm = nil', 'dll', MB_ICONERROR);

    if frmClickerTemplateCallTree = nil then
      MessageBox(Handle, 'frmClickerTemplateCallTree = nil', 'dll', MB_ICONERROR);

    if frmClickerRemoteScreen = nil then
      MessageBox(Handle, 'frmClickerRemoteScreen = nil', 'dll', MB_ICONERROR);

    frmEDAProjectsClickerForm.LoadSettings(Ini);
    frmClickerTemplateCallTree.LoadSettings(Ini);
    frmClickerRemoteScreen.LoadSettings(Ini);
  finally
    Ini.Free;
  end;
end;


procedure TfrmEDAUIClickerMain.SaveSettings;
var
  Ini: TMemIniFile;
begin
  Ini := TMemIniFile.Create(ExtractFilePath(ParamStr(0)) + 'EDAUIClicker.ini');
  try
    Ini.WriteInteger('MainWindow', 'Left', Left);
    Ini.WriteInteger('MainWindow', 'Top', Top);
    Ini.WriteInteger('MainWindow', 'Width', Width);
    Ini.WriteInteger('MainWindow', 'Height', Height);

    frmEDAProjectsClickerForm.SaveSettings(Ini);
    frmClickerTemplateCallTree.SaveSettings(Ini);
    frmClickerRemoteScreen.SaveSettings(Ini);

    Ini.UpdateFile;
  finally
    Ini.Free;
  end;
end;


procedure TfrmEDAUIClickerMain.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  try
    SaveSettings;
  except
  end;
end;


procedure TfrmEDAUIClickerMain.FormCreate(Sender: TObject);
begin
  FUsingMultiSelect := False;
  tmrStartup.Enabled := True;
end;


procedure TfrmEDAUIClickerMain.FormDestroy(Sender: TObject);
begin
  //
end;


procedure TfrmEDAUIClickerMain.rdgrpFileSystemClick(Sender: TObject);
begin
  frmEDAProjectsClickerForm.SetCmbTemplateContentOnAllCmbs;
end;


procedure TfrmEDAUIClickerMain.tmrStartupTimer(Sender: TObject);
begin
  tmrStartup.Enabled := False;

  SetHandles;

  try
    LoadSettings;  //called from main, after creating all forms
  except
    on E: Exception do
      MessageBox(0, PChar('Ex on loading settings: ' + E.Message), PChar(Caption), MB_ICONERROR);
  end;

  if SizeOf(Pointer) = 4 then
    lblBitness.Caption := '32-bit'
  else
    lblBitness.Caption := '64-bit';
end;


procedure TfrmEDAUIClickerMain.btnShowRemoteScreenShotFormClick(
  Sender: TObject);
begin
  frmClickerRemoteScreen.Show;
end;


procedure TfrmEDAUIClickerMain.btnShowEDAProjectsFormClick(
  Sender: TObject);
begin
  frmEDAProjectsClickerForm.Show;
end;


procedure TfrmEDAUIClickerMain.btnDisplayTemplateCallTreeClick(
  Sender: TObject);
begin
  frmClickerTemplateCallTree.Show;
end;


function TfrmEDAUIClickerMain.HandleOnGetConnectionAddress: string;
begin
  Result := frmEDAProjectsClickerForm.ConfiguredRemoteAddress; //not thread safe
end;


function TfrmEDAUIClickerMain.HandleOnGetSelectedCompFromRemoteWin: THandle;
begin
  Result := frmClickerRemoteScreen.SelectedComponentHandle;
end;


procedure TfrmEDAUIClickerMain.HandleOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);
begin
  AFileContent.LoadFromFile(AFileName);
  try
    AFileContent.LoadFromFile(AFileName);
  except
    Sleep(300); //maybe the file is in use by another thread, so wait a bit, then load again
    AFileContent.LoadFromFile(AFileName);
  end;
end;


function TfrmEDAUIClickerMain.HandleOnFileExists(const FileName: string): Boolean;
begin
  Result := FileExists(FileName);
end;


function TfrmEDAUIClickerMain.HandleOnTClkIniReadonlyFileCreate(AFileName: string): TClkIniReadonlyFile;
begin
  Result := TClkIniReadonlyFile.Create(AFileName);
end;


procedure TfrmEDAUIClickerMain.HandleOnSetTemplateOpenDialogInitialDir(AInitialDir: string);
begin
  OpenDialog1.InitialDir := AInitialDir;
end;


function TfrmEDAUIClickerMain.HandleOnOpenDialogExecute(AFilter: string): Boolean;
begin
  OpenDialog1.Filter := AFilter;
  Result := OpenDialog1.Execute;
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
end;


function TfrmEDAUIClickerMain.HandleOnGetOpenDialogFileName: string;
begin
  Result := OpenDialog1.FileName;
end;


procedure TfrmEDAUIClickerMain.HandleOnSetTemplateSaveDialogInitialDir(AInitialDir: string);
begin
  SaveDialog1.InitialDir := AInitialDir;
end;


function TfrmEDAUIClickerMain.HandleOnTemplateSaveDialogExecute: Boolean;
begin
  Result := SaveDialog1.Execute;
end;


function TfrmEDAUIClickerMain.HandleOnGetTemplateSaveDialogFileName: string;
begin
  Result := SaveDialog1.FileName;
end;


procedure TfrmEDAUIClickerMain.HandleOnSetOpenDialogMultiSelect;
begin
  OpenDialog1.Options := OpenDialog1.Options + [ofAllowMultiSelect];
  FUsingMultiSelect := True;
end;



procedure TfrmEDAUIClickerMain.HandleOnSetEDAClickerFileOpenDialogInitialDir(AInitialDir: string);
begin
  OpenDialog2.InitialDir := AInitialDir;
end;


function TfrmEDAUIClickerMain.HandleOnEDAClickerFileOpenDialogExecute: Boolean;
begin
  Result := OpenDialog2.Execute;
end;


function TfrmEDAUIClickerMain.HandleOnGetEDAClickerFileOpenDialogFileName: string;
begin
  Result := OpenDialog2.FileName;
end;


procedure TfrmEDAUIClickerMain.HandleOnSetEDAClickerFileSaveDialogInitialDir(AInitialDir: string);
begin
  SaveDialog2.InitialDir := AInitialDir;
end;


function TfrmEDAUIClickerMain.HandleOnEDAClickerFileSaveDialogExecute: Boolean;
begin
  Result := SaveDialog2.Execute;
end;


function TfrmEDAUIClickerMain.HandleOnGetEDAClickerFileSaveDialogFileName: string;
begin
  Result := SaveDialog2.FileName;
end;


procedure TfrmEDAUIClickerMain.HandleOnGetListOfFilesFromDir(ADir, AFileExtension: string; AListOfFiles: TStringList);
var
  ASearchRec: TSearchRec;
  SearchResult: Integer;
begin
  ADir := StringReplace(ADir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]);
  SearchResult := FindFirst(ADir + '\*' + AFileExtension, faArchive, ASearchRec);
  try
    while SearchResult = 0 do
    begin
      AListOfFiles.Add(ASearchRec.Name);
      SearchResult := FindNext(ASearchRec);
    end;
  finally
    FindClose(ASearchRec);
  end;
end;


procedure TfrmEDAUIClickerMain.HandleOnSaveTextToFile(AStringList: TStringList; const AFileName: string);
begin
  try
    AStringList.SaveToFile(AFileName);
  except
    on E: Exception do
    begin
      frmEDAProjectsClickerForm.memLog.Lines.Add('Error saving file: ' + E.Message);
      MessageBox(Handle, PChar('Error saving file: ' + E.Message), PChar(Caption), MB_ICONERROR);
      raise;
    end;
  end;
end;


procedure TfrmEDAUIClickerMain.HandleOnSetStopAllActionsOnDemand(Value: Boolean);
begin
  StopRemoteTemplateExecution(frmEDAProjectsClickerForm.ConfiguredRemoteAddress, 0);
  StopRemoteTemplateExecution(frmEDAProjectsClickerForm.ConfiguredRemoteAddress, 1);
  ExitRemoteTemplate(frmEDAProjectsClickerForm.ConfiguredRemoteAddress, 1);
end;


end.
