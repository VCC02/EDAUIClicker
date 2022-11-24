{
    Copyright (C) 2022 VCC
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


unit EDAProjectsClickerForm;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

interface

uses
  Windows, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, VirtualTrees, ExtCtrls, ComCtrls, IdSync,
  Buttons, ImgList, Menus, Grids, ValEdit, ClickerUtils, Types,
  IniFiles, ClickerIniFiles, ClickerFileProviderClient;

type
  TEDAProject = string;
  TEDAProjectArr = array of TEDAProject;

  TTemplateType = (ttMain, ttBeforeAll, ttAfterAll);

  TEDAProjectExecStatus = record
    Status: TActionStatus;
    AllVars: string;
  end;


  TOnSaveTextToFile = procedure(AStringList: TStringList; const AFileName: string) of object;
  TOnSetEDAClickerFileOpenDialogInitialDir = procedure(AInitialDir: string) of object;
  TOnEDAClickerFileOpenDialogExecute = function: Boolean of object;
  TOnGetEDAClickerFileOpenDialogFileName = function: string of object;
  TOnSetEDAClickerFileSaveDialogInitialDir = procedure(AInitialDir: string) of object;
  TOnEDAClickerFileSaveDialogExecute = function: Boolean of object;
  TOnGetEDAClickerFileSaveDialogFileName = function: string of object;
  TOnGetListOfFilesFromDir = procedure(ADir, AFileExtension: string; AListOfFiles: TStringList) of object;
  TOnSetStopAllActionsOnDemand = procedure(Value: Boolean) of object;


  TNodeDataProjectRec = record
    FilePath: string;
    DisplayedFile: string; //ExtractFileName(ProjectPath)
    Template: string; //can be any template from the templates folder
    BeforeAllChildTemplates: string;
    AfterAllChildTemplates: string;
    PageSize: string; //used from project if not found in document   (A4, A3, A2, B5 etc)
    PageLayout: string;  //Portrait or Landscape
    PageScaling: string; //1, 1.5, 2, 2.3, 2.5, 3 etc   means 1:1, 1.5:1, 2.3:1, 2.5:1, 3:1
    IconIndex: Integer;
    DocType: string; //a bit redundant, with regard to IconIndex
    SchDocIndex: Integer; //page number - starts at 1
    ExecStatus: TEDAProjectExecStatus;
    ExecStatusBeforeAllChildTemplates: TEDAProjectExecStatus;
    ExecStatusAfterAllChildTemplates: TEDAProjectExecStatus;
    UserNotes: string;
    ListOfEDACustomVars: string; //CRLF separated Key=Value strings   (#4#5 separated in ListOfProjects)
  end;
  PNodeDataProjectRec = ^TNodeDataProjectRec;


  TLoggingSyncObj = class(TIdSync)
  private
    FMsg: string;
  protected
    procedure DoSynchronize; override;
  end;


  { TfrmEDAProjectsClickerForm }

  TfrmEDAProjectsClickerForm = class(TForm)
    btnBrowseActionTemplatesDir: TButton;
    btnBrowseEDAPluginPath64: TButton;
    btnConnect: TButton;
    btnDisconnect: TButton;
    btnSetReplacements: TButton;
    btnBrowseEDAPluginPath32: TButton;
    btnReloadPlugin: TButton;
    grpAllowedFileDirsForServer: TGroupBox;
    grpAllowedFileExtensionsForServer: TGroupBox;
    grpClient: TGroupBox;
    grpVariables: TGroupBox;
    imglstActionStatus: TImageList;
    lbeEDAPluginPath64: TLabeledEdit;
    lblEDAPluginErrorMessage: TLabel;
    lbeEDAPluginPath32: TLabeledEdit;
    lbeClientModeServerAddress: TLabeledEdit;
    lbeConnectTimeout: TLabeledEdit;
    lbePathToTemplates: TLabeledEdit;
    lblClientMode: TLabel;
    memAllowedFileDirsForServer: TMemo;
    memAllowedFileExtensionsForServer: TMemo;
    memLog: TMemo;
    memVariables: TMemo;
    MenuItem_SetDefault: TMenuItem;
    pnlMissingFilesRequest: TPanel;
    spdbtnPlaySelectedFile: TSpeedButton;
    spdbtnPlayAllFiles: TSpeedButton;
    spdbtnStopPlaying: TSpeedButton;
    PageControlMain: TPageControl;
    TabSheetEDAPlugin: TTabSheet;
    TabSheetConnectionSettings: TTabSheet;
    TabSheetProjects: TTabSheet;
    btnAddProject: TButton;
    btnRemoveProject: TButton;
    cmbTemplate: TComboBox;
    lblTemplate: TLabel;
    lbePageSize: TLabeledEdit;
    imglstCalledTemplates: TImageList;
    imglstMainPage: TImageList;
    imglstProjects: TImageList;
    imgLstProjectsHeader: TImageList;
    lblPageLayout: TLabel;
    cmbPageLayout: TComboBox;
    tmrDisplayMissingFilesRequests: TTimer;
    tmrStartup: TTimer;
    pmTemplates: TPopupMenu;
    Selectnone1: TMenuItem;
    spdbtnExtraAddProject: TSpeedButton;
    pmExtraAddProjects: TPopupMenu;
    Addmultipleprojectsfromtexlist1: TMenuItem;
    pmExtraRemoveProjects: TPopupMenu;
    RemoveAllProjects1: TMenuItem;
    spdbtnExtraRemoveProject: TSpeedButton;
    lbeSearchProject: TLabeledEdit;
    chkDisplayFullPath: TCheckBox;
    spdbtnExtraPlayAllFiles: TSpeedButton;
    pmExtraPlayAllFiles: TPopupMenu;
    PlayAllFilesbelowselectedincludingselected1: TMenuItem;
    pmExtraSaveListOfProjects: TPopupMenu;
    SaveListOfProjectsAs1: TMenuItem;
    tmrPlayIconAnimation: TTimer;
    grpListOfProjects: TGroupBox;
    spdbtnExtraSaveListOfProjects: TSpeedButton;
    btnLoadListOfProjects: TButton;
    btnNewListOfProjects: TButton;
    bitbtnSaveListOfProjects: TBitBtn;
    lblLoadedListOfProjects: TLabel;
    imgListOfProjectsIcon: TImage;
    lbePageScaling: TLabeledEdit;
    memPlayingDescription: TMemo;
    spdbtnExtraPlaySelectedFile: TSpeedButton;
    PlayAllFilesInDebuggingMode: TMenuItem;
    PlayAllFilesbelowselectedindebuggingmode1: TMenuItem;
    pmExtraPlaySelectedFile: TPopupMenu;
    PlaySelectedFileInDebuggingMode: TMenuItem;
    cmbBeforeAllChildTemplates: TComboBox;
    lblBeforeAllChildTemplates: TLabel;
    cmbAfterAllChildTemplates: TComboBox;
    lblAfterAllChildTemplates: TLabel;
    chkPlayPrjOnPlayAll: TCheckBox;
    chkPlayPrjMainTemplateOnPlayAll: TCheckBox;
    bitbtnUpdateProject: TBitBtn;
    chkStayOnTop: TCheckBox;
    pmVSTProjects: TPopupMenu;
    ExpandAll1: TMenuItem;
    CollapseAll1: TMenuItem;
    lblListOfProjectsStatus: TLabel;
    lbeUserNotes: TLabeledEdit;
    N1: TMenuItem;
    CheckAllProjects1: TMenuItem;
    CheckAll1: TMenuItem;
    UncheckAll1: TMenuItem;
    N2: TMenuItem;
    Displayprojectsinbold1: TMenuItem;
    N3: TMenuItem;
    CopyFullPathToClipboard1: TMenuItem;
    CopyDirectoryPathToClipboard1: TMenuItem;
    grpExecStatusDebugVars: TGroupBox;
    vallstStatusVariables: TValueListEditor;
    grpEDAFileCustomVars: TGroupBox;
    vallstEDAFileCustomVars: TValueListEditor;
    pmEDAFileCustomVars: TPopupMenu;
    MenuItem_EDAFileAddCustomVarRow: TMenuItem;
    MenuItem_EDAFileRemoveCustomVarRow: TMenuItem;
    N4: TMenuItem;
    SelectAll1: TMenuItem;
    SelectAllProjects1: TMenuItem;
    SelectAllProjectDocuments1: TMenuItem;
    UnselectAll1: TMenuItem;
    CheckAllProjectDocumentsForSelectedProjects: TMenuItem;
    CheckAllSelectedProjects1: TMenuItem;
    pnlvstProjects: TPanel;
    procedure btnAddProjectClick(Sender: TObject);
    procedure btnBrowseActionTemplatesDirClick(Sender: TObject);
    procedure btnBrowseEDAPluginPath32Click(Sender: TObject);
    procedure btnBrowseEDAPluginPath64Click(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure btnDisconnectClick(Sender: TObject);
    procedure btnReloadPluginClick(Sender: TObject);
    procedure cmbAfterAllChildTemplatesMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure cmbBeforeAllChildTemplatesMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure cmbTemplateMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnRemoveProjectClick(Sender: TObject);
    procedure MenuItem_SetDefaultClick(Sender: TObject);
    procedure tmrDisplayMissingFilesRequestsTimer(Sender: TObject);
    procedure vstProjectsGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: {$IFDEF FPC} string {$ELSE} WideString {$ENDIF});
    procedure cmbTemplateDropDown(Sender: TObject);
    procedure spdbtnPlaySelectedFileClick(Sender: TObject);
    procedure spdbtnPlayAllFilesClick(Sender: TObject);
    procedure spdbtnStopPlayingClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure vstProjectsGetImageIndex(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
      var Ghosted: Boolean; var ImageIndex: Integer);
    procedure btnUpdateProjectClick(Sender: TObject);
    procedure btnLoadListOfProjectsClick(Sender: TObject);
    procedure btnSaveListOfProjectsClick(Sender: TObject);
    procedure tmrStartupTimer(Sender: TObject);
    procedure Selectnone1Click(Sender: TObject);
    procedure vstProjectsClick(Sender: TObject);
    procedure cmbTemplateCloseUp(Sender: TObject);
    procedure Addmultipleprojectsfromtextlist1Click(Sender: TObject);
    procedure spdbtnExtraAddProjectClick(Sender: TObject);
    procedure spdbtnExtraRemoveProjectClick(Sender: TObject);
    procedure RemoveAllProjects1Click(Sender: TObject);
    procedure lbeSearchProjectChange(Sender: TObject);
    procedure chkDisplayFullPathClick(Sender: TObject);
    procedure spdbtnExtraPlayAllFilesClick(Sender: TObject);
    procedure PlayAllFilesbelowselectedincludingselected1Click(Sender: TObject);
    procedure btnSetReplacementsClick(Sender: TObject);
    procedure SaveListOfProjectsAs1Click(Sender: TObject);
    procedure btnNewListOfProjectsClick(Sender: TObject);
    procedure tmrPlayIconAnimationTimer(Sender: TObject);
    procedure PlayAllFilesInDebuggingModeClick(Sender: TObject);
    procedure PlayAllFilesbelowselectedindebuggingmode1Click(Sender: TObject);
    procedure spdbtnExtraSaveListOfProjectsClick(Sender: TObject);
    procedure PlaySelectedFileInDebuggingModeClick(Sender: TObject);
    procedure spdbtnExtraPlaySelectedFileClick(Sender: TObject);
    procedure vstProjectsGetImageIndexEx(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
      var Ghosted: Boolean; var ImageIndex: Integer;
      var ImageList: TCustomImageList);
    procedure cmbTemplateMouseEnter(Sender: TObject);
    procedure chkPlayPrjOnPlayAllClick(Sender: TObject);
    procedure chkPlayPrjMainTemplateOnPlayAllClick(Sender: TObject);
    procedure cmbTemplateChange(Sender: TObject);
    procedure cmbBeforeAllChildTemplatesChange(Sender: TObject);
    procedure cmbAfterAllChildTemplatesChange(Sender: TObject);
    procedure lbePageSizeChange(Sender: TObject);
    procedure cmbPageLayoutChange(Sender: TObject);
    procedure lbePageScalingChange(Sender: TObject);
    procedure chkStayOnTopClick(Sender: TObject);
    procedure ExpandAll1Click(Sender: TObject);
    procedure CollapseAll1Click(Sender: TObject);
    procedure lbeUserNotesChange(Sender: TObject);
    procedure vstProjectsKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure vstProjectsInitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure vstProjectsChecked(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure CheckAllProjects1Click(Sender: TObject);
    procedure CheckAll1Click(Sender: TObject);
    procedure UncheckAll1Click(Sender: TObject);
    procedure Displayprojectsinbold1Click(Sender: TObject);
    procedure vstProjectsPaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      TextType: TVSTTextType);
    procedure CopyFullPathToClipboard1Click(Sender: TObject);
    procedure CopyDirectoryPathToClipboard1Click(Sender: TObject);
    procedure vallstEDAFileCustomVarsSetEditText(Sender: TObject; ACol,
      ARow: Integer; const Value: string);
    procedure vallstEDAFileCustomVarsValidate(Sender: TObject; ACol,
      ARow: Integer; const KeyName, KeyValue: string);
    procedure MenuItem_EDAFileAddCustomVarRowClick(Sender: TObject);
    procedure MenuItem_EDAFileRemoveCustomVarRowClick(Sender: TObject);
    procedure SelectAll1Click(Sender: TObject);
    procedure UnselectAll1Click(Sender: TObject);
    procedure SelectAllProjects1Click(Sender: TObject);
    procedure SelectAllProjectDocuments1Click(Sender: TObject);
    procedure vstProjectsMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure vstProjectsMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grpEDAFileCustomVarsMouseLeave(Sender: TObject);
    procedure CheckAllProjectDocumentsForSelectedProjectsClick(Sender: TObject);
    procedure CheckAllSelectedProjects1Click(Sender: TObject);
  private
    { Private declarations }

    FStopAllActionsOnDemand: Boolean;
    FClickerProjectsDir: string;
    FEDAProjectsDir: string;
    FTextListOfProjectsDir: string;
    FListOfProjectsFileName: string;
    FModified: Boolean;
    FToBeUpdated: Boolean;
    FFullTemplatesDir: string;
    FPollForMissingServerFiles: TPollForMissingServerFiles;
    FConfiguredRemoteAddress: string;

    FDefaultTemplate_Main: string;
    FDefaultTemplate_Before: string;
    FDefaultTemplate_After: string;
    LastClickedComboBoxDefaultString: string;

    FProcessingMissingFilesRequestByClient: Boolean; //for activity info

    FOnFileExists: TOnFileExists;
    FOnTClkIniReadonlyFileCreate: TOnTClkIniReadonlyFileCreate;
    FOnSaveTextToFile: TOnSaveTextToFile;
    FOnSetEDAClickerFileOpenDialogInitialDir: TOnSetEDAClickerFileOpenDialogInitialDir;
    FOnEDAClickerFileOpenDialogExecute: TOnEDAClickerFileOpenDialogExecute;
    FOnGetEDAClickerFileOpenDialogFileName: TOnGetEDAClickerFileOpenDialogFileName;
    FOnSetEDAClickerFileSaveDialogInitialDir: TOnSetEDAClickerFileSaveDialogInitialDir;
    FOnEDAClickerFileSaveDialogExecute: TOnEDAClickerFileSaveDialogExecute;
    FOnGetEDAClickerFileSaveDialogFileName: TOnGetEDAClickerFileSaveDialogFileName;
    FOnGetListOfFilesFromDir: TOnGetListOfFilesFromDir;
    FOnSetStopAllActionsOnDemand: TOnSetStopAllActionsOnDemand;

    FOnGetConnectionAddress: TOnGetConnectionAddress;
    FOnLoadMissingFileContent: TOnLoadMissingFileContent;

    FPopUpComboBox: TComboBox;
    vstProjects: TVirtualStringTree;

    function DoOnFileExists(const AFileName: string): Boolean;
    function DoOnTClkIniReadonlyFileCreate(AFileName: string): TClkIniReadonlyFile;
    procedure DoOnSaveTextToFile(AStringList: TStringList; const AFileName: string);

    procedure DoOnSetEDAClickerFileOpenDialogInitialDir(AInitialDir: string);
    function DoOnEDAClickerFileOpenDialogExecute: Boolean;
    function DoOnGetEDAClickerFileOpenDialogFileName: string;
    procedure DoOnSetEDAClickerFileSaveDialogInitialDir(AInitialDir: string);
    function DoOnEDAClickerFileSaveDialogExecute: Boolean;
    function DoOnGetEDAClickerFileSaveDialogFileName: string;

    procedure DoOnGetListOfFilesFromDir(ADir, AFileExtension: string; AListOfFiles: TStringList);
    procedure DoOnSetStopAllActionsOnDemand(Value: Boolean);

    function DoOnGetConnectionAddress: string;
    procedure DoOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);

    procedure HandleOnBeforeRequestingListOfMissingFiles;  //client thread calls this, without UI sync
    procedure HandleOnAfterRequestingListOfMissingFiles;   //client thread calls this, without UI sync
    function HandleOnFileExists(const AFileName: string): Boolean;
    procedure HandleLogMissingServerFile(AMsg: string);
    procedure HandleOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);

    procedure CreateRemainingUIComponents;
    procedure AddToLog(s: string);

    procedure SetModified(Value: Boolean);
    procedure SetToBeUpdated(Value: Boolean);
    procedure SetListOfProjectsFileName(Value: string);

    procedure SetAllNodesToExpandedOrCollapsed(IsExpanded: Boolean);
    procedure UpdateControlsFromNodeData(ANodeData: PNodeDataProjectRec);
    procedure SetEDAPluginErrorMessage(AMsg: string; AColor: TColor);
    procedure PluginStartup;

    procedure HandleProjectSelectionChange;
    procedure RemoveSelectedEDAFiles;
    procedure CheckAllProjects(IsSelectedOnly: Boolean);
    procedure UncheckAllProjects;

    function LoadEDAProjectToVst(FileName: string; PageSize, PageLayout, PageScaling, ActionTemplate, BeforeAllChildTemplates, AfterAllChildTemplates, UserNotes, EDAFileCustomVars: string): PVirtualNode;
    procedure SetCmbTemplateContent(AComboBox: TComboBox);
    procedure PlayNode(Node: PVirtualNode; AIsDebugging: Boolean; ATemplateType: TTemplateType);

    procedure SetDocNodeTemplatesFromClickerProjectFile(Ini: TClkIniReadonlyFile; ProjectNode: PVirtualNode);
    procedure UpdateClickerProjectFileWithDocNodeTemplate(Ini: TStringList; ProjectNode: PVirtualNode);
    procedure LoadListOfProjects(AFileName: string);
    function SaveListOfProjects(AFileName: string): Boolean;
    procedure SaveListOfProjectsWithDialog;
    procedure GetCurrentProjectList(AList: TStringList);
    procedure PlayAllStartingAtNode(Node: PVirtualNode; IsDebugging: Boolean);

    function ExecuteTemplateOnUIClicker(ATemplatePath, AFileLocation, ACustomVarsAndValues: string; AIsDebugging: Boolean): string;

    property ClickerProjectsDir: string read FClickerProjectsDir write FClickerProjectsDir;
    property EDAProjectsDir: string read FEDAProjectsDir write FEDAProjectsDir;
    property TextListOfProjectsDir: string read FTextListOfProjectsDir write FTextListOfProjectsDir;
    property ToBeUpdated: Boolean write SetToBeUpdated;
    property ListOfProjectsFileName: string read FListOfProjectsFileName write SetListOfProjectsFileName;

    property FullTemplatesDir: string read FFullTemplatesDir write FFullTemplatesDir;
  public
    { Public declarations }
    procedure LoadSettings(AIni: TMemIniFile);
    procedure SaveSettings(AIni: TMemIniFile);

    procedure SetCmbTemplateContentOnAllCmbs;

    property Modified: Boolean read FModified write SetModified;
    property ConfiguredRemoteAddress: string read FConfiguredRemoteAddress;

    property OnFileExists: TOnFileExists write FOnFileExists;
    property OnTClkIniReadonlyFileCreate: TOnTClkIniReadonlyFileCreate write FOnTClkIniReadonlyFileCreate;
    property OnSaveTextToFile: TOnSaveTextToFile write FOnSaveTextToFile;

    property OnSetEDAClickerFileOpenDialogInitialDir: TOnSetEDAClickerFileOpenDialogInitialDir write FOnSetEDAClickerFileOpenDialogInitialDir;
    property OnEDAClickerFileOpenDialogExecute: TOnEDAClickerFileOpenDialogExecute write FOnEDAClickerFileOpenDialogExecute;
    property OnGetEDAClickerFileOpenDialogFileName: TOnGetEDAClickerFileOpenDialogFileName write FOnGetEDAClickerFileOpenDialogFileName;
    property OnSetEDAClickerFileSaveDialogInitialDir: TOnSetEDAClickerFileSaveDialogInitialDir write FOnSetEDAClickerFileSaveDialogInitialDir;
    property OnEDAClickerFileSaveDialogExecute: TOnEDAClickerFileSaveDialogExecute write FOnEDAClickerFileSaveDialogExecute;
    property OnGetEDAClickerFileSaveDialogFileName: TOnGetEDAClickerFileSaveDialogFileName write FOnGetEDAClickerFileSaveDialogFileName;

    property OnGetListOfFilesFromDir: TOnGetListOfFilesFromDir write FOnGetListOfFilesFromDir;
    property OnSetStopAllActionsOnDemand: TOnSetStopAllActionsOnDemand write FOnSetStopAllActionsOnDemand;

    property OnGetConnectionAddress: TOnGetConnectionAddress write FOnGetConnectionAddress;
    property OnLoadMissingFileContent: TOnLoadMissingFileContent write FOnLoadMissingFileContent;
  end;


var
  frmEDAProjectsClickerForm: TfrmEDAProjectsClickerForm;


implementation

{$R *.frm}


uses
  Clipbrd, ClickerActionsClient, EDAUIClickerPlugins, EDAUIClickerPluginsCommon;


const
  CDefaultTemplateString = 'Default: ';

procedure TLoggingSyncObj.DoSynchronize;
begin
  frmEDAProjectsClickerForm.AddToLog(FMsg);
end;


procedure AddToLogFromThread(s: string);
var
  SyncObj: TLoggingSyncObj;
begin
  SyncObj := TLoggingSyncObj.Create;
  try
    SyncObj.FMsg := s;
    SyncObj.Synchronize;
  finally
    SyncObj.Free;
  end;
end;


function GetTemplateNameFromNode(AVst: TVirtualStringTree; Node: PVirtualNode; out PrjNodeData: PNodeDataProjectRec; ATemplateType: TTemplateType): string;
var
  NodeData: PNodeDataProjectRec;
  TemplateFileName: string;
begin
  NodeData := AVst.GetNodeData(Node);
  PrjNodeData := NodeData; //point to project first (used by action vars)

  case ATemplateType of
    ttMain: TemplateFileName := NodeData^.Template;
    ttBeforeAll: TemplateFileName := NodeData^.BeforeAllChildTemplates;
    ttAfterAll: TemplateFileName := NodeData^.AfterAllChildTemplates;
  end;

  if (NodeData^.DocType = CSCH) or (NodeData^.DocType = CPCB) then
  begin
    PrjNodeData := AVst.GetNodeData(Node^.Parent);

    case ATemplateType of
      ttMain:
      begin
        if TemplateFileName = '' then
          TemplateFileName := PrjNodeData^.Template;    //use project template if not set for document
      end;

      ttBeforeAll: TemplateFileName := '';
      ttAfterAll: TemplateFileName := '';
    end;
  end;

  Result := TemplateFileName;
end;


procedure UpdateOverridenValuesForActionVars(OverridenValues: TStrings; NodeData, PrjNodeData: PNodeDataProjectRec; DocType: string);
var
  FileExt, PrjFileExt: string;
  FullPathFileName, FullPathFileNameNoExt, FullPathProjectName, FullPathProjectNameNoExt: string;
  CustomVars: TStringList;
begin
  FileExt := UpperCase(ExtractFileExt(NodeData^.DisplayedFile));
  PrjFileExt := UpperCase(ExtractFileExt(PrjNodeData^.DisplayedFile));

  FullPathFileName := NodeData^.FilePath;
  FullPathFileNameNoExt := Copy(NodeData^.FilePath, 1, Length(NodeData^.FilePath) - Length(FileExt));
  FullPathProjectName := PrjNodeData^.FilePath;
  FullPathProjectNameNoExt := Copy(PrjNodeData^.FilePath, 1, Length(PrjNodeData^.FilePath) - Length(PrjFileExt));

  OverridenValues.Add('$PageSize$=' + NodeData^.PageSize);
  OverridenValues.Add('$PageLayout$=' + NodeData^.PageLayout);
  OverridenValues.Add('$PageScaling$=' + NodeData^.PageScaling);
  OverridenValues.Add('$EDADocType$=' + DocType);
  OverridenValues.Add('$SchDocIndex$=' + IntToStr(NodeData^.SchDocIndex));  //starts at 1

  OverridenValues.Add('$FullPathFileName$=' + FullPathFileName);
  OverridenValues.Add('$FullPathFileNameNoExt$=' + FullPathFileNameNoExt);
  OverridenValues.Add('$FullPathProjectName$=' + FullPathProjectName);
  OverridenValues.Add('$FullPathProjectNameNoExt$=' + FullPathProjectNameNoExt);

  OverridenValues.Add('$FileName$=' + ExtractFileName(FullPathFileName));
  OverridenValues.Add('$FileNameNoExt$=' + ExtractFileName(FullPathFileNameNoExt));
  OverridenValues.Add('$ProjectName$=' + ExtractFileName(FullPathProjectName));
  OverridenValues.Add('$ProjectNameNoExt$=' + ExtractFileName(FullPathProjectNameNoExt));

  OverridenValues.Add('$FullPathProjectDir$=' + ExtractFileDir(FullPathFileName));

  CustomVars := TStringList.Create;
  try
    CustomVars.Text := NodeData^.ListOfEDACustomVars;
    OverridenValues.AddStrings(CustomVars);
  finally
    CustomVars.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.LoadSettings(AIni: TMemIniFile);
var
  i, n: Integer;
begin
  Left := AIni.ReadInteger('EDAProjectsWindow', 'Left', Left);
  Top := AIni.ReadInteger('EDAProjectsWindow', 'Top', Top);
  Width := AIni.ReadInteger('EDAProjectsWindow', 'Width', Width);
  Height := AIni.ReadInteger('EDAProjectsWindow', 'Height', Height);
  chkStayOnTop.Checked := AIni.ReadBool('EDAProjectsWindow', 'StayOnTop', chkStayOnTop.Checked);
  Displayprojectsinbold1.Checked := AIni.ReadBool('EDAProjectsWindow', 'DisplayProjectsInBold', Displayprojectsinbold1.Checked);

  lbeClientModeServerAddress.Text := AIni.ReadString('EDAProjectsWindow', 'ClientModeServerAddress', lbeClientModeServerAddress.Text);
  lbeConnectTimeout.Text := AIni.ReadString('EDAProjectsWindow', 'ConnectTimeout', lbeConnectTimeout.Text);

  ClickerProjectsDir := AIni.ReadString('Dirs', 'ClickerProjectsDir', '');
  EDAProjectsDir := AIni.ReadString('Dirs', 'EDAProjectsDir', '');
  TextListOfProjectsDir := AIni.ReadString('Dirs', 'TextListOfProjectsDir', '');
  FullTemplatesDir := AIni.ReadString('Dirs', 'FullTemplatesDir', '$AppDir$\ActionTemplates');
  if FFullTemplatesDir > '' then
    if FFullTemplatesDir[Length(FFullTemplatesDir)] = '\' then
      FFullTemplatesDir := Copy(FFullTemplatesDir, 1, Length(FFullTemplatesDir) - 1);

  lbePathToTemplates.Text := StringReplace(FullTemplatesDir, ExtractFileDir(ParamStr(0)), '$AppDir$', [rfReplaceAll]);

  n := AIni.ReadInteger('EDAProjectsWindow', 'AllowedFileExtensionsForServer.Count', 0);
  if n > 0 then
  begin
    memAllowedFileExtensionsForServer.Clear;
    for i := 0 to n - 1 do
      memAllowedFileExtensionsForServer.Lines.Add(AIni.ReadString('EDAProjectsWindow', 'AllowedFileExtensionsForServer_' + IntToStr(i), '.bmp'));
  end;

  n := AIni.ReadInteger('EDAProjectsWindow', 'AllowedFileDirsForServer.Count', 0);
  if n > 0 then
  begin
    memAllowedFileDirsForServer.Clear;
    for i := 0 to n - 1 do
      memAllowedFileDirsForServer.Lines.Add(AIni.ReadString('EDAProjectsWindow', 'AllowedFileDirsForServer_' + IntToStr(i), ''));
  end
  else
    memAllowedFileDirsForServer.Lines.Add(ExtractFilePath(ParamStr(0)) + 'ActionTemplates');

  lbeEDAPluginPath32.Text := AIni.ReadString('EDAProjectsWindow', 'EDAPluginPath32', lbeEDAPluginPath32.Text);
  lbeEDAPluginPath64.Text := AIni.ReadString('EDAProjectsWindow', 'EDAPluginPath64', lbeEDAPluginPath32.Text);

  FDefaultTemplate_Main := AIni.ReadString('EDAProjectsWindow', 'DefaultTemplate_Main', '');
  FDefaultTemplate_Before := AIni.ReadString('EDAProjectsWindow', 'DefaultTemplate_Before', '');
  FDefaultTemplate_After := AIni.ReadString('EDAProjectsWindow', 'DefaultTemplate_After', '');

  cmbTemplate.Hint := cmbTemplate.Hint + #13#10 + CDefaultTemplateString + FDefaultTemplate_Main;
  cmbBeforeAllChildTemplates.Hint := cmbBeforeAllChildTemplates.Hint + #13#10 + CDefaultTemplateString + FDefaultTemplate_Before;
  cmbAfterAllChildTemplates.Hint := cmbAfterAllChildTemplates.Hint + #13#10 + CDefaultTemplateString + FDefaultTemplate_After;

  if SizeOf(Pointer) = 4 then
  begin
    lbeEDAPluginPath32.Enabled := True;
    btnBrowseEDAPluginPath32.Enabled := True;

    lbeEDAPluginPath64.Enabled := False;
    btnBrowseEDAPluginPath64.Enabled := False;

    btnReloadPlugin.Top := btnBrowseEDAPluginPath32.Top;
  end
  else
  begin
    lbeEDAPluginPath32.Enabled := False;
    btnBrowseEDAPluginPath32.Enabled := False;

    lbeEDAPluginPath64.Enabled := True;
    btnBrowseEDAPluginPath64.Enabled := True;

    btnReloadPlugin.Top := btnBrowseEDAPluginPath64.Top;
  end;

  tmrStartup.Enabled := True;   //LoadSettings is called from main window, so enable the timer here, instead of FormCreate
end;


procedure TfrmEDAProjectsClickerForm.SaveSettings(AIni: TMemIniFile);
var
  i, n: Integer;
begin
  AIni.WriteInteger('EDAProjectsWindow', 'Left', Left);
  AIni.WriteInteger('EDAProjectsWindow', 'Top', Top);
  AIni.WriteInteger('EDAProjectsWindow', 'Width', Width);
  AIni.WriteInteger('EDAProjectsWindow', 'Height', Height);
  AIni.WriteBool('EDAProjectsWindow', 'StayOnTop', chkStayOnTop.Checked);
  AIni.WriteBool('EDAProjectsWindow', 'DisplayProjectsInBold', Displayprojectsinbold1.Checked);

  AIni.WriteString('EDAProjectsWindow', 'ClientModeServerAddress', lbeClientModeServerAddress.Text);
  AIni.WriteString('EDAProjectsWindow', 'ConnectTimeout', lbeConnectTimeout.Text);

  AIni.WriteString('Dirs', 'ClickerProjectsDir', ClickerProjectsDir);
  AIni.WriteString('Dirs', 'EDAProjectsDir', EDAProjectsDir);
  AIni.WriteString('Dirs', 'TextListOfProjectsDir', TextListOfProjectsDir);
  AIni.WriteString('Dirs', 'FullTemplatesDir', StringReplace(FullTemplatesDir, ExtractFileDir(ParamStr(0)), '$AppDir$', [rfReplaceAll]));

  n := memAllowedFileExtensionsForServer.Lines.Count;
  AIni.WriteInteger('EDAProjectsWindow', 'AllowedFileExtensionsForServer.Count', n);

  for i := 0 to n - 1 do
    AIni.WriteString('EDAProjectsWindow', 'AllowedFileExtensionsForServer_' + IntToStr(i), memAllowedFileExtensionsForServer.Lines.Strings[i]);

  n := memAllowedFileDirsForServer.Lines.Count;
  AIni.WriteInteger('EDAProjectsWindow', 'AllowedFileDirsForServer.Count', n);

  for i := 0 to n - 1 do
    AIni.WriteString('EDAProjectsWindow', 'AllowedFileDirsForServer_' + IntToStr(i), memAllowedFileDirsForServer.Lines.Strings[i]);

  AIni.WriteString('EDAProjectsWindow', 'EDAPluginPath32', lbeEDAPluginPath32.Text);
  AIni.WriteString('EDAProjectsWindow', 'EDAPluginPath64', lbeEDAPluginPath64.Text);
end;


function TfrmEDAProjectsClickerForm.DoOnFileExists(const AFileName: string): Boolean;
begin
  if not Assigned(FOnFileExists) then
    raise Exception.Create('OnFileExists is not assigned.')
  else
    Result := FOnFileExists(AFileName);
end;


function TfrmEDAProjectsClickerForm.DoOnTClkIniReadonlyFileCreate(AFileName: string): TClkIniReadonlyFile;
begin
  if not Assigned(FOnTClkIniReadonlyFileCreate) then
    raise Exception.Create('OnTClkIniReadonlyFileCreate is not assigned.')
  else
    Result := FOnTClkIniReadonlyFileCreate(AFileName);
end;


procedure TfrmEDAProjectsClickerForm.DoOnSaveTextToFile(AStringList: TStringList; const AFileName: string);
begin
  if not Assigned(FOnSaveTextToFile) then
    raise Exception.Create('OnSaveTextToFile is not assigned.')
  else
    FOnSaveTextToFile(AStringList, AFileName);
end;


procedure TfrmEDAProjectsClickerForm.DoOnSetEDAClickerFileOpenDialogInitialDir(AInitialDir: string);
begin
  if not Assigned(FOnSetEDAClickerFileOpenDialogInitialDir) then
    raise Exception.Create('OnSetEDAClickerFileOpenDialogInitialDir is not assigned.')
  else
    FOnSetEDAClickerFileOpenDialogInitialDir(AInitialDir);
end;


function TfrmEDAProjectsClickerForm.DoOnEDAClickerFileOpenDialogExecute: Boolean;
begin
  if not Assigned(FOnEDAClickerFileOpenDialogExecute) then
    raise Exception.Create('OnEDAClickerFileOpenDialogExecute is not assigned.')
  else
    Result := FOnEDAClickerFileOpenDialogExecute;
end;


function TfrmEDAProjectsClickerForm.DoOnGetEDAClickerFileOpenDialogFileName: string;
begin
  if not Assigned(FOnGetEDAClickerFileOpenDialogFileName) then
    raise Exception.Create('OnGetEDAClickerFileOpenDialogFileName is not assigned.')
  else
    Result := FOnGetEDAClickerFileOpenDialogFileName;
end;


procedure TfrmEDAProjectsClickerForm.DoOnSetEDAClickerFileSaveDialogInitialDir(AInitialDir: string);
begin
  if not Assigned(FOnSetEDAClickerFileSaveDialogInitialDir) then
    raise Exception.Create('OnSetEDAClickerFileSaveDialogInitialDir is not assigned.')
  else
    FOnSetEDAClickerFileSaveDialogInitialDir(AInitialDir);
end;


function TfrmEDAProjectsClickerForm.DoOnEDAClickerFileSaveDialogExecute: Boolean;
begin
  if not Assigned(FOnEDAClickerFileSaveDialogExecute) then
    raise Exception.Create('OnEDAClickerFileSaveDialogExecute is not assigned.')
  else
    Result := FOnEDAClickerFileSaveDialogExecute;
end;


function TfrmEDAProjectsClickerForm.DoOnGetEDAClickerFileSaveDialogFileName: string;
begin
  if not Assigned(FOnGetEDAClickerFileSaveDialogFileName) then
    raise Exception.Create('OnGetEDAClickerFileSaveDialogFileName is not assigned.')
  else
    Result := FOnGetEDAClickerFileSaveDialogFileName;
end;


procedure TfrmEDAProjectsClickerForm.DoOnGetListOfFilesFromDir(ADir, AFileExtension: string; AListOfFiles: TStringList);
begin
  if not Assigned(FOnGetListOfFilesFromDir) then
    raise Exception.Create('OnGetListOfFilesFromDir is not assigned.')
  else
    FOnGetListOfFilesFromDir(ADir, AFileExtension, AListOfFiles);
end;


procedure TfrmEDAProjectsClickerForm.DoOnSetStopAllActionsOnDemand(Value: Boolean);
begin
  if not Assigned(FOnSetStopAllActionsOnDemand) then
    raise Exception.Create('OnSetStopAllActionsOnDemand is not assigned.')
  else
    FOnSetStopAllActionsOnDemand(Value);
end;


function TfrmEDAProjectsClickerForm.DoOnGetConnectionAddress: string;
begin
  if not Assigned(FOnGetConnectionAddress) then
  begin
    Result := '';
    Exit;
  end;

  Result := FOnGetConnectionAddress();
end;


procedure TfrmEDAProjectsClickerForm.DoOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);
begin
  if (AFileName <> '') and (AFileName[1] = PathDelim) then
    AFileName := UpperCase(StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]) + AFileName);

  if not Assigned(FOnLoadMissingFileContent) then
    raise Exception.Create('OnLoadMissingFileContent is not assigned.')
  else
    FOnLoadMissingFileContent(AFileName, AFileContent);
end;


procedure TfrmEDAProjectsClickerForm.HandleOnBeforeRequestingListOfMissingFiles;  //client thread calls this, without UI sync
begin
  FProcessingMissingFilesRequestByClient := True;
end;


procedure TfrmEDAProjectsClickerForm.HandleOnAfterRequestingListOfMissingFiles;   //client thread calls this, without UI sync
begin
  //FProcessingMissingFilesRequestByClient := False;  //let the timer reset the flag
end;


function TfrmEDAProjectsClickerForm.HandleOnFileExists(const AFileName: string): Boolean;
var
  ExpandedFileName: string;
  WillExpandFileName: Boolean;
begin
  ExpandedFileName := AFileName;
  WillExpandFileName := False;

  if (ExpandedFileName <> '') and (ExpandedFileName[1] = PathDelim) then
    ExpandedFileName := StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]) + ExpandedFileName;

  if ExtractFileName(ExpandedFileName) = ExpandedFileName then //files without paths are expected to be found in $AppDir$\ActionTemplates
  begin
    ExpandedFileName := StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]) + PathDelim + ExpandedFileName;
    WillExpandFileName := True;
  end;

  Result := DoOnFileExists(ExpandedFileName);

  if WillExpandFileName then
    AddToLogFromThread('File existence = ' + IntToStr(Ord(Result)) + ' on ' + ExpandedFileName);
end;


procedure TfrmEDAProjectsClickerForm.HandleLogMissingServerFile(AMsg: string);
begin
  AddToLogFromThread(AMsg);
end;


procedure TfrmEDAProjectsClickerForm.HandleOnLoadMissingFileContent(AFileName: string; AFileContent: TMemoryStream);
var
  ExpandedFileName: string;
begin
  ExpandedFileName := AFileName;

  if ExtractFileName(ExpandedFileName) = ExpandedFileName then //files without paths are expected to be found in $AppDir$\ActionTemplates
  begin
    ExpandedFileName := StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]) + PathDelim + ExpandedFileName;
    AddToLogFromThread('Loading file with expanded name: ' + ExpandedFileName);
  end;

  DoOnLoadMissingFileContent(ExpandedFileName, AFileContent);
end;


function TfrmEDAProjectsClickerForm.LoadEDAProjectToVst(FileName: string; PageSize, PageLayout, PageScaling, ActionTemplate, BeforeAllChildTemplates, AfterAllChildTemplates, UserNotes, EDAFileCustomVars: string): PVirtualNode;
var
  NodeData: PNodeDataProjectRec;
  Node, DocNode: PVirtualNode;
  PrjPageSize: string;
  i: Integer;
  SchDocIndex: Integer;
  ListOfDocuments, ListOfTypes: TStringList;
begin
  Node := vstProjects.InsertNode(vstProjects.RootNode, amAddChildLast);
  Result := Node;
  NodeData := vstProjects.GetNodeData(Node);
  NodeData^.FilePath := FileName;
  NodeData^.DisplayedFile := ExtractFileName(NodeData^.FilePath);
  NodeData^.PageSize := PageSize;
  if NodeData^.PageSize = '' then
    NodeData^.PageSize := 'A4';

  PrjPageSize := NodeData^.PageSize; //used for documents if not provided
  NodeData^.PageLayout := PageLayout;
  NodeData^.PageScaling := PageScaling;
  NodeData^.Template := ActionTemplate;
  NodeData^.BeforeAllChildTemplates := BeforeAllChildTemplates;
  NodeData^.AfterAllChildTemplates := AfterAllChildTemplates;
  NodeData^.UserNotes := UserNotes;
  NodeData^.ListOfEDACustomVars := EDAFileCustomVars;
  NodeData^.SchDocIndex := -2;

  NodeData^.IconIndex := 0;
  NodeData^.DocType := CPRJ;

  if not FileExists(FileName) then
  begin
    NodeData^.DisplayedFile := NodeData^.DisplayedFile + ' (file not found)';
    NodeData^.IconIndex := -1;
    Exit;
  end;

  ListOfDocuments := TStringList.Create;
  ListOfTypes := TStringList.Create;
  try
    GetEDAListsOfDocuments(FileName, ListOfDocuments, ListOfTypes);

    SchDocIndex := 0;
    for i := 0 to ListOfDocuments.Count - 1 do
    begin
      DocNode := vstProjects.InsertNode(Node, amAddChildLast);
      NodeData := vstProjects.GetNodeData(DocNode);

      NodeData^.FilePath := ExtractFilePath(FileName) + ListOfDocuments.Strings[i];
      NodeData^.DisplayedFile := ExtractFileName(NodeData^.FilePath);
      NodeData^.Template := '';
      NodeData^.BeforeAllChildTemplates := '';
      NodeData^.AfterAllChildTemplates := '';
      NodeData^.PageSize := PageSize;
      if NodeData^.PageSize = '' then
        NodeData^.PageSize := PrjPageSize;

      NodeData^.PageLayout := PageLayout;
      NodeData^.PageScaling := PageScaling;
      NodeData^.UserNotes := UserNotes;
      NodeData^.ListOfEDACustomVars := '';

      NodeData^.IconIndex := -1;
      NodeData^.DocType := ListOfTypes.Strings[i];

      if NodeData^.DocType = CPCB then
      begin
        NodeData^.IconIndex := 2;
        NodeData^.SchDocIndex := -3;  //Some unusual value for pcbs.
      end
      else
        if NodeData^.DocType = CSCH then
        begin
          NodeData^.IconIndex := 1;
          Inc(SchDocIndex);  //starts at 1
          NodeData^.SchDocIndex := SchDocIndex;
        end
        else
          if NodeData^.DocType = CETC then
          begin
            NodeData^.IconIndex := 7;
            NodeData^.SchDocIndex := -4;  //Some unusual value for unknown file types.
          end
          else
            if NodeData^.DocType = CPRJ then    //this should not happen, but better display it, than simply show an unknown file
            begin
              NodeData^.IconIndex := 0;
              NodeData^.SchDocIndex := -4;  //Some unusual value for unknown file types.
            end;
    end;
  finally
    ListOfDocuments.Free;
    ListOfTypes.Free;
  end;
end;


//Load
procedure TfrmEDAProjectsClickerForm.SetDocNodeTemplatesFromClickerProjectFile(Ini: TClkIniReadonlyFile; ProjectNode: PVirtualNode);
var
  NodeData, PrjNodeData: PNodeDataProjectRec;
  DocNode: PVirtualNode;
  SectionName: string;
  NodeChecked: Boolean;
begin
  DocNode := vstProjects.GetFirstChild(ProjectNode);   //a SCH or PCB file
  if DocNode = nil then
    Exit;

  PrjNodeData := vstProjects.GetNodeData(ProjectNode);

  repeat
    NodeData := vstProjects.GetNodeData(DocNode);
    SectionName := PrjNodeData^.DisplayedFile + '\' + NodeData^.DisplayedFile;

    NodeData^.Template := Ini.ReadString(SectionName, 'DocTemplateFileName', '');
    NodeData^.BeforeAllChildTemplates := Ini.ReadString(SectionName, 'DocBeforeAllChildTemplatesFileName', '');
    NodeData^.AfterAllChildTemplates := Ini.ReadString(SectionName, 'DocAfterAllChildTemplatesFileName', '');

    NodeData^.PageSize := Ini.ReadString(SectionName, 'PageSize', '');
    NodeData^.PageLayout := Ini.ReadString(SectionName, 'PageLayout', '');
    NodeData^.PageScaling := Ini.ReadString(SectionName, 'PageScaling', '');
    NodeData^.UserNotes := Ini.ReadString(SectionName, 'UserNotes', '');
    NodeData^.ListOfEDACustomVars := StringReplace(Ini.ReadString(SectionName, 'EDAFileCustomVars', ''), #4#5, #13#10, [rfReplaceAll]);
    // NodeData^.SchDocIndex := -1;  - do not reset here

    NodeChecked := Ini.ReadBool(SectionName, 'Enabled', True);
    vstProjects.CheckState[DocNode] := CBoolToCheckState[NodeChecked];

    DocNode := DocNode^.NextSibling;
  until DocNode = nil;
end;


procedure TfrmEDAProjectsClickerForm.LoadListOfProjects(AFileName: string);
var
  i: Integer;
  Ini: TClkIniReadonlyFile;
  ProjectFileName, PageSize, PageLayout, PageScaling, Template, BeforeAllChildTemplates, AfterAllChildTemplates, UserNotes, EDAFileCustomVars: string;
  EDAFileCustomVarsRaw: string;
  ProjectNode: PVirtualNode;
  Suffix: string;
begin
  vstProjects.BeginUpdate;
  try
    vstProjects.Clear;

    Ini := DoOnTClkIniReadonlyFileCreate(AFileName);
    try
      chkPlayPrjOnPlayAll.Checked := Ini.ReadBool('Options', 'PlayPrjOnPlayAll', True);
      chkPlayPrjMainTemplateOnPlayAll.Checked := Ini.ReadBool('Options', 'PlayPrjMainTemplateOnPlayAll', False);

      for i := 0 to Ini.ReadInteger('Projects', 'Count', 0) - 1 do
      begin
        Suffix := IntToStr(i);
        PageSize := Ini.ReadString('Projects', 'PageSize_' + Suffix, 'A4');
        PageLayout := Ini.ReadString('Projects', 'PageLayout_' + Suffix, 'Portrait');
        PageScaling := Ini.ReadString('Projects', 'PageScaling_' + Suffix, '1');
        ProjectFileName := Ini.ReadString('Projects', 'Project_' + Suffix, '');
        Template := Ini.ReadString('Projects', 'Template_' + Suffix, '');
        BeforeAllChildTemplates := Ini.ReadString('Projects', 'BeforeAllChildTemplates_' + Suffix, '');
        AfterAllChildTemplates := Ini.ReadString('Projects', 'AfterAllChildTemplates_' + Suffix, '');
        UserNotes := Ini.ReadString('Projects', 'UserNotes_' + Suffix, '');

        EDAFileCustomVarsRaw := Ini.ReadString('Projects', 'EDAFileCustomVars_' + Suffix, '');
        EDAFileCustomVars := FastReplace_45ToReturn(EDAFileCustomVarsRaw);

        ProjectNode := LoadEDAProjectToVst(ProjectFileName, PageSize, PageLayout, PageScaling, Template, BeforeAllChildTemplates, AfterAllChildTemplates, UserNotes, EDAFileCustomVars);
        vstProjects.CheckState[ProjectNode] := CBoolToCheckState[Ini.ReadBool('Projects', 'Enabled_' + Suffix, True)];
        SetDocNodeTemplatesFromClickerProjectFile(Ini, ProjectNode);
      end;
    finally
      Ini.Free;
    end;
  finally
    vstProjects.EndUpdate;
    vstProjects.UpdateScrollBars(True); //not sure if this fixes the issue
    vstProjects.Repaint;
  end;

  Modified := False;
  ToBeUpdated := True;  //trigger
  ToBeUpdated := False;
end;


procedure TfrmEDAProjectsClickerForm.UncheckAllProjects;
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      vstProjects.CheckState[Node] := csUncheckedNormal;
      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.UncheckAll1Click(Sender: TObject);
begin
  UncheckAllProjects;
end;


procedure TfrmEDAProjectsClickerForm.UnselectAll1Click(Sender: TObject);
begin
  vstProjects.ClearSelection;
  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.UpdateClickerProjectFileWithDocNodeTemplate(Ini: TStringList; ProjectNode: PVirtualNode);
var
  NodeData, PrjNodeData: PNodeDataProjectRec;
  DocNode: PVirtualNode;
  SectionName: string;
begin
  DocNode := vstProjects.GetFirstChild(ProjectNode);
  if DocNode = nil then
    Exit;

  PrjNodeData := vstProjects.GetNodeData(ProjectNode);  

  repeat
    NodeData := vstProjects.GetNodeData(DocNode);
    SectionName := PrjNodeData^.DisplayedFile + '\' + NodeData^.DisplayedFile;
    Ini.Add('[' + SectionName + ']');
    
    Ini.Add('DocTemplateFileName' + '=' + NodeData^.Template);
    Ini.Add('DocBeforeAllChildTemplatesFileName' + '=' + NodeData^.BeforeAllChildTemplates);
    Ini.Add('DocAfterAllChildTemplatesFileName' + '=' + NodeData^.AfterAllChildTemplates);

    Ini.Add('PageSize' + '=' + NodeData^.PageSize);
    Ini.Add('PageLayout' + '=' + NodeData^.PageLayout);
    Ini.Add('PageScaling' + '=' + NodeData^.PageScaling);
    Ini.Add('UserNotes' + '=' + NodeData^.UserNotes);
    Ini.Add('EDAFileCustomVars' + '=' + FastReplace_ReturnTo45(NodeData^.ListOfEDACustomVars));

    Ini.Add('Enabled' + '=' + IntToStr(Ord(DocNode^.CheckState = csCheckedNormal)));
    Ini.Add('');

    DocNode := DocNode^.NextSibling;
  until DocNode = nil;
end;


function TfrmEDAProjectsClickerForm.SaveListOfProjects(AFileName: string): Boolean;
var
  Node: PVirtualNode;
  Ini: TStringList;
  NodeData: PNodeDataProjectRec;
  Count: Integer;
  CountStr: string;
begin
  Result := False;

  Node := vstProjects.GetFirst;
  if Node = nil then
    if MessageBox(Handle, 'Save empty file?', PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDNO then
      Exit;

  Ini := TStringList.Create;
  try
    Ini.Add('[Options]');
    Ini.Add('PlayPrjOnPlayAll=' + IntToStr(Ord(chkPlayPrjOnPlayAll.Checked)));
    Ini.Add('PlayPrjMainTemplateOnPlayAll=' + IntToStr(Ord(chkPlayPrjMainTemplateOnPlayAll.Checked)));

    Ini.Add('');
    Ini.Add('[Projects]');

    Node := vstProjects.GetFirst;
    Count := 0;
    repeat
      Inc(Count);
      Node := Node^.NextSibling;
    until Node = nil;

    Ini.Add('Count=' + IntToStr(Count));

    Count := 0;
    Node := vstProjects.GetFirst;
    repeat
      CountStr := IntToStr(Count);
      NodeData := vstProjects.GetNodeData(Node);
      Ini.Add('PageSize_' + CountStr + '=' + NodeData^.PageSize);
      Ini.Add('PageLayout_' + CountStr + '=' + NodeData^.PageLayout);
      Ini.Add('PageScaling_' + CountStr + '=' + NodeData^.PageScaling);
      Ini.Add('Project_' + CountStr + '=' + NodeData^.FilePath);
      Ini.Add('Template_' + CountStr + '=' + NodeData^.Template);
      Ini.Add('BeforeAllChildTemplates_' + CountStr + '=' + NodeData^.BeforeAllChildTemplates);
      Ini.Add('AfterAllChildTemplates_' + CountStr + '=' + NodeData^.AfterAllChildTemplates);
      Ini.Add('UserNotes_' + CountStr + '=' + NodeData^.UserNotes);
      Ini.Add('EDAFileCustomVars_' + CountStr + '=' + FastReplace_ReturnTo45(NodeData^.ListOfEDACustomVars));
      Ini.Add('Enabled_' + CountStr + '=' + IntToStr(Ord(Node^.CheckState = csCheckedNormal)));

      Ini.Add('');

      Inc(Count);
      Node := Node^.NextSibling;
    until Node = nil;

    Node := vstProjects.GetFirst;
    repeat
      UpdateClickerProjectFileWithDocNodeTemplate(Ini, Node);
      Ini.Add('');

      Node := Node^.NextSibling;
    until Node = nil;

    try
      DoOnSaveTextToFile(Ini, AFileName);
      Result := True;
    except
      Result := False;
    end;
  finally
    Ini.Free;
  end;

  Modified := False;
  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.RemoveAllProjects1Click(Sender: TObject);
begin
  if MessageBox(Handle, 'Are you sure you want to remove all the projects from list?', PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDNO then
    Exit;

  vstProjects.Clear;
  vstProjects.Repaint;
  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.SetModified(Value: Boolean);
begin
  if FModified <> Value then
  begin
    FModified := Value;
    bitbtnSaveListOfProjects.Font.Color := clWindowText;

    if FModified then
      bitbtnSaveListOfProjects.Font.Color := clRed;
  end;
  
  lblListOfProjectsStatus.Visible := Value;
end;


procedure TfrmEDAProjectsClickerForm.SetToBeUpdated(Value: Boolean);
begin
  if FToBeUpdated <> Value then
  begin
    FToBeUpdated := Value;
    bitbtnUpdateProject.Font.Color := clWindowText;
    bitbtnUpdateProject.Font.Style := [];

    if FToBeUpdated then
    begin
      bitbtnUpdateProject.Font.Color := clRed;
      bitbtnUpdateProject.Font.Style := [fsBold];
    end;
  end;
end;


procedure TfrmEDAProjectsClickerForm.SetListOfProjectsFileName(Value: string);
begin
  if FListOfProjectsFileName <> Value then
  begin
    FListOfProjectsFileName := Value;
    
    lblLoadedListOfProjects.Caption := ExtractFileName(FListOfProjectsFileName);
    lblLoadedListOfProjects.Hint := FListOfProjectsFileName;
  end;
end;


procedure TfrmEDAProjectsClickerForm.GetCurrentProjectList(AList: TStringList);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
begin
  AList.Clear;
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  repeat
    NodeData := vstProjects.GetNodeData(Node);
    AList.Add(NodeData^.FilePath);

    Node := Node^.NextSibling;
  until Node = nil;
end;


procedure TfrmEDAProjectsClickerForm.grpEDAFileCustomVarsMouseLeave(Sender: TObject);
{$IFnDEF FPC}
  var
    i: Integer;
{$ENDIF}
begin
  {$IFDEF FPC}
    if vallstEDAFileCustomVars.Focused then
      lbeUserNotes.SetFocus; //make vallst remove focus
  {$ELSE}
    for i := 0 to vallstEDAFileCustomVars.ControlCount - 1 do
      if (vallstEDAFileCustomVars.Controls[i] is TInplaceEditList) then    //TInplaceEditList is the internal editbox, used for inputting new values
        if (vallstEDAFileCustomVars.Controls[i] as TInplaceEditList).Focused then
          lbeUserNotes.SetFocus; //make vallst remove focus
  {$ENDIF}
end;


procedure SearchProject(AVst: TVirtualStringTree; AProjectName: string; SearchFullPath: Boolean);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
  ProjectPath: string;
  UpperCaseSearchName: string;
begin
  Node := AVst.GetFirst;
  if Node = nil then
    Exit;

  UpperCaseSearchName := UpperCase(AProjectName); 

  AVst.BeginUpdate;
  try
    repeat
      NodeData := AVst.GetNodeData(Node);

      if SearchFullPath then
        ProjectPath := UpperCase(NodeData^.FilePath)
      else  
        ProjectPath := UpperCase(NodeData^.DisplayedFile);

      AVst.IsVisible[Node] := (AProjectName = '') or (Pos(UpperCaseSearchName, ProjectPath) > 0);

      Node := Node^.NextSibling;
    until Node = nil;
  finally
    AVst.EndUpdate;
  end;
end;


procedure TfrmEDAProjectsClickerForm.lbePageScalingChange(Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.lbePageSizeChange(Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.lbeSearchProjectChange(Sender: TObject);
begin
  SearchProject(vstProjects, lbeSearchProject.Text, chkDisplayFullPath.Checked);
end;



procedure TfrmEDAProjectsClickerForm.lbeUserNotesChange(Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.MenuItem_EDAFileAddCustomVarRowClick(Sender: TObject);
begin
  vallstEDAFileCustomVars.Strings.Add('');
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.MenuItem_EDAFileRemoveCustomVarRowClick(
  Sender: TObject);
begin
  try
    if MessageBox(Handle, PChar('Remove variable?' + #13#10 + vallstEDAFileCustomVars.Strings[vallstEDAFileCustomVars.Selection.Top - 1]), 'Selection', MB_ICONQUESTION + MB_YESNO) = IDNO then
      Exit;

    vallstEDAFileCustomVars.Strings.Delete(vallstEDAFileCustomVars.Selection.Top - 1);
    ToBeUpdated := True;
  except
  end;
end;


procedure TfrmEDAProjectsClickerForm.Addmultipleprojectsfromtextlist1Click(
  Sender: TObject);
var
  NewList, LoadedList: TStringList;
  AOpenDialog: TOpenDialog;
  i, IndexInLoaded: Integer;
  Template: string;
  BeforeAllChildTemplates: string;
  AfterAllChildTemplates: string;
begin
  NewList := TStringList.Create;
  LoadedList := TStringList.Create;
  AOpenDialog := TOpenDialog.Create(nil);
  try
    AOpenDialog.Filter := 'Text List of project files (*.txt)|*.txt|All files (*.*)|*.*';
    AOpenDialog.InitialDir := FTextListOfProjectsDir;

    if not AOpenDialog.Execute then
      Exit;

    GetCurrentProjectList(LoadedList);
    NewList.LoadFromFile(AOpenDialog.FileName);

    for i := NewList.Count - 1 downto 0 do
    begin
      IndexInLoaded := LoadedList.IndexOf(NewList.Strings[i]);
      if IndexInLoaded > -1 then
        NewList.Delete(i);
    end;

    if cmbTemplate.ItemIndex = -1 then
      Template := ''
    else
      Template := cmbTemplate.Items.Strings[cmbTemplate.ItemIndex];

    if cmbBeforeAllChildTemplates.ItemIndex = -1 then
      BeforeAllChildTemplates := ''
    else
      BeforeAllChildTemplates := cmbBeforeAllChildTemplates.Items.Strings[cmbBeforeAllChildTemplates.ItemIndex];

    if cmbAfterAllChildTemplates.ItemIndex = -1 then
      AfterAllChildTemplates := ''
    else
      AfterAllChildTemplates := cmbAfterAllChildTemplates.Items.Strings[cmbAfterAllChildTemplates.ItemIndex];

    for i := 0 to NewList.Count - 1 do
      LoadEDAProjectToVst(NewList.Strings[i],
                          lbePageSize.Text,
                          cmbPageLayout.Items.Strings[cmbPageLayout.ItemIndex],
                          lbePageScaling.Text,
                          Template,
                          BeforeAllChildTemplates,
                          AfterAllChildTemplates,
                          lbeUserNotes.Text,
                          vallstEDAFileCustomVars.Strings.Text);

    vstProjects.Repaint;

    FTextListOfProjectsDir := ExtractFileDir(AOpenDialog.FileName);  
  finally
    NewList.Free;
    LoadedList.Free;
    AOpenDialog.Free;
  end;
end;


function GetProjectFileTypeForDialogFilter: string;
var
  ListOfExtensions: TStringList;
  i: Integer;
  CurrentExt: string;
begin
  Result := '';
  ListOfExtensions := TStringList.Create;
  try
    GetListsOfKnownEDAProjectExtensions(ListOfExtensions);
    for i := 0 to ListOfExtensions.Count - 1 do
    begin
      CurrentExt := ListOfExtensions.Strings[i];
      Result := Result + 'Project files (*' + CurrentExt + ')|*' + CurrentExt + '|';
    end;
  finally
    ListOfExtensions.Free;
  end;

  Result := Result + 'All files (*.*)|*.*';
end;


procedure TfrmEDAProjectsClickerForm.btnAddProjectClick(Sender: TObject);
var
  Template: string;
  BeforeAllChildTemplates: string;
  AfterAllChildTemplates: string;
  LoadedList: TStringList;
  AOpenDialog: TOpenDialog;
begin
  AOpenDialog := TOpenDialog.Create(nil);
  try
    AOpenDialog.Filter := GetProjectFileTypeForDialogFilter;
    AOpenDialog.InitialDir := EDAProjectsDir;
    if not AOpenDialog.Execute then
      Exit;

    LoadedList := TStringList.Create;
    try
      GetCurrentProjectList(LoadedList);
      if LoadedList.IndexOf(AOpenDialog.FileName) > -1 then
      begin
        MessageBox(Handle, 'This project is already in the list.', PChar(Caption), MB_ICONERROR);
        Exit;
      end;

      if cmbTemplate.ItemIndex = -1 then
        Template := ''
      else
        Template := cmbTemplate.Items.Strings[cmbTemplate.ItemIndex];

      if cmbBeforeAllChildTemplates.ItemIndex = -1 then
        BeforeAllChildTemplates := ''
      else
        BeforeAllChildTemplates := cmbBeforeAllChildTemplates.Items.Strings[cmbBeforeAllChildTemplates.ItemIndex];

      if cmbAfterAllChildTemplates.ItemIndex = -1 then
        AfterAllChildTemplates := ''
      else
        AfterAllChildTemplates := cmbAfterAllChildTemplates.Items.Strings[cmbAfterAllChildTemplates.ItemIndex];

      LoadEDAProjectToVst(AOpenDialog.FileName,
                          lbePageSize.Text,
                          cmbPageLayout.Items.Strings[cmbPageLayout.ItemIndex],
                          lbePageScaling.Text,
                          Template,
                          BeforeAllChildTemplates,
                          AfterAllChildTemplates,
                          lbeUserNotes.Text,
                          vallstEDAFileCustomVars.Strings.Text);

      vstProjects.Repaint;
      EDAProjectsDir := ExtractFileDir(AOpenDialog.FileName);
      Modified := True;
    finally
      LoadedList.Free;
    end;
  finally
    AOpenDialog.Free;
  end;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.btnBrowseActionTemplatesDirClick(
  Sender: TObject);
var
  AOpenDialog: TSelectDirectoryDialog;
begin
  AOpenDialog := TSelectDirectoryDialog.Create(nil);
  try
    AOpenDialog.Filter := 'Clicker template files (*.clktmpl)|*.clktmpl|All files (*.*)|*.*';
    AOpenDialog.InitialDir := StringReplace(lbePathToTemplates.Text, '$AppDir$', ParamStr(0), [rfReplaceAll]);

    if AOpenDialog.Execute then
    begin
      lbePathToTemplates.Text := StringReplace(AOpenDialog.FileName, ExtractFileDir(ParamStr(0)), '$AppDir$', [rfReplaceAll]);
      FullTemplatesDir := lbePathToTemplates.Text;
      if FFullTemplatesDir > '' then
        if FFullTemplatesDir[Length(FFullTemplatesDir)] = '\' then
          FFullTemplatesDir := Copy(FFullTemplatesDir, 1, Length(FFullTemplatesDir) - 1);
    end;
  finally
    AOpenDialog.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.btnBrowseEDAPluginPath32Click(Sender: TObject);
var
  AOpenDialog: TOpenDialog;
begin
  AOpenDialog := TOpenDialog.Create(nil);
  try
    AOpenDialog.Filter := 'DLL Files (*.dll)|*.dll|All Files (*.*)|*.*';
    if AOpenDialog.Execute then
      lbeEDAPluginPath32.Text := AOpenDialog.FileName;
  finally
    AOpenDialog.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.btnBrowseEDAPluginPath64Click(
  Sender: TObject);
var
  AOpenDialog: TOpenDialog;
begin
  AOpenDialog := TOpenDialog.Create(nil);
  try
    AOpenDialog.Filter := 'DLL Files (*.dll)|*.dll|All Files (*.*)|*.*';
    if AOpenDialog.Execute then
      lbeEDAPluginPath64.Text := AOpenDialog.FileName;
  finally
    AOpenDialog.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.btnConnectClick(Sender: TObject);
begin
  FConfiguredRemoteAddress := lbeClientModeServerAddress.Text;
  GeneralConnectTimeout := StrToIntDef(lbeConnectTimeout.Text, 1000);

  if FPollForMissingServerFiles <> nil then
  begin
    try
      FPollForMissingServerFiles.Terminate; //terminate existing thread, so a new one can be created
    except
      on E: Exception do
        MessageBox(0, PChar('Exception on stopping client thread. Maybe it''s already done: ' + E.Message), PChar(Application.Title), MB_ICONERROR);
    end;

    Exit;
  end;

  FPollForMissingServerFiles := TPollForMissingServerFiles.Create(True);
  FPollForMissingServerFiles.RemoteAddress := lbeClientModeServerAddress.Text;
  FPollForMissingServerFiles.ConnectTimeout := StrToIntDef(lbeConnectTimeout.Text, 1000);
  FPollForMissingServerFiles.FullTemplatesDir := FFullTemplatesDir;
  FPollForMissingServerFiles.AddListOfAccessibleDirs(memAllowedFileDirsForServer.Lines);
  FPollForMissingServerFiles.AddListOfAccessibleFileExtensions(memAllowedFileExtensionsForServer.Lines);
  FPollForMissingServerFiles.OnBeforeRequestingListOfMissingFiles := HandleOnBeforeRequestingListOfMissingFiles;
  FPollForMissingServerFiles.OnAfterRequestingListOfMissingFiles := HandleOnAfterRequestingListOfMissingFiles;
  FPollForMissingServerFiles.OnFileExists := HandleOnFileExists;
  FPollForMissingServerFiles.OnLogMissingServerFile := HandleLogMissingServerFile;
  FPollForMissingServerFiles.OnLoadMissingFileContent := HandleOnLoadMissingFileContent;
  FPollForMissingServerFiles.Start;

  AddToLog('Started "missing files" monitoring thread.');
  lblClientMode.Caption := 'Client on';
  lblClientMode.Font.Color := clGreen;
end;


procedure TfrmEDAProjectsClickerForm.btnDisconnectClick(Sender: TObject);
var
  tk: QWord;
begin
  if FPollForMissingServerFiles <> nil then
  begin
    AddToLog('Stopping "missing files" monitoring thread for client.');
    FPollForMissingServerFiles.Terminate;

    btnConnect.Enabled := False;
    try
      tk := GetTickCount64;
      repeat
        Application.ProcessMessages;
        Sleep(10);

        if FPollForMissingServerFiles.Done then
        begin
          AddToLog('Monitoring thread terminated.');
          Break;
        end;

        if GetTickCount64 - tk > 1500 then
        begin
          AddToLog('Timeout waiting for monitoring thread to terminate. The thread is still running (probably waiting).');
          Break;
        end;
      until False;
    finally
      btnConnect.Enabled := True;
    end;

    FPollForMissingServerFiles := nil;
    lblClientMode.Caption := 'Client off';
    lblClientMode.Font.Color := clOlive;
  end;
end;


procedure TfrmEDAProjectsClickerForm.SetEDAPluginErrorMessage(AMsg: string; AColor: TColor);
begin
  lblEDAPluginErrorMessage.Caption := AMsg;
  lblEDAPluginErrorMessage.Font.Color := AColor;
  lblEDAPluginErrorMessage.Repaint; //should not be needed, but simetimes, labels are not properly repainted
end;


procedure TfrmEDAProjectsClickerForm.btnUpdateProjectClick(Sender: TObject);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'No project selected. Please select a project to be updated.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  if vstProjects.SelectedCount > 1 then
    if MessageBox(Handle, 'Update all selected items with the same settings/vars?', PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDNO then
      Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      if vstProjects.Selected[Node] then
      begin
        NodeData := vstProjects.GetNodeData(Node);
        NodeData^.PageSize := lbePageSize.Text;

        if cmbPageLayout.ItemIndex > -1 then
          NodeData^.PageLayout := cmbPageLayout.Items.Strings[cmbPageLayout.ItemIndex];
          
        NodeData^.PageScaling := lbePageScaling.Text;
        NodeData^.UserNotes := lbeUserNotes.Text;
        NodeData^.ListOfEDACustomVars := vallstEDAFileCustomVars.Strings.Text;

        if cmbTemplate.ItemIndex = -1 then
          NodeData^.Template := ''
        else
          NodeData^.Template := cmbTemplate.Items.Strings[cmbTemplate.ItemIndex];

        if cmbBeforeAllChildTemplates.ItemIndex = -1 then
          NodeData^.BeforeAllChildTemplates := ''
        else
          NodeData^.BeforeAllChildTemplates := cmbBeforeAllChildTemplates.Items.Strings[cmbBeforeAllChildTemplates.ItemIndex];

        if cmbAfterAllChildTemplates.ItemIndex = -1 then
          NodeData^.AfterAllChildTemplates := ''
        else
          NodeData^.AfterAllChildTemplates := cmbAfterAllChildTemplates.Items.Strings[cmbAfterAllChildTemplates.ItemIndex];
      end;

      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;
  
  vstProjects.Repaint;
  Modified := True;
  ToBeUpdated := False;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.CheckAll1Click(Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      vstProjects.CheckState[Node] := csCheckedNormal;
      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.CheckAllProjectDocumentsForSelectedProjectsClick(
  Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      if vstProjects.GetNodeLevel(Node) > 0 then //document  (sch or pcb)
        if vstProjects.Selected[Node^.Parent] then
          vstProjects.CheckState[Node] := csCheckedNormal;

      Node := vstProjects.GetNext(Node);    
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.CheckAllProjects(IsSelectedOnly: Boolean);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      if (IsSelectedOnly and vstProjects.Selected[Node]) or not IsSelectedOnly then
        vstProjects.CheckState[Node] := csCheckedNormal;

      Node := Node^.NextSibling;
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.CheckAllProjects1Click(Sender: TObject);
begin
  UncheckAllProjects;
  CheckAllProjects(False);
end;


procedure TfrmEDAProjectsClickerForm.CheckAllSelectedProjects1Click(Sender: TObject);
begin
  UncheckAllProjects;
  CheckAllProjects(True);
end;


procedure TfrmEDAProjectsClickerForm.RemoveSelectedEDAFiles;
var
  Node: PVirtualNode;
  s: string;
begin
  Node := vstProjects.GetFirstSelected;

  if Node = nil then
  begin
    MessageBox(Handle, 'No project selected. Please select a project to be removed.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  if vstProjects.SelectedCount > 1 then
    s := 's'
  else
    s := '';

  if MessageBox(Handle, PChar('Are you sure you want to remove the selected item' + s + '?'), PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDNO then
    Exit;

  Node := vstProjects.GetLast;
  vstProjects.BeginUpdate;
  try
    repeat
      if vstProjects.Selected[Node] then
        vstProjects.DeleteNode(Node);

      Node := vstProjects.GetPrevious(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  vstProjects.Repaint;
  Modified := True;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.btnRemoveProjectClick(Sender: TObject);
begin
  RemoveSelectedEDAFiles;
end;


procedure TfrmEDAProjectsClickerForm.MenuItem_SetDefaultClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FPopUpComboBox <> nil then
  begin
    Idx := FPopUpComboBox.Items.IndexOf(LastClickedComboBoxDefaultString);
    if Idx <> -1 then
      FPopUpComboBox.ItemIndex := Idx;
  end;
end;


procedure TfrmEDAProjectsClickerForm.tmrDisplayMissingFilesRequestsTimer(
  Sender: TObject);
const
  CRequestColor: array[Boolean] of TColor = (clGreen, clLime);
  CRequestFontColor: array[Boolean] of TColor = (clWhite, clBlack);
begin
  if FProcessingMissingFilesRequestByClient then
  begin
    FProcessingMissingFilesRequestByClient := False;
    pnlMissingFilesRequest.Color := CRequestColor[True];
    pnlMissingFilesRequest.Font.Color := CRequestFontColor[True];
  end
  else
  begin   //displaying "False" on next timer iteration, to allow it to be visible
    pnlMissingFilesRequest.Color := CRequestColor[False];
    pnlMissingFilesRequest.Font.Color := CRequestFontColor[False];
  end;
end;


procedure TfrmEDAProjectsClickerForm.btnLoadListOfProjectsClick(Sender: TObject);
var
  Fnm: string;
begin
  DoOnSetEDAClickerFileOpenDialogInitialDir(ClickerProjectsDir);
  if not DoOnEDAClickerFileOpenDialogExecute then
    Exit;

  if vstProjects.RootNodeCount > 0 then
    if MessageBox(Handle, 'Discard existing list of projects?', PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDNO then
      Exit;

  Fnm := DoOnGetEDAClickerFileOpenDialogFileName;
  LoadListOfProjects(Fnm);
  ClickerProjectsDir := ExtractFileDir(Fnm);
  ListOfProjectsFileName := Fnm;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.btnNewListOfProjectsClick(Sender: TObject);
begin
  if FModified  then
    if MessageBox(Handle, 'The current list of projects is modified. Save?', PChar(Caption), MB_ICONQUESTION + MB_YESNO) = IDYES then
    begin
      if FListOfProjectsFileName = '' then
        SaveListOfProjectsWithDialog
      else
        if not SaveListOfProjects(FListOfProjectsFileName) then
          Exit;
    end;

  ListOfProjectsFileName := '';

  chkPlayPrjOnPlayAll.Checked := True;
  chkPlayPrjMainTemplateOnPlayAll.Checked := False;

  vstProjects.Clear;
  vstProjects.Repaint;
  Modified := False;
  ToBeUpdated := False;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.SaveListOfProjectsAs1Click(Sender: TObject);
begin
  SaveListOfProjectsWithDialog;
end;


procedure TfrmEDAProjectsClickerForm.btnSaveListOfProjectsClick(Sender: TObject);
begin
  if DoOnFileExists(FListOfProjectsFileName) then
    SaveListOfProjects(FListOfProjectsFileName)
  else
    SaveListOfProjectsWithDialog;
end;


procedure TfrmEDAProjectsClickerForm.SaveListOfProjectsWithDialog;
var
  Fnm: string;
begin
  DoOnSetEDAClickerFileSaveDialogInitialDir(ClickerProjectsDir);
  if not DoOnEDAClickerFileSaveDialogExecute then
    Exit;

  Fnm := DoOnGetEDAClickerFileSaveDialogFileName;

  if UpperCase(ExtractFileExt(Fnm)) <> '.EDACLK' then
    Fnm := Fnm + '.edaclk';

  if DoOnFileExists(Fnm) then
    if MessageBox(Handle, 'File already exists. Overwrite?', PChar(Caption), MB_ICONWARNING + MB_YESNO) = IDNO then
      Exit;

  SaveListOfProjects(Fnm);
  ClickerProjectsDir := ExtractFileDir(Fnm);
  ListOfProjectsFileName := Fnm;
end;


procedure TfrmEDAProjectsClickerForm.btnSetReplacementsClick(Sender: TObject);
var
  OverridenValues: TStringList;
  Node: PVirtualNode;
  NodeData, PrjNodeData: PNodeDataProjectRec;
  i: Integer;
  Key, Value, KeyValue: string;
  SetVarOptions: TClkSetVarOptions;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'No project or document selected. Please select one.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  NodeData := vstProjects.GetNodeData(Node);
  PrjNodeData := NodeData;

  GetTemplateNameFromNode(vstProjects, Node, PrjNodeData, ttMain);

  OverridenValues := TStringList.Create;
  try
    UpdateOverridenValuesForActionVars(OverridenValues, NodeData, PrjNodeData, NodeData^.DocType);

    SetVarOptions.ListOfVarNames := '';
    SetVarOptions.ListOfVarValues := '';
    SetVarOptions.ListOfVarEvalBefore := '';
    for i := 0 to OverridenValues.Count - 1 do
    begin
      KeyValue := OverridenValues.Strings[i];
      Key := Copy(KeyValue, 1, Pos('=', KeyValue) - 1);
      Value := Copy(KeyValue, Pos('=', KeyValue) + 1, MaxInt); 

      SetVarOptions.ListOfVarNames := SetVarOptions.ListOfVarNames + Key + #13#10;
      SetVarOptions.ListOfVarValues := SetVarOptions.ListOfVarValues + Value + #13#10;
      SetVarOptions.ListOfVarEvalBefore := SetVarOptions.ListOfVarEvalBefore + '0'#13#10;
    end;

    ExecuteSetVarAction(DoOnGetConnectionAddress, SetVarOptions);
  finally
    OverridenValues.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.chkDisplayFullPathClick(Sender: TObject);
begin
  if vstProjects.RootNodeCount > 0 then
    vstProjects.Repaint;
end;


procedure TfrmEDAProjectsClickerForm.chkPlayPrjMainTemplateOnPlayAllClick(
  Sender: TObject);
begin
  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.chkPlayPrjOnPlayAllClick(Sender: TObject);
begin
  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.chkStayOnTopClick(Sender: TObject);
begin
  if chkStayOnTop.Checked then
    FormStyle := fsStayOnTop
  else
    FormStyle := fsNormal;
end;


procedure TfrmEDAProjectsClickerForm.SetCmbTemplateContent(AComboBox: TComboBox);
var
  AvailableTemplates: TStringList;
  SelectedTemplate: string;
begin
  AvailableTemplates := TStringList.Create;
  try
    DoOnGetListOfFilesFromDir(FullTemplatesDir, '.clktmpl', AvailableTemplates);

    SelectedTemplate := '';
    if AComboBox.ItemIndex > -1 then
      SelectedTemplate := AComboBox.Items.Strings[AComboBox.ItemIndex];

    AComboBox.Items.Clear;
    AComboBox.Items.AddStrings(AvailableTemplates);  //load the existing files

    if AvailableTemplates.IndexOf(SelectedTemplate) > -1 then //currently selected template still exists on disk
      AComboBox.ItemIndex := AComboBox.Items.IndexOf(SelectedTemplate);       //set the ItemIndex to the old selected file
  finally
    AvailableTemplates.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.cmbAfterAllChildTemplatesChange(
  Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.cmbBeforeAllChildTemplatesChange(
  Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.cmbPageLayoutChange(Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.cmbTemplateChange(Sender: TObject);
begin
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.cmbTemplateCloseUp(Sender: TObject);
var
  AComboBox: TComboBox;
begin
  AComboBox := Sender as TComboBox;
  AComboBox.Width := AComboBox.Width - 100;
end;


procedure TfrmEDAProjectsClickerForm.cmbTemplateDropDown(Sender: TObject);
var
  AComboBox: TComboBox;
begin
  AComboBox := Sender as TComboBox;
  AComboBox.DropDownCount := 20;
  AComboBox.Width := AComboBox.Width + 100;
  SetCmbTemplateContent(AComboBox);
end;


procedure TfrmEDAProjectsClickerForm.cmbTemplateMouseEnter(Sender: TObject);
begin
  FPopUpComboBox := Sender as TComboBox;
end;


procedure TfrmEDAProjectsClickerForm.SetAllNodesToExpandedOrCollapsed(IsExpanded: Boolean);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;
  if Node = nil then
    Exit;

  vstProjects.BeginUpdate;
  try
    repeat
      vstProjects.Expanded[Node] := IsExpanded;
      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;
end;


procedure TfrmEDAProjectsClickerForm.CollapseAll1Click(Sender: TObject);
begin
  SetAllNodesToExpandedOrCollapsed(False);
end;


procedure TfrmEDAProjectsClickerForm.CopyDirectoryPathToClipboard1Click(
  Sender: TObject);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
  ClipContent: string;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'Please select an item.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  ClipContent := '';
  try
    vstProjects.BeginUpdate;
    repeat
      if vstProjects.Selected[Node] then
      begin
        NodeData := vstProjects.GetNodeData(Node);
        ClipContent := ClipContent + ExtractFileDir(NodeData^.FilePath) + #13#10;
      end;

      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  if ClipContent > '' then
    Delete(ClipContent, Length(ClipContent) - 1, 2);
    
  Clipboard.AsText := ClipContent;
end;


procedure TfrmEDAProjectsClickerForm.CopyFullPathToClipboard1Click(Sender: TObject);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
  ClipContent: string;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'Please select an item.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  ClipContent := '';
  try
    vstProjects.BeginUpdate;
    repeat
      if vstProjects.Selected[Node] then
      begin
        NodeData := vstProjects.GetNodeData(Node);
        ClipContent := ClipContent + NodeData^.FilePath + #13#10;
      end;

      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  if ClipContent > '' then
    Delete(ClipContent, Length(ClipContent) - 1, 2);
    
  Clipboard.AsText := ClipContent;
end;


procedure TfrmEDAProjectsClickerForm.Displayprojectsinbold1Click(Sender: TObject);
begin
  vstProjects.Repaint;
end;


procedure TfrmEDAProjectsClickerForm.ExpandAll1Click(Sender: TObject);
begin
  SetAllNodesToExpandedOrCollapsed(True);
end;


procedure TfrmEDAProjectsClickerForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  //
end;


procedure TfrmEDAProjectsClickerForm.CreateRemainingUIComponents;
var
  NewColum: TVirtualTreeColumn;
begin
  vstProjects := TVirtualStringTree.Create(Self);
  vstProjects.Parent := TabSheetProjects;

  vstProjects.Left := 3;
  vstProjects.Top := 7;
  vstProjects.Width := 653;
  vstProjects.Height := 338;
  vstProjects.Anchors := [akLeft, akTop, akRight, akBottom];
  vstProjects.CheckImageKind := ckXP;
  vstProjects.Header.AutoSizeIndex := -1;
  vstProjects.Header.DefaultHeight := 21;
  vstProjects.Header.Font.Charset := DEFAULT_CHARSET;
  vstProjects.Header.Font.Color := clWindowText;
  vstProjects.Header.Font.Height := -11;
  vstProjects.Header.Font.Name := 'Tahoma';
  vstProjects.Header.Font.Style := [];
  vstProjects.Header.Height := 21;
  vstProjects.Header.Images := imgLstProjectsHeader;
  vstProjects.Header.Options := [hoColumnResize, hoDblClickResize, hoDrag, hoShowImages, hoShowSortGlyphs, hoVisible];
  vstProjects.Header.Style := hsFlatButtons;
  vstProjects.PopupMenu := pmVSTProjects;
  vstProjects.StateImages := imglstProjects;
  vstProjects.TabOrder := 0;
  vstProjects.TreeOptions.AutoOptions := [toAutoDropExpand, toAutoScrollOnExpand, toAutoTristateTracking, toAutoDeleteMovedNodes, toDisableAutoscrollOnFocus, toDisableAutoscrollOnEdit];
  vstProjects.TreeOptions.MiscOptions := [toAcceptOLEDrop, toCheckSupport, toFullRepaintOnResize, toInitOnSave, toToggleOnDblClick, toWheelPanning];
  vstProjects.TreeOptions.PaintOptions := [toShowButtons, toShowDropmark, toShowRoot, toThemeAware, toUseBlendedImages];
  vstProjects.TreeOptions.SelectionOptions := [toFullRowSelect, toMiddleClickSelect, toMultiSelect, toRightClickSelect];
  vstProjects.OnChecked := vstProjectsChecked;
  vstProjects.OnClick := vstProjectsClick;
  vstProjects.OnGetText := vstProjectsGetText;
  vstProjects.OnPaintText := vstProjectsPaintText;
  vstProjects.OnGetImageIndex := vstProjectsGetImageIndex;
  vstProjects.OnGetImageIndexEx := vstProjectsGetImageIndexEx;
  vstProjects.OnInitNode := vstProjectsInitNode;
  vstProjects.OnKeyUp := vstProjectsKeyUp;
  vstProjects.OnMouseDown := vstProjectsMouseDown;
  vstProjects.OnMouseUp := vstProjectsMouseUp;

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := 0;
  NewColum.MinWidth := 250;
  NewColum.Position := 0;
  NewColum.Width := 250;
  NewColum.Text := 'Project / Schematic / PCB';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := 1;
  NewColum.MinWidth := 270;
  NewColum.Position := 1;
  NewColum.Width := 270;
  NewColum.Text := 'Actions Templates: [Main / BeforeAll / AfterAll]';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 120;
  NewColum.Position := 2;
  NewColum.Width := 120;
  NewColum.Text := 'Status';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 70;
  NewColum.Position := 3;
  NewColum.Width := 70;
  NewColum.Text := 'Page Size';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 80;
  NewColum.Position := 4;
  NewColum.Width := 80;
  NewColum.Text := 'Page Layout';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 80;
  NewColum.Position := 5;
  NewColum.Width := 80;
  NewColum.Text := 'Page Scaling';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 50;
  NewColum.Position := 6;
  NewColum.Width := 50;
  NewColum.Text := 'Index';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 200;
  NewColum.Position := 7;
  NewColum.Width := 200;
  NewColum.Text := 'User Notes';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 200;
  NewColum.Position := 8;
  NewColum.Width := 200;
  NewColum.Text := 'Custom Variables';

  NewColum := vstProjects.Header.Columns.Add;
  NewColum.ImageIndex := -1;
  NewColum.MinWidth := 70;
  NewColum.Position := 9;
  NewColum.Width := 70;
  NewColum.Text := 'Doc Type';
end;


procedure TfrmEDAProjectsClickerForm.FormCreate(Sender: TObject);
begin
  CreateRemainingUIComponents;

  FOnFileExists := nil;
  FOnTClkIniReadonlyFileCreate := nil;
  FOnSaveTextToFile := nil;
  FOnSetEDAClickerFileOpenDialogInitialDir := nil;
  FOnEDAClickerFileOpenDialogExecute := nil;
  FOnGetEDAClickerFileOpenDialogFileName := nil;
  FOnSetEDAClickerFileSaveDialogInitialDir := nil;
  FOnEDAClickerFileSaveDialogExecute := nil;
  FOnGetEDAClickerFileSaveDialogFileName := nil;
  FOnGetListOfFilesFromDir := nil;
  FOnSetStopAllActionsOnDemand := nil;
  FOnGetConnectionAddress := nil;
  FOnLoadMissingFileContent := nil;

  FProcessingMissingFilesRequestByClient := False;
  FPollForMissingServerFiles := nil;
  FConfiguredRemoteAddress := 'unconfigured address';

  FStopAllActionsOnDemand := False;
  FFullTemplatesDir := 'not set'; //

  FDefaultTemplate_Main := '';
  FDefaultTemplate_Before := '';
  FDefaultTemplate_After := '';
  LastClickedComboBoxDefaultString := '';

  PageControlMain.ActivePageIndex := 0;
  FListOfProjectsFileName := '';
  btnSetReplacements.Hint := 'Updates $PageSize$, $PageLayout$, $PageScaling$, $EDADocType$, $FullPathProjectDir$, $FileName$, $PojectName$' + #13#10 +
                             'and their derivatives replacements to current selected EDA file from the "Projects" tab.'+ #13#10 +
                             'This can be useful when playing a template without playing the actions of the project.' + #13#10 +
                             'Uses main template';

  spdbtnExtraAddProject.Hint := 'Adding multiple projects from text list, requires full paths to EDA project files.' + #13#10 +
                                'Those which are already loaded, are ignored (to avoid duplicates).' + #13#10 +
                                'The currently selected templates are applied to EDA project files.' + #13#10 +
                                'SCH/PCB files will have empty templates.';

  vstProjects.NodeDataSize := SizeOf(TNodeDataProjectRec);

  //tmrStartup.Enabled := True; //commented here, to allow loading from main window
end;


procedure TfrmEDAProjectsClickerForm.FormDestroy(Sender: TObject);
var
  tk: QWord;
begin
  try
    DonePlugin;
  except
    on E: Exception do
      MessageBox(Handle, PChar('Error terminating EDA plugin:' + #13#10 + E.Message), PChar(Caption), MB_ICONERROR);
  end;

  try
    UnloadPlugin;
  except
    on E: Exception do
      MessageBox(Handle, PChar('Error unloading EDA plugin:' + #13#10 + E.Message), PChar(Caption), MB_ICONERROR);
  end;

  if FPollForMissingServerFiles <> nil then
  begin
    FPollForMissingServerFiles.Terminate;

    tk := GetTickCount64;
    repeat
      Application.ProcessMessages;
      Sleep(10);

      if FPollForMissingServerFiles.Done then
        Break;

      if GetTickCount64 - tk > 1500 then
        Break;
    until False;

    FPollForMissingServerFiles := nil;
  end;
end;


procedure TfrmEDAProjectsClickerForm.AddToLog(s: string);
begin
  memLog.Lines.Add(DateTimeToStr(Now) + '  ' + s);
end;


procedure TfrmEDAProjectsClickerForm.SelectAll1Click(Sender: TObject);
begin
  vstProjects.SelectAll(False);
  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.SelectAllProjectDocuments1Click(
  Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;

  if Node = nil then
    Exit;

  vstProjects.ClearSelection;  

  try
    vstProjects.BeginUpdate;
    repeat
      if vstProjects.GetNodeLevel(Node) > 0 then
        vstProjects.Selected[Node] := True;
        
      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.SelectAllProjects1Click(Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirst;

  if Node = nil then
    Exit;

  vstProjects.ClearSelection;

  try
    vstProjects.BeginUpdate;
    repeat
      vstProjects.Selected[Node] := True;
      Node := Node^.NextSibling;
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;

  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.Selectnone1Click(Sender: TObject);
begin
  FPopUpComboBox.ItemIndex := -1;
  FPopUpComboBox := nil; //reset, to avoid dangling pointers
  ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.spdbtnExtraAddProjectClick(Sender: TObject);
var
  tp: TPoint;
begin
  GetCursorPos(tp);
  pmExtraAddProjects.Popup(tp.X, tp.Y);
end;


procedure TfrmEDAProjectsClickerForm.spdbtnExtraPlayAllFilesClick(Sender: TObject);
var
  tp: TPoint;
begin
  GetCursorPos(tp);
  pmExtraPlayAllFiles.Popup(tp.X, tp.Y);
end;


procedure TfrmEDAProjectsClickerForm.spdbtnExtraPlaySelectedFileClick(
  Sender: TObject);
var
  tp: TPoint;
begin
  GetCursorPos(tp);
  pmExtraPlaySelectedFile.Popup(tp.X, tp.Y);
end;


procedure TfrmEDAProjectsClickerForm.spdbtnExtraRemoveProjectClick(Sender: TObject);
var
  tp: TPoint;
begin
  GetCursorPos(tp);
  pmExtraRemoveProjects.Popup(tp.X, tp.Y);
end;


procedure TfrmEDAProjectsClickerForm.spdbtnExtraSaveListOfProjectsClick(
  Sender: TObject);
var
  tp: TPoint;
begin
  GetCursorPos(tp);
  pmExtraSaveListOfProjects.Popup(tp.X, tp.Y);
end;


function TfrmEDAProjectsClickerForm.ExecuteTemplateOnUIClicker(ATemplatePath, AFileLocation, ACustomVarsAndValues: string; AIsDebugging: Boolean): string;
var
  CallTemplateOptions: TClkCallTemplateOptions;
begin
  CallTemplateOptions.TemplateFileName := ATemplatePath;
  CallTemplateOptions.ListOfCustomVarsAndValues := ACustomVarsAndValues;
  CallTemplateOptions.EvaluateBeforeCalling := False;
  CallTemplateOptions.CallTemplateLoop.Enabled := False;
  CallTemplateOptions.CallTemplateLoop.Direction := ldInc;
  CallTemplateOptions.CallTemplateLoop.EvalBreakPosition := lebpAfterContent;

  //This debugging setup requires UIClicker to be run on the same machine as EDAUIClicker, because EDAUIClicker does not have any debugger implementation.
  Result := FastReplace_87ToReturn(ExecuteCallTemplateAction(DoOnGetConnectionAddress, CallTemplateOptions, AIsDebugging, AIsDebugging, AFileLocation));
end;                                                                                                        //Using AIsDebugging to control the debugger type, which is always local debugger for this application.


procedure TfrmEDAProjectsClickerForm.PlayNode(Node: PVirtualNode; AIsDebugging: Boolean; ATemplateType: TTemplateType);
var
  OverridenValues, ResultedVars: TStringList;
  TemplateFileName, ExecResult: string;
  NodeData, PrjNodeData: PNodeDataProjectRec;
  ExpandedFullTemplatesDir: string;
begin
  NodeData := vstProjects.GetNodeData(Node);

  TemplateFileName := GetTemplateNameFromNode(vstProjects, Node, PrjNodeData, ATemplateType);
  if TemplateFileName = '' then
  begin
    MessageBox(Handle, PChar('The current selected item does not have a template. Please update it first.' + #13#10 + 'EDA file: ' + NodeData^.DisplayedFile), PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  ExpandedFullTemplatesDir := StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]);
  TemplateFileName := ExpandedFullTemplatesDir + '\' + TemplateFileName;

  OverridenValues := TStringList.Create;
  ResultedVars := TStringList.Create;
  try
    ResultedVars.AddStrings(memVariables.Lines);

    UpdateOverridenValuesForActionVars(OverridenValues, NodeData, PrjNodeData, NodeData^.DocType);   /////////////////// maybe send OverridenValues with SetVar
    ExecResult := ExecuteTemplateOnUIClicker(TemplateFileName, CREParam_FileLocation_ValueMem, OverridenValues.Text, AIsDebugging);

    NodeData^.ExecStatus.Status := asInProgress;
    vstProjects.RepaintNode(Node);

    ResultedVars.Text := ExecResult;

    if (ResultedVars.Values['$ExecAction_Err$'] = '') or (ResultedVars.Values['$ExecAction_Err$'] = 'Action condition is false.') then
      NodeData^.ExecStatus.Status := asSuccessful               //FileName is ignored in PlayAllActions
    else
      NodeData^.ExecStatus.Status := asFailed;

    NodeData^.ExecStatus.AllVars := ExecResult;

    if NodeData <> PrjNodeData then
      if NodeData^.ExecStatus.Status = asFailed then
      begin
        PrjNodeData^.ExecStatus.Status := asFailed;
        vstProjects.RepaintNode(Node^.Parent); 
      end;

    vstProjects.RepaintNode(Node);
  finally
    OverridenValues.Free;
    ResultedVars.Free;
  end;
end;


procedure TfrmEDAProjectsClickerForm.PlayAllStartingAtNode(Node: PVirtualNode; IsDebugging: Boolean);
var
  NodeData, PrjNodeData: PNodeDataProjectRec;
  StartNode, OldNode, LastNode, NextNode: PVirtualNode;
  ListOfProjectsWithNoTemplate: TStringList;
  TemplateFileName: string;
  ExpandedFullTemplatesDir: string;
begin
  ExpandedFullTemplatesDir := StringReplace(FFullTemplatesDir, '$AppDir$', ExtractFileDir(ParamStr(0)), [rfReplaceAll]);
  StartNode := Node;
  ListOfProjectsWithNoTemplate := TStringList.Create;
  try
    LastNode := vstProjects.GetLast;
    repeat
      NodeData := vstProjects.GetNodeData(Node);

      NodeData^.ExecStatus.Status := asNotStarted;
      NodeData^.ExecStatus.AllVars := '';

      TemplateFileName := GetTemplateNameFromNode(vstProjects, Node, PrjNodeData, ttMain); //only main template is verified for now

      TemplateFileName := ExpandedFullTemplatesDir + '\' + TemplateFileName;

      if not DoOnFileExists(TemplateFileName) then
        ListOfProjectsWithNoTemplate.Add(NodeData^.FilePath);

      OldNode := Node;
      Node := vstProjects.GetNext(Node);
    until OldNode = LastNode;

    vstProjects.Repaint; //to update all nodes to NotStarted

    if ListOfProjectsWithNoTemplate.Count > 0 then
    begin
      MessageBox(Handle, PChar('The following projects do not have a valid actions template. If Schematic and PCB files do not have a template, they use the one from their project. Please update them: ' + #13#10#13#10 + ListOfProjectsWithNoTemplate.Text + #13#10#13#10 + 'Also, make sure the path to action templates directory, is properly set.'), PChar(Caption), MB_ICONINFORMATION);
      Exit;
    end;
  finally
    ListOfProjectsWithNoTemplate.Free;
  end;

  vstProjects.ClearSelection;
  vstProjects.Repaint;

  vstProjects.TreeOptions.SelectionOptions := vstProjects.TreeOptions.SelectionOptions - [toMultiSelect];
  try
    Node := StartNode;
    repeat
      vstProjects.Selected[Node] := True;

      if FStopAllActionsOnDemand then
      begin
        FStopAllActionsOnDemand := False;   //has to be reset somewhere, not only by restarting on "play all"
        DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
        Exit;
      end;

      vstProjects.ScrollIntoView(Node, True);

      NodeData := vstProjects.GetNodeData(Node);
      UpdateControlsFromNodeData(NodeData);

      if NodeData^.DocType <> CPRJ then
      begin
        if Node^.CheckState = csCheckedNormal then
          PlayNode(Node, IsDebugging, ttMain);
      end
      else
      begin  //project node
        if chkPlayPrjOnPlayAll.Checked then
          if Node^.CheckState = csCheckedNormal then
          begin
            if chkPlayPrjMainTemplateOnPlayAll.Checked then
              PlayNode(Node, IsDebugging, ttMain);

            PlayNode(Node, IsDebugging, ttBeforeAll);
          end;
      end;

      OldNode := Node;
      NextNode := vstProjects.GetNext(Node);

      if ((NextNode = nil) or (vstProjects.GetNodeLevel(NextNode) = 0)) and (vstProjects.GetNodeLevel(Node) > 0) then  //next node is a project node, so Node should be the last SCH/PCB from this one
      begin   //last child
        if chkPlayPrjOnPlayAll.Checked then
        begin
          if Node^.Parent^.CheckState = csCheckedNormal then
            PlayNode(Node^.Parent, IsDebugging, ttAfterAll);
          //do not update any parent here
        end
        else
        begin
          if Node^.Parent <> nil then  //should always be the case, because GetNodeLevel(Node) > 0
          begin
            PrjNodeData := vstProjects.GetNodeData(Node^.Parent);
            if PrjNodeData^.ExecStatus.Status = asNotStarted then
              PrjNodeData^.ExecStatus.Status := asSuccessful;      //PlayNode sets it to failed if that's the case, so it can only be asInProgress at this point

            vstProjects.RepaintNode(Node^.Parent);
          end;
        end;  
      end;

      Node := NextNode
    until OldNode = LastNode;
  finally
    vstProjects.TreeOptions.SelectionOptions := vstProjects.TreeOptions.SelectionOptions + [toMultiSelect];
  end;

  FStopAllActionsOnDemand := False;   //has to be reset here again, in case the user presses the stop button on last item
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
end;


procedure TfrmEDAProjectsClickerForm.PlaySelectedFileInDebuggingModeClick(Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'No project or document is selected to play. Please select one.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;

  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    PlayNode(Node, True, ttMain);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
    FStopAllActionsOnDemand := False; //reset here as well, because PlayNode can be interrupted
    DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  end;
end;


procedure TfrmEDAProjectsClickerForm.spdbtnPlaySelectedFileClick(Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := vstProjects.GetFirstSelected;
  if Node = nil then
  begin
    MessageBox(Handle, 'No project or document is selected to play. Please select one.', PChar(Caption), MB_ICONINFORMATION);
    Exit;
  end;

  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;

  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    PlayNode(Node, False, ttMain);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
    FStopAllActionsOnDemand := False; //reset here as well, because PlayNode can be interrupted
    DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  end;
end;


procedure TfrmEDAProjectsClickerForm.spdbtnPlayAllFilesClick(Sender: TObject);
var
  Node: PVirtualNode;
begin
  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;
  PlayAllFilesInDebuggingMode.Enabled := False;
  PlayAllFilesbelowselectedindebuggingmode1.Enabled := False;

  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    Node := vstProjects.GetFirst;
    if Node = nil then
      Exit;

    PlayAllStartingAtNode(Node, False);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
    PlayAllFilesInDebuggingMode.Enabled := True;
    PlayAllFilesbelowselectedindebuggingmode1.Enabled := True;
  end;
end;


procedure TfrmEDAProjectsClickerForm.PlayAllFilesbelowselectedincludingselected1Click(
  Sender: TObject);
var
  Node: PVirtualNode;
begin
  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;
  
  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    Node := vstProjects.GetFirstSelected;   //difference from spdbtnPlayAllFilesClick
    if Node = nil then
    begin
      MessageBox(Handle, 'Please select a project to start with.', PChar(Caption), MB_ICONINFORMATION);
      Exit;
    end;

    PlayAllStartingAtNode(Node, False);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
  end;
end;


procedure TfrmEDAProjectsClickerForm.PlayAllFilesbelowselectedindebuggingmode1Click(
  Sender: TObject);
var
  Node: PVirtualNode;
begin
  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;
  
  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    Node := vstProjects.GetFirstSelected;   //difference from spdbtnPlayAllFilesClick
    if Node = nil then
    begin
      MessageBox(Handle, 'Please select a project to start with.', PChar(Caption), MB_ICONINFORMATION);
      Exit;
    end;

    PlayAllStartingAtNode(Node, True);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
  end;
end;


procedure TfrmEDAProjectsClickerForm.PlayAllFilesInDebuggingModeClick(Sender: TObject);
var
  Node: PVirtualNode;
begin
  spdbtnStopPlaying.Enabled := True;
  spdbtnPlaySelectedFile.Enabled := False;
  spdbtnPlayAllFiles.Enabled := False;
  PlayAllFilesbelowselectedincludingselected1.Enabled := False;

  FStopAllActionsOnDemand := False;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
  try
    Node := vstProjects.GetFirst;
    if Node = nil then
      Exit;

    PlayAllStartingAtNode(Node, True);
  finally
    spdbtnStopPlaying.Enabled := False;
    spdbtnPlaySelectedFile.Enabled := True;
    spdbtnPlayAllFiles.Enabled := True;
    PlayAllFilesbelowselectedincludingselected1.Enabled := True;
  end;
end;


procedure TfrmEDAProjectsClickerForm.spdbtnStopPlayingClick(Sender: TObject);
begin
  spdbtnStopPlaying.Enabled := False;
  spdbtnPlaySelectedFile.Enabled := True;
  spdbtnPlayAllFiles.Enabled := True;
  PlayAllFilesbelowselectedincludingselected1.Enabled := True;
  FStopAllActionsOnDemand := True;
  DoOnSetStopAllActionsOnDemand(FStopAllActionsOnDemand);
end;


procedure TfrmEDAProjectsClickerForm.tmrPlayIconAnimationTimer(Sender: TObject);
begin
  tmrPlayIconAnimation.Tag := tmrPlayIconAnimation.Tag + 1;
  
  if tmrPlayIconAnimation.Tag > 3 then
    tmrPlayIconAnimation.Tag := 0;
end;


procedure TfrmEDAProjectsClickerForm.SetCmbTemplateContentOnAllCmbs;
begin
  SetCmbTemplateContent(cmbTemplate);
  SetCmbTemplateContent(cmbBeforeAllChildTemplates);
  SetCmbTemplateContent(cmbAfterAllChildTemplates);
end;


procedure TfrmEDAProjectsClickerForm.PluginStartup;
begin
  if PluginIsLoaded then
  begin
    try
      DonePlugin;
    except
      on E: Exception do
        SetEDAPluginErrorMessage('Error terminating EDA plugin:' + #13#10 + E.Message, clRed);
    end;

    try
      UnloadPlugin;
    except
      on E: Exception do
        SetEDAPluginErrorMessage('Error unloading EDA plugin:' + #13#10 + E.Message, clRed);
    end;
  end;

  try
    if SizeOf(Pointer) = 4 then
      LoadPlugin(lbeEDAPluginPath32.Text)
    else
      LoadPlugin(lbeEDAPluginPath64.Text);

    try
      InitPlugin(nil);
      SetEDAPluginErrorMessage('Plugin initialized successfully.', clGreen);
    except
      on E: Exception do
        SetEDAPluginErrorMessage('Error initializing EDA plugin: ' + E.Message, clRed);
    end;
  except
    on E: Exception do
      SetEDAPluginErrorMessage('Error loading EDA plugin:' + E.Message, clRed);
  end;

  AddToLog(FastReplace_ReturnToCSV(lblEDAPluginErrorMessage.Caption));
end;


procedure TfrmEDAProjectsClickerForm.tmrStartupTimer(Sender: TObject);
begin
  tmrStartup.Enabled := False;
  SetCmbTemplateContent(cmbTemplate);
  SetCmbTemplateContent(cmbBeforeAllChildTemplates);
  SetCmbTemplateContent(cmbAfterAllChildTemplates);

  if FDefaultTemplate_Main <> '' then
    cmbTemplate.ItemIndex := cmbTemplate.Items.IndexOf(FDefaultTemplate_Main);

  if FDefaultTemplate_Before <> '' then
    cmbBeforeAllChildTemplates.ItemIndex := cmbBeforeAllChildTemplates.Items.IndexOf(FDefaultTemplate_Before);

  if FDefaultTemplate_After <> '' then
    cmbAfterAllChildTemplates.ItemIndex := cmbAfterAllChildTemplates.Items.IndexOf(FDefaultTemplate_After);

  Randomize;

  if ParamCount > 0 then
  begin
    LoadListOfProjects(ParamStr(1));
    Show;
  end;

  PluginStartup;
end;


procedure TfrmEDAProjectsClickerForm.btnReloadPluginClick(Sender: TObject);
begin
  PluginStartup;
end;


procedure TfrmEDAProjectsClickerForm.cmbAfterAllChildTemplatesMouseDown(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  LastClickedComboBoxDefaultString := FDefaultTemplate_After;
end;


procedure TfrmEDAProjectsClickerForm.cmbBeforeAllChildTemplatesMouseDown(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  LastClickedComboBoxDefaultString := FDefaultTemplate_Before;
end;


procedure TfrmEDAProjectsClickerForm.cmbTemplateMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  LastClickedComboBoxDefaultString := FDefaultTemplate_Main;
end;


procedure TfrmEDAProjectsClickerForm.UpdateControlsFromNodeData(ANodeData: PNodeDataProjectRec);
begin
  cmbTemplate.ItemIndex := cmbTemplate.Items.IndexOf(ANodeData^.Template);
  cmbBeforeAllChildTemplates.ItemIndex := cmbBeforeAllChildTemplates.Items.IndexOf(ANodeData^.BeforeAllChildTemplates);
  cmbAfterAllChildTemplates.ItemIndex := cmbAfterAllChildTemplates.Items.IndexOf(ANodeData^.AfterAllChildTemplates);
  
  lbePageSize.Text := ANodeData^.PageSize;
  cmbPageLayout.ItemIndex := cmbPageLayout.Items.IndexOf(ANodeData^.PageLayout);
  lbePageScaling.Text := ANodeData^.PageScaling;
  lbeUserNotes.Text := ANodeData^.UserNotes;

  vallstStatusVariables.Strings.Text := ANodeData^.ExecStatus.AllVars;
  vallstEDAFileCustomVars.Strings.Text := ANodeData^.ListOfEDACustomVars;
end;


procedure TfrmEDAProjectsClickerForm.vallstEDAFileCustomVarsSetEditText(
  Sender: TObject; ACol, ARow: Integer; const Value: string);
begin
  if vallstEDAFileCustomVars.Cells[ACol, ARow] <> Value then
    ToBeUpdated := True;
end;


procedure TfrmEDAProjectsClickerForm.vallstEDAFileCustomVarsValidate(Sender: TObject;
  ACol, ARow: Integer; const KeyName, KeyValue: string);
begin
  if ACol = 0 then
  begin
    if KeyName > '' then
      if {(Length(KeyName) < 2) or} (KeyName[1] <> '$') or (KeyName[Length(KeyName)] <> '$') then
        raise Exception.Create('Variable name must be enclosed by two "$" characters. E.g. "$my_var$" (without double quotes).');
  end;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsChecked(Sender: TBaseVirtualTree;
  Node: PVirtualNode);
begin
  Modified := True;
end;


procedure TfrmEDAProjectsClickerForm.HandleProjectSelectionChange;
var
  Node: PVirtualNode;
  NodeData, FirstNodeData: PNodeDataProjectRec;
begin
  cmbTemplate.Color := clWindow;
  cmbBeforeAllChildTemplates.Color := clWindow;
  cmbAfterAllChildTemplates.Color := clWindow;
  lbePageSize.Color := clWindow;
  cmbPageLayout.Color := clWindow;
  lbePageScaling.Color := clWindow;
  lbeUserNotes.Color := clWindow;
  vallstEDAFileCustomVars.Color := clWindow;

  Node := vstProjects.GetFirstSelected;
  if Node = nil then
    Exit;

  FirstNodeData := vstProjects.GetNodeData(Node);
  UpdateControlsFromNodeData(FirstNodeData);

  vstProjects.BeginUpdate;
  try
    repeat
      if vstProjects.Selected[Node] then
      begin
        NodeData := vstProjects.GetNodeData(Node);

        if NodeData^.Template <> FirstNodeData^.Template then
          cmbTemplate.Color := clYellow;

        if NodeData^.BeforeAllChildTemplates <> FirstNodeData^.BeforeAllChildTemplates then
          cmbBeforeAllChildTemplates.Color := clYellow;

        if NodeData^.AfterAllChildTemplates <> FirstNodeData^.AfterAllChildTemplates then
          cmbAfterAllChildTemplates.Color := clYellow;

        if NodeData^.PageSize <> FirstNodeData^.PageSize then
          lbePageSize.Color := clYellow;

        if NodeData^.PageLayout <> FirstNodeData^.PageLayout then
          cmbPageLayout.Color := clYellow;

        if NodeData^.PageScaling <> FirstNodeData^.PageScaling then
          lbePageScaling.Color := clYellow;

        if NodeData^.UserNotes <> FirstNodeData^.UserNotes then
          lbeUserNotes.Color := clYellow;

        if NodeData^.ListOfEDACustomVars <> FirstNodeData^.ListOfEDACustomVars then
          vallstEDAFileCustomVars.Color := clYellow;
      end;

      Node := vstProjects.GetNext(Node);
    until Node = nil;
  finally
    vstProjects.EndUpdate;
  end;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsClick(Sender: TObject);
begin
  HandleProjectSelectionChange;
  ToBeUpdated := True;   //trigger
  ToBeUpdated := False;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsGetImageIndex(
  Sender: TBaseVirtualTree; Node: PVirtualNode; Kind: TVTImageKind;
  Column: TColumnIndex; var Ghosted: Boolean; var ImageIndex: Integer);
begin
  // dummy
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsGetImageIndexEx(
  Sender: TBaseVirtualTree; Node: PVirtualNode; Kind: TVTImageKind;
  Column: TColumnIndex; var Ghosted: Boolean; var ImageIndex: Integer;
  var ImageList: TCustomImageList);
var
  NodeData: PNodeDataProjectRec;  
begin
  try
    NodeData := vstProjects.GetNodeData(Node);

    case Column of
      0:
      begin
        ImageList := imglstProjects;
        ImageIndex := NodeData^.IconIndex;
      end;

      1:
      begin
        ImageList := imglstProjects;
        ImageIndex := 3;
      end;

      2:
      begin
        ImageList := imglstActionStatus;
        ImageIndex := Ord(NodeData^.ExecStatus.Status);
      end;
    end;
  except
  end;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: {$IFDEF FPC} string {$ELSE} WideString {$ENDIF});
var
  NodeData: PNodeDataProjectRec;
begin
  try
    NodeData := vstProjects.GetNodeData(Node);

    case Column of
      0:
      begin
        if chkDisplayFullPath.Checked then
          CellText := NodeData^.FilePath
        else
          CellText := NodeData^.DisplayedFile;
      end;
      
      1: CellText := NodeData^.Template + ' / ' + NodeData^.BeforeAllChildTemplates + ' /'  + NodeData^.AfterAllChildTemplates;
      2: CellText := CActionStatusStr[NodeData^.ExecStatus.Status];
      3: CellText := NodeData^.PageSize;
      4: CellText := NodeData^.PageLayout;
      5: CellText := NodeData^.PageScaling;
      6:
      begin
        if NodeData^.IconIndex = 0 then
          CellText := IntToStr(Integer(Node^.Index) + 1)
        else
          CellText := '';
      end;
      7: CellText := NodeData^.UserNotes;
      8: CellText := FastReplace_ReturnToCSV(NodeData^.ListOfEDACustomVars);
      9: CellText := NodeData^.DocType;
    end;
  except
    CellText := 'bug';
  end;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsInitNode(Sender: TBaseVirtualTree;
  ParentNode, Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
begin
  Node^.CheckType := ctCheckBox;
  //Node^.CheckState := csMixedNormal;   //commented because of BeginUpdate/EndUpdate
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Node: PVirtualNode;
  NodeData: PNodeDataProjectRec;
begin
  HandleProjectSelectionChange;

  Node := vstProjects.GetFirstSelected;
  if Node = nil then
    Exit;

  NodeData := vstProjects.GetNodeData(Node); 
  UpdateControlsFromNodeData(NodeData);
  ToBeUpdated := False;

  if Key = VK_DELETE then
    RemoveSelectedEDAFiles;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HandleProjectSelectionChange;
end;


procedure TfrmEDAProjectsClickerForm.vstProjectsPaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
begin
  if Displayprojectsinbold1.Checked and (Node^.ChildCount > 0) then
    TargetCanvas.Font.Style := [fsBold]
  else
    TargetCanvas.Font.Style := [];
end;


end.
