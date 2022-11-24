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


unit EDAUIClickerPlugins;

{$mode Delphi}

interface

uses
  Windows, Classes, SysUtils;


type
  TInitPlugin = function(AExtraData: Pointer; AErrMsg: Pointer): Integer; cdecl;
  TGetListsOfKnownEDAProjectExtensions = function(AListOfExtensions: Pointer): Integer; cdecl;
  TGetEDAListsOfDocuments = procedure(AEDAProjectFileNameStr: Pointer; AListOfDocuments, AListOfTypes: Pointer; out AListOfDocumentsLen, AListOfTypesLen: Integer); cdecl;
  TDonePlugin = procedure; cdecl;



procedure LoadPlugin(APluginPath: string);
function PluginIsLoaded: Boolean;
procedure UnloadPlugin;

procedure InitPlugin(AExtraData: Pointer);
procedure GetListsOfKnownEDAProjectExtensions(AListOfExtensions: TStringList);
procedure GetEDAListsOfDocuments(AEDAProjectFileName: string; AListOfDocuments, AListOfTypes: TStringList);
procedure DonePlugin;


implementation


uses
  DllUtils;


const
  CEDAPluginIsNotLoadedErr = 'EDA plugin is not loaded.';

var
  PluginHandle: THandle = 0;

  InitPlugin_Proc: TInitPlugin = nil;
  GetListsOfKnownEDAProjectExtensions_Proc: TGetListsOfKnownEDAProjectExtensions = nil;
  GetEDAListsOfDocuments_Proc: TGetEDAListsOfDocuments = nil;
  DonePlugin_Proc: TDonePlugin = nil;


procedure LoadPlugin(APluginPath: string);
begin
  if not FileExists(APluginPath) then
    raise Exception.Create(APluginPath + ' does not exist.');

  PluginHandle := LoadLibrary(PChar(APluginPath));
  if PluginHandle > 0 then
  begin
    @InitPlugin_Proc := GetProcAddress(PluginHandle, 'InitPlugin');
    @GetListsOfKnownEDAProjectExtensions_Proc := GetProcAddress(PluginHandle, 'GetListsOfKnownEDAProjectExtensions');
    @GetEDAListsOfDocuments_Proc := GetProcAddress(PluginHandle, 'GetEDAListsOfDocuments');
    @DonePlugin_Proc := GetProcAddress(PluginHandle, 'DonePlugin');
  end;
end;


function PluginIsLoaded: Boolean;
begin
  Result := PluginHandle > 0;
end;


procedure UnloadPlugin;
begin
  if PluginIsLoaded then
    FreeLibrary(PluginHandle);
end;


procedure InitPlugin(AExtraData: Pointer);
var
  Err: string;
begin
  if @InitPlugin_Proc <> nil then
  begin
    SetLength(Err, CMaxSharedStringLength); //preallocate
    SetLength(Err, InitPlugin_Proc(AExtraData, @Err[1]));
    if Err <> '' then
      raise Exception.Create(Err);
  end
  else
    raise Exception.Create('InitPlugin function is not available.');
end;


procedure GetListsOfKnownEDAProjectExtensions(AListOfExtensions: TStringList);
var
  Extensions: string;
begin
  if not PluginIsLoaded then
    raise Exception.Create(CEDAPluginIsNotLoadedErr);

  if @GetListsOfKnownEDAProjectExtensions_Proc <> nil then
  begin
    SetLength(Extensions, CMaxSharedStringLength); //preallocate
    SetLength(Extensions, GetListsOfKnownEDAProjectExtensions_Proc(@Extensions[1]));
    AListOfExtensions.Text := Extensions;
  end
  else
    raise Exception.Create('GetListsOfKnownEDAProjectExtensions function is not available.');
end;


procedure GetEDAListsOfDocuments(AEDAProjectFileName: string; AListOfDocuments, AListOfTypes: TStringList);
var
  Documents, DocumentTypes: string;
  DocLen, DocTypeLen: Integer;
begin
  if not PluginIsLoaded then
    raise Exception.Create(CEDAPluginIsNotLoadedErr);

  if @GetListsOfKnownEDAProjectExtensions_Proc <> nil then
  begin
    SetLength(Documents, CMaxSharedStringLength); //preallocate
    SetLength(DocumentTypes, CMaxSharedStringLength); //preallocate

    GetEDAListsOfDocuments_Proc(@AEDAProjectFileName[1], @Documents[1], @DocumentTypes[1], DocLen, DocTypeLen);

    SetLength(Documents, DocLen); //preallocate
    SetLength(DocumentTypes, DocTypeLen); //preallocate
    AListOfDocuments.Text := Documents;
    AListOfTypes.Text := DocumentTypes;
  end
  else
    raise Exception.Create('GetEDAListsOfDocuments procedure is not available.');
end;


procedure DonePlugin;
begin
  if @DonePlugin_Proc <> nil then
    DonePlugin_Proc
  else
    raise Exception.Create('DonePlugin procedure is not available.');
end;

end.


