unit unMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, System.Actions, Vcl.ActnList,
  cxLocalization, System.ImageList, Vcl.ImgList, cxImageList, cxGraphics,
  cxClasses, cxShellBrowserDialog, cxLookAndFeels, cxLookAndFeelPainters,
  Vcl.Menus, cxControls, cxContainer, cxEdit, cxLabel, Vcl.StdCtrls, cxButtons,
  dxShellDialogs, System.Types, System.IOUtils, System.Masks, dxCore,
  dxCoreClasses, dxHashUtils, dxSpreadSheetCore, dxSpreadSheetCoreFormulas,
  dxSpreadSheetCoreHistory, dxSpreadSheetCoreStyles, dxSpreadSheetCoreStrs,
  dxSpreadSheetConditionalFormatting, dxSpreadSheetConditionalFormattingRules,
  dxSpreadSheetClasses, dxSpreadSheetContainers, dxSpreadSheetFormulas,
  dxSpreadSheetHyperlinks, dxSpreadSheetFunctions, dxSpreadSheetStyles,
  dxSpreadSheetGraphics, dxSpreadSheetPrinting, dxSpreadSheetTypes,
  dxSpreadSheetUtils, dxSpreadSheetFormattedTextUtils, dxBarBuiltInMenu,
  dxSpreadSheet, ComObj;

type
  TfrmMain = class(TForm)
    cxShellBrowserDialog: TcxShellBrowserDialog;
    cxImageList: TcxImageList;
    Localizer: TcxLocalizer;
    ActionList: TActionList;
    acOpenFolder: TAction;
    acExecute: TAction;
    acAbout: TAction;
    acSaveExcelFile: TAction;
    Label1: TLabel;
    btnOpenFolder: TcxButton;
    lbFolder: TcxLabel;
    cxButton1: TcxButton;
    Label2: TLabel;
    acSelectShablon: TAction;
    ofdOpenShablon: TdxOpenFileDialog;
    cxButton2: TcxButton;
    lbFileShablon: TcxLabel;
    cxButton3: TcxButton;
    acOpenFolderOut: TAction;
    lbFolderOut: TcxLabel;
    Label3: TLabel;
    cxShellBrowserDialog_OUT: TcxShellBrowserDialog;
    SST: TdxSpreadSheet;
    SSC: TdxSpreadSheet;
    procedure FormCreate(Sender: TObject);
    procedure acExecuteExecute(Sender: TObject);
    procedure acSelectShablonExecute(Sender: TObject);
    procedure acOpenFolderExecute(Sender: TObject);
    procedure acOpenFolderOutExecute(Sender: TObject);
  private
    { Private declarations }
  const Mask ='вислуга*.xlsx';
  //const Mask ='вислуга БОЙЧЕНКО*.xlsx';

  function GetFileCount(const APath:string; Mask: string): Integer;
  function CopyFile_ (sFileName:string):string;
  procedure RangeCopyDate(WorkbookSrc, WorkbookDst:variant;
                          SheetSrc, SheetDst:integer;
                          RangeSrc, RangeDst:String);
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;


implementation

uses
 unExcelFunctions, unToolsFunctions;

{$R *.dfm}
{$R RUS.RES}

function GetAppVersion(FileName: string): string;
var
  Dummy, InfoSize: DWORD;
  Info: Pointer;
  VerValue: PVSFixedFileInfo;
  VerSize: UINT;
begin
  Result := '';
  InfoSize := GetFileVersionInfoSize(PChar(FileName), Dummy);
  if InfoSize = 0 then Exit;

  GetMem(Info, InfoSize);
  try
    if GetFileVersionInfo(PChar(FileName), Dummy, InfoSize, Info) then
      if VerQueryValue(Info, '\', Pointer(VerValue), VerSize) then
        with VerValue^ do
          Result := Format('%d.%d.%d.%d',
            [HiWord(dwFileVersionMS), LoWord(dwFileVersionMS),
             HiWord(dwFileVersionLS), LoWord(dwFileVersionLS)]);
  finally
    FreeMem(Info);
  end;
end;


procedure TfrmMain.FormCreate(Sender: TObject);
var
  rs: TResourceStream;
begin
  Caption:=Caption+' версия '+GetAppVersion(Application.ExeName);

  try
    rs := TResourceStream.Create(HInstance, 'RUSINI', RT_RCDATA);
    try
      Localizer.LoadFromStream(rs);
      Localizer.LanguageIndex := 1;
      Localizer.Translate;
    finally
      FreeAndNil(rs);
    end;
  except
    //
  end;

  lbFolder.Caption:=ReadStringIniFiles('Prosto','PATH_FILES', '\\A4785\Disc\ВИСЛУГА РОКІВ А4785\Вислуга років похибка');
  lbFileShablon.Caption:=ReadStringIniFiles('Prosto','FILE_SHABLON', '\\A4785\Disc\ВИСЛУГА РОКІВ А4785\Вислуга Виправлена+++.xlsx');
  lbFolderOut.Caption:=ReadStringIniFiles('Prosto','PATH_FILES_OUT', '\\A4785\Disc\ВИСЛУГА РОКІВ А4785\Вислуга років без похибки');

  if (DirectoryExists(lbFolder.Caption))
  and (FileExists(lbFileShablon.Caption))
  and (DirectoryExists(lbFolderOut.Caption))
  then
  acExecute.Enabled:=True;
end;


procedure TfrmMain.acOpenFolderExecute(Sender: TObject);
begin
  if cxShellBrowserDialog.Execute then
  begin
    lbFolder.Caption:=cxShellBrowserDialog.Path;
    SaveStringIniFiles('Prosto','PATH_FILES', cxShellBrowserDialog.Path);

    if (DirectoryExists(lbFolder.Caption))
    and (FileExists(lbFileShablon.Caption))
    and (DirectoryExists(lbFolderOut.Caption))
    then
      if acExecute.Enabled=False then acExecute.Enabled:=True;
  end;
end;

procedure TfrmMain.acSelectShablonExecute(Sender: TObject);
begin
  if ofdOpenShablon.Execute then
  begin
    lbFileShablon.Caption:=ofdOpenShablon.FileName;
    SaveStringIniFiles('Prosto','FILE_SHABLON', ofdOpenShablon.FileName);

    if (DirectoryExists(lbFolder.Caption))
    and (FileExists(lbFileShablon.Caption))
    and (DirectoryExists(lbFolderOut.Caption))
    then
      if acExecute.Enabled=False then acExecute.Enabled:=True;
  end;
end;

procedure TfrmMain.acOpenFolderOutExecute(Sender: TObject);
begin
  if cxShellBrowserDialog_OUT.Execute then
  begin
    lbFolderOut.Caption:=cxShellBrowserDialog_OUT.Path;
    SaveStringIniFiles('Prosto','PATH_FILES_OUT', cxShellBrowserDialog_OUT.Path);

    if (DirectoryExists(lbFolder.Caption))
    and (FileExists(lbFileShablon.Caption))
    and (DirectoryExists(lbFolderOut.Caption))
    then
      if acExecute.Enabled=False then acExecute.Enabled:=True;
  end;
end;

function TfrmMain.GetFileCount(const APath:string; Mask: string): Integer;
var
  Files: TStringDynArray;
begin
  // Получаем массив путей к файлам
  Files := TDirectory.GetFiles(APath,
           //TSearchOption.soAllDirectories,
           TSearchOption.soTopDirectoryOnly,
           function(const Path: string; const SearchRec: TSearchRec): Boolean
           begin
             Result := MatchesMask(WideUpperCase(SearchRec.Name), WideUpperCase(Mask));
           end);

  // Возвращаем размер массива
  Result := Length(Files);
 end;

function TfrmMain.CopyFile_(sFileName:string):string;
begin
  if FileExists(sFileName) then
  begin
   CopyFile(PWideChar(sFileName), PWideChar(ChangeFileExt(ExtractFileName(sFileName),'')+'_временный.xlsx'), False);
   Result:=ChangeFileExt(ExtractFileName(sFileName),'')+'_временный.xlsx';
  end;
end;

procedure TfrmMain.acExecuteExecute(Sender: TObject);
 var
 Folder, FolderOut, FileShablon, FileName, FileName_temp, FileName_correct:string;
 row,n:integer;
 ExcelApp, SrcWb, DstWb, SrcRange, DstRange: Variant;
begin
  Folder:=lbFolder.Caption;
  FolderOut:=lbFolderOut.Caption;
  FileShablon:=lbFileShablon.Caption;

  if (DirectoryExists(Folder))
  and (DirectoryExists(FolderOut))
  and (FileExists(FileShablon))
  and (GetFileCount(Folder, Mask)>0) then
  begin
     //Копируем файл шаблона
     CopyFile(PWideChar(FileShablon),PWideChar(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileShablon)),False);


    //ВРЕМЕННО
    //n:=0;

    ShowProgress(Self, '', GetFileCount(Folder, Mask), 'Анализ указанной папки...');
    for FileName in TDirectory.GetFiles(Folder,
                     //TSearchOption.soAllDirectories,
                     TSearchOption.soTopDirectoryOnly,
                     function(const Path: string; const SearchRec: TSearchRec): Boolean
                     begin
                       Result := MatchesMask(WideUpperCase(SearchRec.Name), WideUpperCase(Mask));
                     end)
    do
    begin
      try
        UpdateProgressCaption(ExtractFileName(FileName), 'Анализ указанной папки...');
        StepProgress;
        Application.ProcessMessages;

         //ВРЕМЕННО
         //n:=n+1;

         CopyFile(
         PWideChar(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileShablon)),
         PWideChar(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileName)),False);

         FileName_temp:=ExtractFileDir(Application.ExeName)+'\'+ CopyFile_(FileName);
         FileName_correct:=ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileName);


         ExcelApp := CreateOleObject('Excel.Application');

         ExcelApp.Visible := False;
         ExcelApp.DisplayAlerts := False;

         SrcWb := ExcelApp.Workbooks.Open(FileName_temp);
         DstWb := ExcelApp.Workbooks.Open(FileName_correct);


         //Копируем имя человека
         SrcRange := SrcWb.Worksheets[1].Range['B1'];
         DstRange := DstWb.Worksheets[1].Range['B1'];
         SrcRange.Copy(DstRange);

         row:=4;
         while row<=118 do
         begin

           {
           SrcRange := SrcWb.Worksheets[1].Range['B'+InttoStr(row)+':Q'+InttoStr(row+1)];
           DstRange := DstWb.Worksheets[1].Range['B'+InttoStr(row)+':Q'+InttoStr(row+1)];
           SrcRange.Copy(DstRange);
           }

           RangeCopyDate(SrcWb, DstWb,
                          1,1, 'B'+InttoStr(row)+':Q'+InttoStr(row+1),
                          'B'+InttoStr(row)+':Q'+InttoStr(row+1)
                        );

           row:=row+6;
         end;


         DstWb.Save;

         //Переносим файл в указанную папку
         CopyFile(
         PWideChar(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileName_correct)),
         PWideChar(FolderOut+'\'+ExtractFileName(FileName_correct)), False);


         Application.ProcessMessages;
      finally
        if not VarIsEmpty(SrcWb) then SrcWb.Close(False);
        if not VarIsEmpty(DstWb) then DstWb.Close(False);
        if not VarIsEmpty(ExcelApp) then ExcelApp.Quit;
        ExcelApp := Unassigned;

        if FileExists(FileName_temp) then
        DeleteFile(FileName_temp);
        if FileExists(FileName_correct) then
        DeleteFile(FileName_correct);
      end;

       //ВРЕМЕННО
       //if n=1 then Break
    end;
    DelProgress;

    if FileExists(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileShablon)) then
    DeleteFile(ExtractFileDir(Application.ExeName)+'\'+ExtractFileName(FileShablon));
  end;
  ShowMessage('Все');
end;



procedure TfrmMain.RangeCopyDate(WorkbookSrc, WorkbookDst:variant;
                                 SheetSrc, SheetDst:integer;
                                 RangeSrc, RangeDst:String);
var
 DataArray: OleVariant;
 Src1904, Dst1904: Boolean;
 Row, Col: Integer;
begin
  Src1904 := WorkbookSrc.Date1904;
  Dst1904 := WorkbookDst.Date1904;

  DataArray := WorkbookSrc.Sheets[SheetSrc].Range[RangeSrc].Value2;

  if Src1904 <> Dst1904 then
  begin
    for Row := VarArrayLowBound(DataArray, 1) to VarArrayHighBound(DataArray, 1) do
      for Col := VarArrayLowBound(DataArray, 2) to VarArrayHighBound(DataArray, 2) do
      begin
        if VarIsNumeric(DataArray[Row, Col]) then
        begin
          if Src1904 and not Dst1904 then
            DataArray[Row, Col] := DataArray[Row, Col] + 1462
          else
            DataArray[Row, Col] := DataArray[Row, Col] - 1462;
        end;
      end;
  end;

  WorkbookDst.Sheets[SheetDst].Range[RangeDst].Value2 := DataArray;
end;



end.
