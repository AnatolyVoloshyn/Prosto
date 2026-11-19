program Prosto;

uses
  Vcl.Forms,
  unMain in '..\SOURCE\unMain.pas' {frmMain},
  Process_ in '..\SOURCE\Process_.pas' {frmProcess},
  unToolsFunctions in '..\SOURCE\unToolsFunctions.pas',
  unIni in '..\SOURCE\unIni.pas',
  unExcelFunctions in '..\SOURCE\unExcelFunctions.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
