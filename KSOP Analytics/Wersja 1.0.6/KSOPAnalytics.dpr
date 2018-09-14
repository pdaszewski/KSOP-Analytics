program KSOPAnalytics;

uses
  Vcl.Forms,
  AOknoGl_frm in 'OknaProgramu\AOknoGl_frm.pas' {AOknoGl},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'KSOP Analytics';
  TStyleManager.TrySetStyle('Carbon');
  Application.CreateForm(TAOknoGl, AOknoGl);
  Application.Run;
end.
