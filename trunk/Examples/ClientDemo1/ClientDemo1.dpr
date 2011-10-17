program ClientDemo1;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  BrowseUnit in 'BrowseUnit.pas' {BrowseDlg};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TBrowseDlg, BrowseDlg);
  Application.Run;
end.
