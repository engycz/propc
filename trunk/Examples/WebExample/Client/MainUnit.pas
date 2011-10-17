unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, prOpcVSC;

type
  TForm1 = class(TForm)
    WidthShrinkButton: TSpeedButton;
    WidthGrowButton: TSpeedButton;
    HeightShrinkButton: TSpeedButton;
    HeightGrowButton: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure WidthShrinkButtonClick(Sender: TObject);
    procedure HeightGrowButtonClick(Sender: TObject);
    procedure HeightShrinkButtonClick(Sender: TObject);
    procedure WidthGrowButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    VSC: Variant;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  VSC:= OpcClient('', 'Wizzy1.TWizzyServer.1')
end;

procedure TForm1.WidthGrowButtonClick(Sender: TObject);
begin
  VSC.FormWidth:= VSC.FormWidth + 20
end;

procedure TForm1.WidthShrinkButtonClick(Sender: TObject);
begin
  VSC.FormWidth:= VSC.FormWidth - 20
end;

procedure TForm1.HeightGrowButtonClick(Sender: TObject);
begin
  VSC.FormHeight:= VSC.FormHeight + 20
end;

procedure TForm1.HeightShrinkButtonClick(Sender: TObject);
begin
  VSC.FormHeight:= VSC.FormHeight - 20
end;


end.
