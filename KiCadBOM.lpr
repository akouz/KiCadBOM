program KiCadBOM;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, main, sch_to_bom, bom_to_pl, csvdocument, inventory_unit
  { you can add units after this };

{$R *.res}

begin
  Application.Title:='KiCadBOM';
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
