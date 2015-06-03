unit main;

{
* Author    A.Kouznetsov
* Rev       1.01 dated 4/6/2015
Redistribution and use in source and binary forms, with or without modification, are permitted.
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.}


{$mode objfpc}{$H+}

// ###########################################################################
interface
// ###########################################################################

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Menus,
  StdCtrls, Grids, ExtCtrls, ComCtrls, IniFiles, sch_to_bom, CSVdocument;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    GroupBox1: TGroupBox;
    ListBox1: TListBox;
    MainMenu1: TMainMenu;
    FileMenuItem: TMenuItem;
    AboutMenuItem: TMenuItem;
    MenuItem1: TMenuItem;
    MenuItem3: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl: TPageControl;
    SaveDialog1: TSaveDialog;
    SaveMenuItem: TMenuItem;
    MenuItem2: TMenuItem;
    ExitMenuItem: TMenuItem;
    Splitter1: TSplitter;
    StatusBar: TStatusBar;
    SG1: TStringGrid;
    SG2: TStringGrid;
    StringGrid1: TStringGrid;
    TabSheet1: TTabSheet;
    TabSheetSep: TTabSheet;
    TabSheetGr: TTabSheet;
    procedure AboutMenuItemClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure ExitMenuItemClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure SaveMenuItemClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  SchBOM : TKicadSchBOM;
  BomFileName : string;
const
  revision = '1.01';

// ###########################################################################
implementation
// ###########################################################################


{$R *.lfm}

// ==================================================
// Remove file ext
// ==================================================
function file_name_no_ext(fn : string) : string;
var i : integer;
    c : char;
begin
  result := fn;
  for i := Length(fn) downto 2 do begin
    c := fn[i];
    if c = '.' then begin
      result := Copy(fn,1,i-1);
      break;
    end;
  end;
end;

{ TForm1 }
// ==================================================
// Create
// ==================================================
procedure TForm1.FormCreate(Sender: TObject);
var Inif : TIniFile;
    s, si : string;
    i : integer;
begin
  SchBOM := TKicadSchBOM.Create;
  Inif := TIniFile.Create('KiCadBOM.ini');
  // --------------------------
  // Exclude components with certain RefDes
  // --------------------------
  for i:=1 to 1000 do begin
    si := 'Ref'+IntToStr(i);
    s := AnsiUpperCase(Inif.ReadString('Excluded Ref',si,''));
    if s = '' then
      break
    else
      SchBOM.FExcludedRef.Add(s);
  end;
  // --------------------------
  // Exclude library prefixes from footprints
  // --------------------------
  for i:=1 to 1000 do begin
    si := 'Prefix'+IntToStr(i);
    s := AnsiUpperCase(Inif.ReadString('Excluded Footprint Prefix',si,''));
    if s = '' then
      break
    else
      SchBOM.FExcludedFootPrefix.Add(s);
  end;
  Inif.Free;
end;

// ==================================================
// New BOM
// ==================================================
procedure TForm1.MenuItem1Click(Sender: TObject);
begin
  Button1Click(Sender);
  Button2Click(Sender);
  BomFileName := ExtractFileName(OpenDialog1.FileName);
  BomFileName := file_name_no_ext(BomFileName);
  StatusBar.Panels[2].Text := BomFileName;
end;

// ==================================================
// Add SCH file
// ==================================================
procedure TForm1.MenuItem3Click(Sender: TObject);
begin
  Button2Click(Sender);
end;

// ==================================================
// Save BOM
// ==================================================
procedure TForm1.SaveMenuItemClick(Sender: TObject);
begin
  Button4Click(Sender);
end;

// ==================================================
// Close
// ==================================================
procedure TForm1.ExitMenuItemClick(Sender: TObject);
begin
  SchBOM.Free;
  Application.Terminate;
end;

// ==================================================
// About
// ==================================================
procedure TForm1.AboutMenuItemClick(Sender: TObject);
begin
  ShowMessage('KiCAD BOM Generator rev '+revision+char(13)+'This program is free software');
end;

// ==================================================
// New BOM
// ==================================================
procedure TForm1.Button1Click(Sender: TObject);
begin
  SchBOM.Reset;
  SchBOM.ToStringGrid(SG1);
  SG1.RowCount := 50;
  SchBOM.FPartsList.ToStringGrid(SG2);
  SG2.RowCount := 50;
  StatusBar.Panels[0].Text:=IntToStr(SchBOM.Count-1) + ' parts';
  StatusBar.Panels[1].Text:=IntToStr(SchBOM.FPartsList.Count-1) + ' items';
end;

// ==================================================
// Add SCH
// ==================================================
procedure TForm1.Button2Click(Sender: TObject);
begin
  if OpenDialog1.execute then begin
     SchBOM.ReadFile(OpenDialog1.FileName, true);
     SchBOM.ToStringGrid(SG1);
     SchBOM.FPartsList.ToStringGrid(SG2);
     StatusBar.Panels[0].Text:=IntToStr(SchBOM.Count-1) + ' parts';
     StatusBar.Panels[1].Text:=IntToStr(SchBOM.FPartsList.Count-1) + ' items';
  end;
end;

// ==================================================
// Update missed fields
// ==================================================
procedure TForm1.Button3Click(Sender: TObject);
begin
  SchBOM.UpdateFields;
  SchBOM.ToStringGrid(SG1);
end;

// ==================================================
// Save BOM
// ==================================================
procedure TForm1.Button4Click(Sender: TObject);
begin
  if PageControl.TabIndex = 0 then begin;
    SaveDialog1.Title := 'BOM with individual items - save as';
    SaveDialog1.FileName := BomFileName+'_BOM.csv';
    if SaveDialog1.execute then begin
       SchBOM.SaveToFile(SaveDialog1.FileName);
    end;
  end else begin
  SaveDialog1.Title := 'BOM with grouped items - save as';
    SaveDialog1.FileName := BomFileName+'_PL.csv';
    if SaveDialog1.execute then begin
       SchBOM.FPartsList.SaveToFile(SaveDialog1.FileName);
    end;
  end;
end;

// ==================================================
// Make parts list
// ==================================================
procedure TForm1.Button5Click(Sender: TObject);
begin
  SchBOM.MakePartList;
end;


end.
