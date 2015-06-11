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
  Classes, SysUtils, Windows, FileUtil, Forms, Controls, Graphics, Dialogs, Menus,
  StdCtrls, Grids, ExtCtrls, ComCtrls, IniFiles, sch_to_bom, inventory_unit;

type

  { TForm1 }

  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    MainMenu1: TMainMenu;
    FileMenuItem: TMenuItem;
    Pop1CopyTo: TMenuItem;
    Pop1CopyFrom: TMenuItem;
    Pop1Find: TMenuItem;
    MiAbout: TMenuItem;
    MiNewBom: TMenuItem;
    MiAddSch: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    PageControl2: TPageControl;
    Pop1: TPopupMenu;
    Pop2: TPopupMenu;
    SaveDialog1: TSaveDialog;
    MiSaveBom: TMenuItem;
    MenuItem2: TMenuItem;
    MiExit: TMenuItem;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    StatusBar: TStatusBar;
    SG1: TStringGrid;
    SG2: TStringGrid;
    SG3: TStringGrid;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheetSep: TTabSheet;
    TabSheetGr: TTabSheet;
    TreeView1: TTreeView;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure MiAboutClick(Sender: TObject);
    procedure MiExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure MiNewBomClick(Sender: TObject);
    procedure MiAddSchClick(Sender: TObject);
    procedure MiSaveBomClick(Sender: TObject);
    procedure Pop1FindClick(Sender: TObject);
    procedure SG2GetCellHint(Sender: TObject; ACol, ARow: Integer;
      var HintText: String);
    procedure SG2MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure TreeView1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  SchBOM : TKicadSchBOM;
  Stock : TKicadStock;
  BomFileName : string;
const
  revision = '1.02';

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
    s, ss, sn, si, inv_path, fn : string;
    i, j, k, x, n : integer;
    sec: TStringList;
    inv : boolean;
    ttn : TTreeNode;
begin
  SchBOM := TKicadSchBOM.Create;
  Stock := TKicadStock.Create ;
  Inif := TIniFile.Create('KiCadBOM.ini');
  // --------------------------
  // Restore window
  // --------------------------
  i := Inif.ReadInteger('Window','Width',0);
  if i > 400 then
    Form1.Width := i;
  i := Inif.ReadInteger('Window','Height',0);
  if i > 200 then
    Form1.Height := i;
  i := Inif.ReadInteger('Window','Left',0);
  Form1.Left := i;
  i := Inif.ReadInteger('Window','Top',0);
  Form1.Top := i;
  PageControl1.Width := Inif.ReadInteger('Window','Split',200);
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
  // --------------------------
  // Read inventory (stock) path
  // --------------------------
  inv := false;
  si := 'Path';
  inv_path := AnsiUpperCase(Inif.ReadString('Inventory',si,''));
  Inif.Free;
  // --------------------------
  // Read inventory
  // --------------------------
  if inv_path <> '' then begin
    k := 0;
    sec := TStringList.Create;
    Inif := TIniFile.Create(inv_path+'\Inventory.ini');
    Inif.ReadSections(sec);
    for j:=0 to sec.Count-1 do begin
     sn := sec.Strings[j];
     if j = 0 then
       TreeView1.Items.Add(nil,sn)
     else
       TreeView1.Items.Add(TreeView1.Items[k],sn);
     k := TreeView1.Items.Count-1;
     for i:=1 to 1000 do begin
      // ---------------------
      // get file name
      // ---------------------
       si := 'File'+IntToStr(i);
       s := Inif.ReadString(sn,si,'');
       if s = '' then
         break                   // finish section
       else begin
         // ---------------------
         // get name for TreeView
         // ---------------------
         fn := inv_path + '\' + s;
         if FileExists(fn) then begin
           // ---------------------
           // get Ref name
           // ---------------------
           si := 'Ref'+IntToStr(i);
           ss := Trim(Inif.ReadString(sn,si,'?'));
           Stock.RefNames.Add(ss);
           // ---------------------
           // get name
           // ---------------------
           si := 'Name'+IntToStr(i);
           ss := Inif.ReadString(sn,si,'');
           if ss = '' then
             ss := file_name_no_ext(s);      // use file name
           Stock.FileNames.Add(fn);
           x := Stock.FileNames.Count;
           ttn := TreeView1.Items.AddChildObject(TreeView1.Items[k],ss, Pointer(x));
           // ---------------------
           // load first to SG3
           // ---------------------
           if not inv then begin
             inv := true;
             ttn.Selected := true;
             n := Integer(ttn.Data);
             if n > 0 then begin
               Stock.ToStringGrid(SG3, n-1);
             end;
           end;
         end;
       end;
     end;
    end;
    sec.Free;
    Inif.Free;
  end;
end;

// ==================================================
// Destroy
// ==================================================
procedure TForm1.FormDestroy(Sender: TObject);
begin
  SchBOM.Free;
  Stock.Free;
end;

// ==================================================
// Form resize
// ==================================================
procedure TForm1.FormResize(Sender: TObject);
var
  x, y, w, aw, i: integer;
  MaxWidth: integer;
  SG : TStringGrid;
begin
  for i := 1 to 3 do begin
    case i of
     1 : SG := SG1;
     2 : SG := SG2;
     3 : SG := SG3;
    end;
    aw := 0;
    with SG do  begin
      for x := 0 to ColCount-2 do   begin
        MaxWidth := Canvas.TextWidth(Columns[x].Title.Caption)+10;
        if MaxWidth < 50 then
          MaxWidth := 50;
        for y := 0 to RowCount-1 do
        begin
          w := Canvas.TextWidth(Cells[x,y]);
          if w > MaxWidth then
            MaxWidth := w;
        end;
        ColWidths[x] := MaxWidth + 10;
        aw := aw + ColWidths[x];
      end;
      ColWidths[ColCount-1] := ClientWidth - aw;
    end;
  end;
end;

// ==================================================
// New BOM
// ==================================================
procedure TForm1.MiNewBomClick(Sender: TObject);
begin
  SchBOM.Reset;
  SchBOM.ToStringGrid(SG1);
  SG1.RowCount := 50;
  SchBOM.FPartsList.ToStringGrid(SG2);
  SG2.RowCount := 50;
  MiAddSchClick(Sender);
  BomFileName := ExtractFileName(OpenDialog1.FileName);
  BomFileName := file_name_no_ext(BomFileName);
  StatusBar.Panels[2].Text := BomFileName;
  FormResize(Sender);
end;

// ==================================================
// Add SCH file
// ==================================================
procedure TForm1.MiAddSchClick(Sender: TObject);
begin
  if OpenDialog1.execute then begin
     SchBOM.ReadFile(OpenDialog1.FileName, true);
     SchBOM.ToStringGrid(SG1);
     SchBOM.FPartsList.ToStringGrid(SG2);
     FormResize(Sender);
  end;
  StatusBar.Panels[0].Text:=IntToStr(SchBOM.Count-1) + ' items';
  StatusBar.Panels[1].Text:=IntToStr(SchBOM.FPartsList.Count-1) + ' groups';
end;

// ==================================================
// Save BOM
// ==================================================
procedure TForm1.MiSaveBomClick(Sender: TObject);
begin
  if PageControl1.TabIndex = 0 then begin;
    SaveDialog1.Title := 'Save BOM';
    SaveDialog1.FileName := BomFileName+'_BOM.csv';
    if SaveDialog1.execute then begin
      SchBOM.SaveCSV(SaveDialog1.FileName, BomFileName, SG1);
    end;
  end else begin
    SaveDialog1.Title := 'Save Grouped BOM';
    SaveDialog1.FileName := BomFileName+'_PL.csv';
    if SaveDialog1.execute then begin
      SchBOM.SaveCSV(SaveDialog1.FileName, BomFileName, SG2);
    end;
  end;
  FormResize(Sender);
end;

// ==================================================
// Find selected in inventory
// ==================================================
procedure TForm1.Pop1FindClick(Sender: TObject);
begin

end;

// ==================================================
// SG2 cell hint
// ==================================================
procedure TForm1.SG2GetCellHint(Sender: TObject; ACol, ARow: Integer;
  var HintText: String);
begin
  if ACol = 0 then
    SG2.Hint:=HintText;
end;

// ==================================================
// SG2 mouse move
// ==================================================
procedure TForm1.SG2MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer
  );
var
  Row, Col: Integer;
  s : string;
begin
  Row := 1;
  Col := 1;
  SG2.MouseToCell(X, Y, Col, Row);
  if Col = 0 then begin
    s := SchBOM.FPartsList.GetRef(Row);
    SG2GetCellHint(Sender, Col, Row, s);
  end else
    Application.CancelHint;
end;

// ==================================================
// Click on inventory group
// ==================================================
procedure TForm1.TreeView1Click(Sender: TObject);
var ttn : TTreeNode;
    i : integer;
begin
  ttn := TreeView1.Selected;
  if ttn = nil then
    exit;
  if not ttn.HasChildren then begin
    i := Integer(ttn.Data);
    if i > 0 then begin
      Stock.ToStringGrid(SG3, i-1);
    end;
  end;
end;

// ==================================================
// Exit
// ==================================================
procedure TForm1.MiExitClick(Sender: TObject);
begin
  SchBOM.Free;
  Application.Terminate;
end;

// ==================================================
// About
// ==================================================
procedure TForm1.MiAboutClick(Sender: TObject);
begin
  ShowMessage('KiCAD BOM Generator rev '+revision+char(13)+'This program is free software');
end;

// ==================================================
// Close
// ==================================================
procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var Inif : TIniFile;
begin
  Inif := TIniFile.Create('KiCadBOM.ini');
  Inif.WriteInteger('Window','Width',Form1.Width);
  Inif.WriteInteger('Window','Height',Form1.Height);
  Inif.WriteInteger('Window','Left',Form1.Left);
  Inif.WriteInteger('Window','Top',Form1.Top);
  Inif.WriteInteger('Window','Split',Splitter2.Left);
  Inif.Free;
end;

end.

