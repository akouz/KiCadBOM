unit sch_to_bom;

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
  Classes, SysUtils, Grids, bom_to_pl;

type

{ TKicadSchBOM }

TKicadSchBOM = class(TStringList)
private
    quote_mode : boolean;      // quotation mode
    comp_mode : boolean;      // component mode
    comp_cnt : integer;
    Ftmp : TStringList;
    FComp: TStringList;  // current component
    FSrc : TStringList;
    procedure ClrComp;
    procedure AddComp;
    function  ExFootprintPrefix(val : string) : string;
    procedure FillCompField(fld : integer; val : string);
    procedure FillComp;
    procedure ParseFile;
    procedure ParseString(ss:string);
public
    FileName : string;
    FExcludedRef : TStringList;
    FRefList : TStringList;   // list of RefDes
    FExcludedFootPrefix : TStringList;
    FPartsList : TKicadSchPL;
    procedure MakePartList;
    procedure UpdateFields;
    procedure ReadFile(fn:string; appnd : boolean);
    procedure SaveCSV(fn:string; BOMname: string; SG : TStringGrid);
    procedure ToStringGrid(var SG : TStringGrid);
    procedure Reset;
    constructor Create;
    destructor Destroy; override;
end;

function ExtractRefDesName(ss : string; var n : integer) : string;
function ExpandNumber(ss : string): string;
function MySort(List: TStringList; Index1, Index2: Integer): Integer;

// ###########################################################################
implementation
// ###########################################################################

// ==================================================
// Extract name from sepRef
// ==================================================
function ExtractRefDesName(ss : string; var n : integer) : string;
var i : integer;
    s : string;
    c : char;
    name : boolean;
begin
  n := 0;
  s := '';
  name := true;
  for i := 1 to Length(ss) do begin
    c := ss[i];
    if c in ['0'..'9'] then begin
      name := false;
      n := n*10 + ord(c) - ord('0');
    end else begin
      if name then
        s := s+c
      else
        break; // char after number - the rest is ignored
    end;
  end;
  result := s;
end;

// ==================================================
// Expand numbers in the string by adding leading zeroes
// ==================================================
function ExpandNumber(ss : string): string;
var s : string;
    i, val : integer;
    c : char;
    number : boolean;
begin
  s := '';
  val := 0;
  number := false;
  for i:=1 to Length(ss) do begin
    c := ss[i];
    // ----------------------------
    // if digit
    // ----------------------------
    if c in ['0'..'9'] then begin
      if number then
        val := val*10 + ord(c) - ord('0')
      else begin
        val := ord(c) - ord('0');
        number := true;
      end;
    // ----------------------------
    // if letter
    // ----------------------------
    end else begin
      if number then begin
        number := false;  // finish number
        if val < 10000000 then
          s := s + '0';
        if val < 1000000 then
          s := s + '0';
        if val < 100000 then
          s := s + '0';
        if val < 10000 then
          s := s + '0';
        if val < 1000 then
          s := s + '0';
        if val < 100 then
          s := s + '0';
        if val < 10 then
          s := s + '0';
        s := s + IntToStr(val);
      end else
        s := s + c;
    end;
  end;
  result := s;
end;

// ==================================================
// Custom sort compare
// ==================================================
function MySort(List: TStringList; Index1, Index2: Integer): Integer;
var s1, s2 : string;
begin
  s1 := ExpandNumber(List.Strings[Index1]);
  s2 := ExpandNumber(List.Strings[Index2]);
  Result:= AnsiCompareText(s1,s2);
end;

// ##################################################
{ TKicadSchBOM }
// ##################################################

// ==================================================
// Clear component
// ==================================================
procedure TKicadSchBOM.ClrComp;
var i : integer;
begin
  FComp.Clear;
  for i:=0 to ord(sepOther) do
    FComp.Add('');
end;

// ==================================================
// Add component to BOM
// ==================================================
procedure TKicadSchBOM.AddComp;
var i : integer;
    s : string;
    tmp : TStringList;
begin
  if Trim(FComp.Strings[0]) <> '' then begin
    s := '';
    tmp := TStringList.Create;
    for i:=0 to ord(sepOther) do begin
      if i <= ord(sepNote) then begin
        if i = ord(sepPrice) then
          s := s + FComp.Strings[i]               // price is number
        else
          s := s + '"' + FComp.Strings[i] + '"'; // other cells are strings
        if i < ord(sepNote) then
          s := s + char(9);         // add TABs
      end;
      tmp.Add(FComp.Strings[i]);
    end;
    self.AddObject(s, (tmp as TObject));
    inc(comp_cnt);
  end;
end;

// ==================================================
// Separate and exclude prefix from footprint string
// ==================================================
function TKicadSchBOM.ExFootprintPrefix(val: string): string;
var i : integer;
    s : string;
    c : char;
begin
  s := '';
  result := val;
  for i:=1 to Length(val)-1 do begin
    c := val[i];
    s := s + UpperCase(c);
    if FExcludedFootPrefix.IndexOf(s) >= 0 then begin
      result := Copy(val, i+1, Length(val));
      break;
    end;
  end;
end;

// ==================================================
// Fill specified component field with value
// ==================================================
procedure TKicadSchBOM.FillCompField(fld: integer; val: string);
var n : integer;
    refname : string;
begin
  n := 0;
  case fld of
    // ------------------------------
    ord(sepRef): begin
      if val[1] = '#' then  // power symbol
        comp_mode := false
      else if FRefList.IndexOf(AnsiUpperCase(val)) >= 0 then // this RefDes already exists
        comp_mode := false
      else begin
        refname := ExtractRefDesName(val,n);
        if FExcludedRef.IndexOf(AnsiUpperCase(refname)) >= 0 then
          comp_mode := false
        else begin
          FRefList.Add(AnsiUpperCase(val));
          FComp.Strings[ord(sepRef)]  := val;
          FComp.Strings[ord(sepRefName)] := refname;
          FComp.Strings[ord(sepRefIndex)] := IntToStr(n);
        end;
      end;
    end;
    // ------------------------------
    ord(sepValue): begin
      FComp.Strings[ord(sepValue)] := val;
    end;
    // ------------------------------
    ord(sepFootprint): begin
      FComp.Strings[ord(sepFootprint)] := ExFootprintPrefix(val);
    end;
    // ------------------------------
    ord(sepBrand): begin
       FComp.Strings[ord(sepBrand)] := val;
    end;
    // ------------------------------
    ord(sepOrder): begin
       FComp.Strings[ord(sepOrder)] := val;
    end;
    // ------------------------------
    ord(sepSupplier): begin
       FComp.Strings[ord(sepSupplier)] := val;
    end;
    // ------------------------------
    ord(sepPrice): begin
       FComp.Strings[ord(sepPrice)] := val;
    end;
    // ------------------------------
    ord(sepNote): begin
       FComp.Strings[ord(sepNote)] := val;
    end;
    // ------------------------------
    ord(sepRefUnit): begin
       FComp.Strings[ord(sepRefUnit)] := val;
    end;
  end;
end;

// ==================================================
// Fill component fields from parsed string (FTmp)
// ==================================================
procedure TKicadSchBOM.FillComp;
var fld : string;
    fno : string;
    fval : string;
    fnm : string; // field name
begin
  if FTmp.Count < 3 then
      exit;
  fld := FTmp.Strings[0];
  fno := FTmp.Strings[1];
  fval := Ftmp.Strings[2];
  // -----------------------------------
  // if it is a unit (part) string
  // -----------------------------------
  if fld = 'U' then begin
     FillCompField(ord(sepRefUnit), fno);
  // -----------------------------------
  // if it is a filed string
  // -----------------------------------
  end else if fld = 'F' then begin
    if FTmp.Count < 10 then
      exit
    else begin
      // ---------------------------
      // pre-defined KiCad fields
      // ---------------------------
      if fno = '0' then begin
        FillCompField(ord(sepRef), fval);
      end else if fno = '1' then begin
        FillCompField(ord(sepValue), fval);
      end  else if fno = '2' then begin
        FillCompField(ord(sepFootprint), fval);
      // ---------------------------
      // user-defined fields
      // ---------------------------
      end else begin
        fnm := AnsiLowerCase(FTmp.Strings[FTmp.Count-1]); // last
        if (fnm = 'brand') or (fnm = 'manuf') or (fnm = 'manufacturer') then begin
          FillCompField(ord(sepBrand), fval);
        end else if (fnm = 'ord') or (fnm = 'order') or (fnm = 'ordering') or (fnm = 'order code')  or (fnm = 'ordering code') then begin
          FillCompField(ord(sepOrder), fval);
        end else if (fnm = 'suppl') or (fnm = 'supplier') or (fnm = 'suppliers') then begin
          FillCompField(ord(sepSupplier), fval);
        end else if (fnm = 'price') or (fnm = 'cost') or (fnm = '$') then begin
          FillCompField(ord(sepPrice), fval);
        end else if (fnm = 'note') or (fnm = 'notes') or (fnm = 'comment') or (fnm = 'comments') then begin
          FillCompField(ord(sepNote), fval);
        end else if (fnm = 'inventory') or (fnm = 'inv') or (fnm = 'inv #') or (fnm = 'inv no') or
            (fnm = '#') or (fnm = 'stock') or (fnm = 'stock no') then begin
          FillCompField(ord(sepInv), fval);
        end;
      end;
    end;
  end;
end;

// ==================================================
// Read source file
// ==================================================
procedure TKicadSchBOM.ReadFile(fn:string; appnd : boolean);
begin
  // --------------------------
  // Load new or append existing
  // --------------------------
  if appnd = false then begin
    Reset;
  end;
  // --------------------------
  // If file has not been loaded yet
  // --------------------------
  FSrc.LoadFromFile(fn);
  ParseFile;
  self.Sorted := false;
  self.CustomSort(@MySort);
  self.MakePartList;
end;

// ==================================================
// Save as CSV file
// ==================================================
procedure TKicadSchBOM.SaveCSV(fn: string;  BOMname: string; SG : TStringGrid);
var cf : TStringList;
    i : integer;
    s : string;
begin
  cf := TStringList.Create;
  // ---------------------
  // Header block
  // ---------------------
  cf.Add(BOMname);         // row #1 - BOM name
  cf.Add('');              // row #2 - spare
  // ---------------------
  // Field names - row #3
  // ---------------------
  s := '';
  for i:=0 to SG.Columns.Count-1 do begin
    s := s + SG.Columns[i].Title.Caption;
    if i < SG.Columns.Count-1 then
      s := s + char(9);
  end;
  cf.Add(s);
   // ---------------------
   // Data starts from row #4
   // ---------------------
   for i:=0 to Count-1 do
     cf.Add(Strings[i]);
   cf.SaveToFile(fn);
 cf.Free;
end;

// ==================================================
// Parse file
// ==================================================
procedure TKicadSchBOM.ParseFile;
var i : integer;
    s : string;
begin
  comp_mode := false;
  for i:=0 to FSrc.Count-1 do begin
    s := FSrc.Strings[i];
    ParseString(s);
    if FTmp.Count > 0 then begin
      // ----------------------------
      // component starts
      // ----------------------------
      if AnsiLowerCase(FTmp.Strings[0]) = '$comp' then begin
        comp_mode := true;
        ClrComp;
      // ----------------------------
      // component finished
      // ----------------------------
      end else if AnsiLowerCase(FTmp.Strings[0]) = '$endcomp' then begin
        comp_mode := false;
        AddComp;
      end else begin
        // ----------------------------
        // fill component
        // ----------------------------
        if comp_mode then begin
          FillComp;
        end;
      end;
    end;
  end;
end;

// ==================================================
// Parse string - separate into words and store words in Ftmp
// ==================================================
procedure TKicadSchBOM.ParseString(ss: string);
var i : integer;
    s : string;
    c : char;
    was_backslash : boolean;
begin
  quote_mode := false;
  was_backslash := false;
  Ftmp.Clear;
  s := '';
  for i:=1 to length(ss) do begin
    c := ss[i];
    // --------------------------
    // quoted text
    // --------------------------
    if quote_mode then begin
      if c = char(9) then         // replace TABs
        c := ' ';
      // ------------------
      // it is a second char of esc-pair
      // ------------------
      if was_backslash then begin
        if (c = '\') or (c = '"') then
          s := s + c              // convert legal esc-pair into a single char
        else
          s := s + '\' + c;       // it should not happen: incomplete esc-pair
        was_backslash := false;   // esc-pair finished
      // ------------------
      // normal (eg it is not an esc-pair)
      // ------------------
      end else begin
        if c = '"' then begin     // closing quote is a separator
          quote_mode := false;
          Ftmp.Add(s);            // even if string is empty
          s := '';
        end else if c = '\' then  // esc-pair starts
          was_backslash := true
        else
          s := s + c;
      end;
      // --------------------------
    // without quote
    // --------------------------
    end else begin
      if c = '"' then begin
        quote_mode := true;        // quoted text started
        s := '';
      end else begin
        if ord(c) <= ord(' ') then begin  // if a separator char
          if s <> '' then
            Ftmp.Add(s);
          s := '';
        end else
          s := s + c;
      end;
    end;
  end;
  if s <> '' then
    Ftmp.Add(s);
end;

// ==================================================
// Make grouped BOM out of separate items
// ==================================================
procedure TKicadSchBOM.MakePartList;
var i : integer;
    comp : TStringList;
begin
  FPartsList.Reset;
  for i:=0 to Count-1 do begin
    comp := Objects[i] as TStringList;
    FPartsList.AddComp(comp);
  end;
  FPartsList.MakeCSV;
end;

// ==================================================
// Find similar components and update their fields
// ==================================================
procedure TKicadSchBOM.UpdateFields;
var i, j : integer;
    refnm, val, footprint, order, brand, inv : string;
    refnmX, valX, footprintX, orderX, brandX, invX : string;
    sl, slX : TStringList;
begin
  for i := 0 to Count-2 do begin
    sl := (self.Objects[i] as TStringList);
    refnm      := sl.Strings[ord(sepRefName)];
    val        := sl.Strings[ord(sepValue)];
    footprint  := sl.Strings[ord(sepFootprint)];
    order      := sl.Strings[ord(sepOrder)];
    brand      := sl.Strings[ord(sepBrand)];
    inv        := sl.Strings[ord(sepInv)];
    for j := i+1 to Count-2 do begin
      slX := (self.Objects[j] as TStringList);
      refnmX      := slX.Strings[ord(sepRefName)];
      valX        := slX.Strings[ord(sepValue)];
      footprintX  := slX.Strings[ord(sepFootprint)];
      orderX      := slX.Strings[ord(sepOrder)];
      brandX      := slX.Strings[ord(sepBrand)];
      invX        := slX.Strings[ord(sepInv)];
      if (refnm = refnmX) and (val = valX) and (val <> 'DNP') then begin
        if (footprint = '') then begin
          sl.Strings[ord(sepFootprint)] := footprintX;
          footprint := footprintX;
        end;
        if (footprintX = '') then begin
          slX.Strings[ord(sepFootprint)] := footprint;
          footprintX := footprint;
        end;
        if (footprint = footprintX) then begin
          if (order = '') then
            sl.Strings[ord(sepOrder)] := orderX;
          if (orderX = '') then
            slX.Strings[ord(sepOrder)] := order;
          if (brand = '') then
            sl.Strings[ord(sepBrand)] := brandX;
          if (brandX = '') then
            slX.Strings[ord(sepBrand)] := brand;
          if (invX = '') then
            slX.Strings[ord(sepInv)] := inv;
        end;
      end;
    end;
  end;
end;

// ==================================================
// Copy BOM to StringGrid
// ==================================================
procedure TKicadSchBOM.ToStringGrid(var SG: TStringGrid);
var i, j : integer;
    sl : TStringList;
    s : string;
begin
  SG.RowCount := Count;
  for i:=0 to Count-1 do begin
    sl := (self.Objects[i] as TStringList);
    for j:=0 to ord(sepNote) do begin
      if j < sl.Count then begin
        s := sl.Strings[j];
        SG.Cells[j,i] := s;
      end else
        SG.Cells[j,i] := '~';
    end;
  end;
end;

// ==================================================
// Reset
// ==================================================
procedure TKicadSchBOM.Reset;
var i : integer;
begin
  quote_mode := false;
  comp_mode := false;
  comp_cnt := 0;
  for i:=0 to Count-1 do
     Objects[i].Free;
  Clear;
  ClrComp;
  FPartsList.Reset;
end;

// ==================================================
// Create
// ==================================================
constructor TKicadSchBOM.Create;
begin
  FExcludedRef := TStringList.Create;
  FExcludedFootPrefix := TStringList.Create;
  FSrc := TStringList.Create;
  FTmp := TStringList.Create;
  FComp := TStringList.Create;
  FPartsList := TKicadSchPL.Create;
  FRefList := TStringList.Create;
  Reset;
end;

// ==================================================
// Destroy
// ==================================================
destructor TKicadSchBOM.Destroy;
begin
  Reset;
  FSrc.Free;
  FTmp.Free;
  FComp.Free;
  FExcludedRef.Free;
  FExcludedFootPrefix.Free;
  FPartsList.Free;
  FRefList.Free;
  inherited Destroy;
end;

end.

