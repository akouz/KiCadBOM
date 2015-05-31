unit sch_to_bom;

{
* Author    A.Kouznetsov
* Rev       1.0 dated 30/5/2015
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
    quot_mode : boolean;      // quotation mode
    comp_mode : boolean;      // component mode
    comp_cnt : integer;
    Ftmp : TStringList;
    FComp: TStringList;  // current component
    FSrc : TStringList;
    FFileNames : TStringList;
    function is_separator(c: char): boolean;
    function is_excluded_ref(ss:string) : boolean;
    procedure ClrComp;
    procedure AddComp;
    procedure FillComp;
    procedure ParseFile;
    procedure ParseString(ss:string);
public
    FExcludedRef : TStringList;
    FPartsList : TKicadSchPL;
    procedure MakePartList;
    procedure UpdateFields;
    procedure ReadFile(fn:string; appnd : boolean);
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

{ TKicadSchBOM }

// ==================================================
// Check if char is a separator
// ==================================================
function TKicadSchBOM.is_separator(c: char): boolean;
begin
  result := false;
  if quot_mode then begin
     if c = '"' then begin
       quot_mode := false;
       result := true;
     end;
  end else begin
    if ord(c) <= ord(' ') then
      result := true
    else begin
      if c = '"' then begin
        quot_mode := true;
        result := true;
      end;
    end;
  end;
end;

// ==================================================
// Check if specified sepRef should be excluded from the list
// ==================================================
function TKicadSchBOM.is_excluded_ref(ss: string): boolean;
var s : string;
    c : char;
    i : integer;
begin
  result := false;
  s := '';
  for i := 1 to Length(ss) do begin
    c := ss[i];
    if c in ['0'..'9'] then
      break
    else
      s := s+c;
  end;
  if FExcludedRef.IndexOf(AnsiUpperCase(s)) >= 0 then
    result := true;
end;

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
    if is_excluded_ref(FComp.Strings[ord(sepRefName)]) = false then begin
      s := '';
      tmp := TStringList.Create;
      for i:=0 to ord(sepOther)-1 do begin
        s := s + FComp.Strings[i] + ' ; ';
        tmp.Add(FComp.Strings[i]);
      end;
      self.AddObject(s, (tmp as TObject));
      inc(comp_cnt);
    end;
  end;
end;

// ==================================================
// Fill component fields from parsed string (FTmp)
// ==================================================
procedure TKicadSchBOM.FillComp;
var fnm : string; // field name
    fno : string;
    fval : string;
    n : integer;
begin
  n := 0;
  // -----------------------------------
  // if it is a filed string
  // -----------------------------------
  if FTmp.Count >= 10 then begin
    if FTmp.Strings[0] = 'F' then begin
      fval := Ftmp.Strings[2];
      // ---------------------------
      // pre-defined fields
      // ---------------------------
      fno := FTmp.Strings[1];
      if fno = '0' then begin
        if fval[1] = '#' then  // power symbol
          comp_mode := false
        else begin
          FComp.Strings[ord(sepRef)]  := fval;
          FComp.Strings[ord(sepRefName)] := ExtractRefDesName(fval,n);
        end;
      end else if fno = '1' then begin
        FComp.Strings[ord(sepValue)] := fval
      end  else if fno = '2' then begin
        FComp.Strings[ord(sepFootprint)] := fval
      end
      // ---------------------------
      // user-defined fields
      // ---------------------------
      else if FTmp.Count > 10 then begin
        fnm := AnsiLowerCase(FTmp.Strings[10]);
        if (fnm = 'brand') or (fnm = 'manuf') or (fnm = 'manufacturer') then begin
          FComp.Strings[ord(sepBrand)] := fval
        end else if (fnm = 'order') then begin
          FComp.Strings[ord(sepOrder)] := fval;
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
  if FFileNames.IndexOf(fn) < 0 then begin
    FFileNames.Add(fn);
    FSrc.LoadFromFile(fn);
    ParseFile;
    self.Sorted := false;
    self.CustomSort(@MySort);
    self.MakePartList;
  end;
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
        if comp_mode and (FTmp.Strings[0] = 'F') then begin
          FillComp;
        end;
      end;
    end;
  end;
end;

// ==================================================
// Parse string
// ==================================================
procedure TKicadSchBOM.ParseString(ss: string);
var i : integer;
    s : string;
    c : char;
begin
  quot_mode := false;
  Ftmp.Clear;
  s := '';
  for i:=1 to length(ss) do begin
    c := ss[i];
    if is_separator(c) then begin
      s := Trim(s);
      if s <> '' then begin
        Ftmp.Add(s);
      end;
      s := '';
    end else
      s := s + c;
  end;
  s := Trim(s);
  if s <> '' then begin
    Ftmp.Add(s);
  end;
end;

// ==================================================
// Make parts list out of BOM
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
  FPartsList.SaveToFile('test.txt');
end;

// ==================================================
// Find similar components and update their fields
// ==================================================
procedure TKicadSchBOM.UpdateFields;
var i, j : integer;
    refnm, val, footprint, order, brand : string;
    refnmX, valX, footprintX, orderX, brandX : string;
    sl, slX : TStringList;
begin
  for i := 0 to Count-2 do begin
    sl := (self.Objects[i] as TStringList);
    refnm      := sl.Strings[ord(sepRefName)];
    val        := sl.Strings[ord(sepValue)];
    footprint  := sl.Strings[ord(sepFootprint)];
    order      := sl.Strings[ord(sepOrder)];
    brand      := sl.Strings[ord(sepBrand)];
    for j := i+1 to Count-2 do begin
      slX := (self.Objects[j] as TStringList);
      refnmX      := slX.Strings[ord(sepRefName)];
      valX        := slX.Strings[ord(sepValue)];
      footprintX  := slX.Strings[ord(sepFootprint)];
      orderX      := slX.Strings[ord(sepOrder)];
      brandX      := slX.Strings[ord(sepBrand)];
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
begin
  SG.RowCount := Count;
  for i:=0 to Count-1 do begin
    sl := (self.Objects[i] as TStringList);
    for j:=0 to 7 do begin
      if j < sl.Count then
        SG.Cells[j,i] := sl.Strings[j]
      else
        SG.Cells[j,i] := '';
    end;
  end;
end;

// ==================================================
// Reset
// ==================================================
procedure TKicadSchBOM.Reset;
var i : integer;
begin
  quot_mode := false;
  comp_mode := false;
  comp_cnt := 0;
  for i:=0 to Count-1 do
     Objects[i].Free;
  Clear;
  ClrComp;
  FPartsList.Reset;
  FFileNames.Clear;
end;

// ==================================================
// Create
// ==================================================
constructor TKicadSchBOM.Create;
begin
  FExcludedRef := TStringList.Create;
  FSrc := TStringList.Create;
  FTmp := TStringList.Create;
  FComp := TStringList.Create;
  FFileNames := TStringList.Create;
  FPartsList := TKicadSchPL.Create;
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
  FFileNames.Free;
  FExcludedRef.Free;
  FPartsList.Free;
  inherited Destroy;
end;

end.

