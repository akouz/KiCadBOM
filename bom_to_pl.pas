unit bom_to_pl;
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
  Classes, SysUtils, Grids;

type
  sep_bom_fields = (  sepRef, sepValue, sepFootprint, sepInv, sepOrder, sepBrand, sepSupplier,
            sepPrice, sepNote,  sepRefName, sepRefIndex,  sepRefUnit, sepOther );
  gr_bom_fields = (  grRef, grValue, grFootprint,  grInv, grOrder, grBrand, grSupplier,
            grPrice, grQty, grNote, grRefName, grOther );

{ TKicadSchPL }

TKicadSchPL = class(TStringList)
private
  same_comp : array [0..$FF] of integer;
  procedure AddRefDes(i : integer; ss : string);
  procedure IncQty(i : integer);
  procedure AddField(i, fld : integer; val : string);
  procedure SetField(i, fld : integer; val : string);
  function GetField(i, fld : integer) : string;
  function FindField(fld : integer; val : string; fp : string; refname : string) : integer;
public
  procedure ToStringGrid(SG : TStringGrid);
  function GetRef(Row : integer) : string;
  procedure MakeCSV;
  procedure SaveCSV(fn:string;  BOMname: string; SG : TStringGrid);
  procedure AddComp(comp : TStringList);
  procedure Reset;
  constructor Create;
  destructor Destroy; override;
end;

// ###########################################################################
implementation
// ###########################################################################

// ##################################################
{ TKicadSchPL }
// ##################################################

// ==================================================
// Add RefDes
// ==================================================
procedure TKicadSchPL.AddRefDes(i: integer; ss : string);
var sl : TStringList;
    s : string;
begin
  if (i >= 0) and (i < Count) then begin
    sl := Objects[i] as TStringList;
    s := sl.Strings[ord(grRef)];  // list of RefDes
    sl.Strings[ord(grRef)] := s +', ' + ss;
  end;
end;

// ==================================================
// Increment Q-ty field
// ==================================================
procedure TKicadSchPL.IncQty(i: integer);
var sl : TStringList;
    n : integer;
begin
  if (i >= 0) and (i < Count) then begin
    sl := Objects[i] as TStringList;
    n := StrToIntDef(sl.Strings[ord(grQty)], 1);
    inc(n);
    sl.Strings[ord(grQty)] := IntToStr(n);
  end;
end;

// ==================================================
// Add string to the specified field
// ==================================================
// do not add if val already is a part of the field
procedure TKicadSchPL.AddField(i, fld: integer; val: string);
var sl : TStringList;
    j,len : integer;
    s, ss : string;
    same : boolean;
begin
  if (i >= 0) and (i < Count) and (val <> '') then begin
    sl := Objects[i] as TStringList;
    if (fld >= 0) and (fld < sl.Count) then begin
       len := Length(val);
       ss := sl.Strings[fld];
       // ------------------------
       // parse existing field
       // ------------------------
       if len < Length(ss) then begin
          same := false;
          for j := 1 to (Length(ss) - len + 1) do begin
            s := Copy(ss, j, len);    // select substring of the same length
            if s = val then begin     // substring matches
              same := true;           // do not add val
              break;
            end;
          end;
          if not same then
            sl.Strings[fld] := ss + ', ' + val;
       // ------------------------
       // just add a new value
       // ------------------------
       end else if ((len = Length(ss)) and (val <> ss)) or  (len > Length(ss)) then begin
         if ss <> '' then
           sl.Strings[fld] := ss + '; ' + val
         else
           sl.Strings[fld] := val;
       end;

    end;
  end;
end;

// ==================================================
// Write to the specified field
// ==================================================
procedure TKicadSchPL.SetField(i, fld: integer; val: string);
var sl : TStringList;
begin
  if (i >= 0) and (i < Count) then begin
    sl := Objects[i] as TStringList;
    if (fld >= 0) and (fld < sl.Count) then
      sl.Strings[fld] := val;
  end;
end;

// ==================================================
// Read specified field
// ==================================================
function TKicadSchPL.GetField(i, fld: integer): string;
var sl : TStringList;
begin
  result := '';
  if (i >= 0) and (i < Count) then begin
    sl := Objects[i] as TStringList;
    if (fld >= 0) and (fld < sl.Count) then
      result := sl.Strings[fld];
  end;
end;

// ==================================================
// Find specified field, return count
// ==================================================
function TKicadSchPL.FindField(fld: integer; val: string; fp : string; refname : string): integer;
var sl : TStringList;
    i : integer;
    s : string;
begin
  result := 0;
  if Trim(val) <> '' then begin   // if valid
    for i:=0 to Count-1 do begin
      sl := Objects[i] as TStringList;
      s := sl.Strings[fld];
      if AnsiUpperCase(s) =  AnsiUpperCase(val) then begin
        // ------------------------------
        // also reference type must be the same
        // ------------------------------
        s := sl.Strings[ord(grRefName)];
        if AnsiUpperCase(s) = AnsiUpperCase(refname) then
          // ------------------------------
          // also footprints must be the same
          // ------------------------------
          s := sl.Strings[ord(grFootprint)];
          if ( AnsiUpperCase(s) = AnsiUpperCase(fp)) then begin
            same_comp[result] := i;
            if result < $FF  then
              inc(result);
          // ------------------------------
          // or one of footprints should not be defined
          // ------------------------------
          end else if (fp = '') or (sl.Strings[ord(grFootprint)] = '') then begin
            same_comp[result] := i;
            if result < $FF  then
              inc(result);
          end;
      end;
    end;
  end;
end;

// ==================================================
// Convert RefDes list into [first..last]
// ==================================================
function convert_list_to_first_last(ss : string) : string;
var s , s1, sx: string;
    c : char;
    i, cnt : integer;
begin
  cnt := 0;
  s := '';
  s1 := '';
  sx := '';
  // ----------------------
  // select RefDes
  // ----------------------
  for i := 1 to Length(ss) do begin
    c := ss[i];
    if (ord(c) <= ord(' ')) or (c = ',') then begin // separator
      if (s <> '') then begin      // if valid
        if cnt = 0 then begin
          s1 := s;
        end else begin
          sx := s;
        end;
        inc(cnt);
      end;
      s := '';
    end else
      s := s + c;
  end;
  // ----------------------
  // Last RefDes
  // ----------------------
  if s <> '' then begin
    sx := s;
    inc(cnt);
  end;
  // ----------------------
  // result
  // ----------------------
  if cnt < 3 then
    result := ss
  else
    result := '{' + s1 + ', ..., ' + sx + '}';
end;

// ==================================================
// Copy parts list to string grid
// ==================================================
procedure TKicadSchPL.ToStringGrid(SG: TStringGrid);
var i, j : integer;
    sl : TStringList;
begin
  SG.RowCount := Count;
  // ----------------------------
  // For every row
  // ----------------------------
  for i:=0 to Count-1 do begin
    sl := (self.Objects[i] as TStringList);
    // ----------------------------
    // For visible columns, up to Q-ty
    // ----------------------------
    for j:=0 to ord(grNote) do begin
      if j < sl.Count then begin
        if j = 0 then  begin
          SG.Cells[j,i] := convert_list_to_first_last(sl.Strings[j]);  // RefDes
        end else begin
          SG.Cells[j,i] := sl.Strings[j];
        end;
      end else begin
        SG.Cells[j,i] := '';
      end;
    end;
  end;
end;

// ==================================================
// Get RefDes
// ==================================================
function TKicadSchPL.GetRef(Row: integer): string;
var sl : TStringList;
begin
  result := '';
  if (Row >=0) and (Row < Count) then begin
    sl := Objects[Row] as TStringList;
    result := sl.Strings[0];
  end;
end;

// ==================================================
// Read fields and make CSV
// ==================================================
procedure TKicadSchPL.MakeCSV;
var sl : TStringList;
    i,j, last : integer;
    s : string;
begin
  last := ord(grNote);                       // last cell in a row
  // --------------------------------
  // For every row
  // --------------------------------
  for i:= 0 to Count-1 do begin
    sl := Objects[i] as TStringList;
    s := '';
    // --------------------------------
    // For columns up to Notes
    // --------------------------------
    for j:=0 to last do begin
      if (j = ord(grPrice)) or (j = ord(grQty)) then begin
        s := s + sl.Strings[j];              // numbers
        if j < last then
           s := s + char(9);                 // add TABs except the last cell
      end else begin
        s := s + '"' +  sl.Strings[j] + '"'; // text cells
        if j < last then
          s := s + char(9);                  // add TABs except the last cell
      end;
    end;
    Strings[i] := s;
  end;
end;

// ==================================================
// Save as CSV file
// ==================================================
procedure TKicadSchPL.SaveCSV(fn: string;  BOMname: string; SG: TStringGrid);
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
// Add component
// ==================================================
procedure TKicadSchPL.AddComp(comp: TStringList);
var tmp : TStringList;
    j, cnt, same : integer;
    accept : boolean;
    refname, refdes, value, footprint, order, brand, supplier, inv : string;
    footprintX, orderX, brandX, supplierX, invX: string;
begin
  // --------------------------------
  // RefDes and Value must exist
  // --------------------------------
  if (comp.Strings[ord(sepRef)] <> '') then begin
    if (comp.Strings[ord(sepValue)] <> '') then begin
      refname   := AnsiUpperCase(comp.Strings[ord(sepRefName)]);
      refdes    := AnsiUpperCase(comp.Strings[ord(sepRef)]);
      value     := AnsiUpperCase(comp.Strings[ord(sepValue)]);
      footprint := AnsiUpperCase(comp.Strings[ord(sepFootprint)]);
      inv       := AnsiUpperCase(comp.Strings[ord(sepInv)]);
      order     := AnsiUpperCase(comp.Strings[ord(sepOrder)]);
      brand     := AnsiUpperCase(comp.Strings[ord(sepBrand)]);
      supplier  := AnsiUpperCase(comp.Strings[ord(sepSupplier)]);
      // ----------------------------
      // If value defined
      // ----------------------------
         cnt := FindField(ord(grValue), value, footprint, refname);
         // ------------------------
         // If such a field already exists - compare fields
         // ------------------------
         accept := false;
         if cnt > 0 then begin
           for j:=0 to cnt-1 do begin
             same := same_comp[j];
             footprintX  := AnsiUpperCase(GetField(same, ord(grFootprint)));
             invX        := AnsiUpperCase(GetField(same, ord(grInv)));
             orderX      := AnsiUpperCase(GetField(same, ord(grOrder)));
             brandX      := AnsiUpperCase(GetField(same, ord(grBrand)));
             supplierX   := AnsiUpperCase(GetField(same, ord(grSupplier)));
             accept := true;
             if (inv <> '') and (invX <> '') and (inv <> invX) then
               accept := false
             else if (order <> '') and (orderX <> '') and (order <> orderX) then
               accept := false
             else if (brand <> '') and (brandX <> '') and (brand <> brandX) then
               accept := false
             else if (inv <> '') and (invX <> '') and (inv <> invX) then
               accept := false;
             if accept then
               break;
           end;
         end;
         // ------------------------
         // If it is the same component
         // ------------------------
         if accept then begin
           IncQty(same);        // increment Q-ty
           AddRefDes(same, refdes);
           // update empty fields
           if (footprintX = '') and (footprint <> '') then
              SetField(same, ord(grFootprint), comp.Strings[ord(sepFootprint)]);
           if (invX = '') and (inv <> '') then
              SetField(same, ord(grInv), comp.Strings[ord(sepInv)]);
           if (orderX = '') and (order <> '') then
              SetField(same, ord(grOrder), comp.Strings[ord(sepOrder)]);
           if (brandX = '') and (brand <> '') then
              SetField(same, ord(grBrand), comp.Strings[ord(sepBrand)]);
           if (supplierX = '') and (supplier <> '') then
              SetField(same, ord(grSupplier), comp.Strings[ord(sepSupplier)]);
           AddField(same, ord(grNote),  comp.Strings[ord(sepNote)]);
         // ------------------------
         // If it is a new component
         // ------------------------
         end else begin
{           pl_fields = (  grRef, grValue, grFootprint, grOrder, grBrand, grSupplier,
                     grPrice, grQty, grRefName, grOther ); }
           tmp := TStringList.Create;
           tmp.Add(comp.Strings[ord(sepRef)]);
           tmp.Add(comp.Strings[ord(sepValue)]);
           tmp.Add(comp.Strings[ord(sepFootprint)]);
           tmp.Add(comp.Strings[ord(sepInv)]);
           tmp.Add(comp.Strings[ord(sepOrder)]);
           tmp.Add(comp.Strings[ord(sepBrand)]);
           tmp.Add(comp.Strings[ord(sepSupplier)]);
           tmp.Add(comp.Strings[ord(sepPrice)]);
           tmp.Add('1');  // Q-ty = 1
           tmp.Add(Trim(comp.Strings[ord(sepNote)]));
           tmp.Add(comp.Strings[ord(sepRefName)]);
           tmp.Add('');
           self.AddObject('',tmp as TObject);
         end;
      end;
  end;
end;

// ==================================================
// Reset
// ==================================================
procedure TKicadSchPL.Reset;
var i : integer;
begin
  for i:=0 to Count-1 do
//    if i = ord(sepNote) then
//       (Objects[i] as TStringList).Free;
    Objects[i].Free;
  Clear;
end;

// ==================================================
// Create
// ==================================================
constructor TKicadSchPL.Create;
begin

end;

// ==================================================
// Destroy
// ==================================================
destructor TKicadSchPL.Destroy;
begin
  Reset;
  inherited Destroy;
end;

end.

