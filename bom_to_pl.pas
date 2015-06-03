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
  sep_bom_fields = (  sepRef, sepValue, sepFootprint, sepOrder, sepBrand, sepSupplier,
            sepPrice, sepNote, sepRefName, sepRefIndex, sepOther );
  gr_bom_fields = (  grRef, grValue, grFootprint, grOrder, grBrand, grSupplier,
            grPrice, grQty, grRefName, grOther );

{ TKicadSchPL }

TKicadSchPL = class(TStringList)
private
  same_comp : array [0..$FF] of integer;
  procedure AddRefDes(i : integer; ss : string);
  procedure IncQty(i : integer);
  procedure SetField(i, fld : integer; val : string);
  function GetField(i, fld : integer) : string;
  function FindField(fld : integer; val : string; fp : string; refname : string) : integer;
public
  procedure ToStringGrid(SG : TStringGrid);
  procedure MakeCSV;
  procedure AddComp(comp : TStringList);
  procedure Reset;
  constructor Create;
  destructor Destroy; override;
end;

// ###########################################################################
implementation
// ###########################################################################

{ TKicadSchPL }

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
// Write to specified field
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
// Copy parts list to string grid
// ==================================================
procedure TKicadSchPL.ToStringGrid(SG: TStringGrid);
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
// Read fields and make CSV
// ==================================================
procedure TKicadSchPL.MakeCSV;
var sl : TStringList;
    i,j : integer;
    s : string;
begin
  for i:= 0 to Count-1 do begin
    sl := Objects[i] as TStringList;
    s := '';
    for j:=0 to ord(grQty) do begin
      if j < ord(grPrice) then begin
        s := s + '"' +  sl.Strings[j] + '"'; // first cells are text
        s := s + char(9);                    // add TABs
      end else begin
        s := s + sl.Strings[j];              // last 2 cells are numbers
        if j < ord(grQty) then
           s := s + char(9);                 // add TABs except the last cell
      end;
    end;
    Strings[i] := s;
  end;
end;

// ==================================================
// Add component
// ==================================================
procedure TKicadSchPL.AddComp(comp: TStringList);
var tmp : TStringList;
    i, j, cnt : integer;
    accept : boolean;
    refname, refdes, value, footprint, order, brand, supplier : string;
    footprintX, orderX, brandX, supplierX : string;
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
      order     := AnsiUpperCase(comp.Strings[ord(sepOrder)]);
      brand     := AnsiUpperCase(comp.Strings[ord(sepBrand)]);
      supplier  := AnsiUpperCase(comp.Strings[ord(sepSupplier)]);
      if refname = 'T' then
        cnt := 0;
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
             i := same_comp[j];
             footprintX := AnsiUpperCase(GetField(i, ord(grFootprint)));
             orderX     := AnsiUpperCase(GetField(i, ord(grOrder)));
             brandX     := AnsiUpperCase(GetField(i, ord(grBrand)));
             supplierX   := AnsiUpperCase(GetField(i, ord(grSupplier)));
             accept := true;
             if (footprint <> '') and (footprintX <> '') and (footprint <> footprintX) then
               accept := false
             else if (order <> '') and (orderX <> '') and (order <> orderX) then
               accept := false
             else if (brand <> '') and (brandX <> '') and (brand <> brandX) then
               accept := false;
             if accept then
               break;
           end;
         end;
         // ------------------------
         // If it is the same component
         // ------------------------
         if accept then begin
           IncQty(i);        // increment Q-ty
           AddRefDes(i, refdes);
           // update empty fields
           if (footprintX = '') and (footprint <> '') then
              SetField(i, ord(grFootprint), comp.Strings[ord(sepFootprint)]);
           if (orderX = '') and (order <> '') then
              SetField(i, ord(grOrder), comp.Strings[ord(sepOrder)]);
           if (brandX = '') and (brand <> '') then
              SetField(i, ord(grBrand), comp.Strings[ord(sepBrand)]);
           if (supplierX = '') and (supplier <> '') then
              SetField(i, ord(grSupplier), comp.Strings[ord(sepSupplier)]);
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
           tmp.Add(comp.Strings[ord(sepOrder)]);
           tmp.Add(comp.Strings[ord(sepBrand)]);
           tmp.Add(comp.Strings[ord(sepSupplier)]);
           tmp.Add(comp.Strings[ord(sepPrice)]);
           tmp.Add('1');  // Q-ty = 1
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
