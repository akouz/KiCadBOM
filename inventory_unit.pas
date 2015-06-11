unit inventory_unit;

{
* Author    A.Kouznetsov
* Rev       1.02 dated 4/6/2015
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
  Classes, SysUtils, Grids, CSVdocument;

type

  { TKicadStock }

TKicadStock = class(TStringList)
  private
    Doc : TCSVDocument;
  public
    FileNames : TStringList;
    RefNames : TStringList;
    procedure ToStringGrid(SG : TStringGrid; index : integer);
    procedure LoadFromFile(const FileName : string); override;
    constructor Create;
    destructor Destroy; override;
end;

// ###########################################################################
implementation
// ###########################################################################

// ##################################################
{ TKicadStock }
// ##################################################

// ==================================================
// Copy to string grid
// ==================================================
procedure TKicadStock.ToStringGrid(SG: TStringGrid; index : integer);
var i,j, lim : integer;
    s, fn : string;
begin
  if (index >= 0) and (index < FileNames.Count) then begin
    fn := FileNames.Strings[index];
    Doc.LoadFromFile(fn);
    // ---------------------------------
    // Fill headers from row[0]
    // ---------------------------------
    lim := Doc.ColCount[0]-1;
    for j := 0 to lim do begin
      s := Doc.Cells[j,0];
      SG.Columns[j].Title.Caption := s;
      if j >= SG.ColCount-1 then
        break;
    end;
    // ---------------------------------
    // Fill grid
    // ---------------------------------
    SG.RowCount := Doc.RowCount;
    lim := Doc.RowCount-1;
    for i:=1 to lim do begin
      for j:=0 to SG.ColCount-1 do begin
        s := Doc.Cells[j,i];
        SG.Cells[j,i] := s;
      end;
    end;
  end;
end;

// ==================================================
// Load from file
// ==================================================
procedure TKicadStock.LoadFromFile(const FileName: string);
begin
  inherited LoadFromFile(FileName);
end;

// ==================================================
// Create
// ==================================================
constructor TKicadStock.Create;
begin
  FileNames := TStringList.Create;
  RefNames := TStringList.Create;
  Doc := TCSVDocument.Create;
  Doc.Delimiter:=char(9);
end;

// ==================================================
// Destroy
// ==================================================
destructor TKicadStock.Destroy;
begin
  Doc.Free;
  FileNames.Free;
  RefNames.Free;
  inherited Destroy;
end;

end.
