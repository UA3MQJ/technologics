program xls2pdf;

{$APPTYPE CONSOLE}

uses
  SysUtils, comobj, variants, activex;

var
  VExcel:variant;
  infile,outfile:string;
begin
  infile := paramstr(1);
  outfile := paramstr(2);

  //infile  := 'c:\temp\Template_22072014 93027.xls';
  //outfile := 'c:\temp\Template_22072014 93027.pdf';

  if ((infile='')or(outfile='')) then
  begin
    writeln('xls2pdf input_file.xls output_file.pdf');
  end else begin
    writeln('infile '+infile);
    writeln('outfile '+outfile);

    CoInitialize(nil);
    VExcel := CreateOleObject('Excel.Application');
    VExcel.Visible := false;
    VExcel.DisplayAlerts := false;
    VExcel.WorkBooks.Open(infile);
    VExcel.ActiveWorkBook.ExportAsFixedFormat(0,outfile, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    VExcel.ActiveWorkbook.Close;
    VExcel.Quit;
    VExcel:= Unassigned;
  end;
end.
