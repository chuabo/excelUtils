unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ActnList, Spin, Buttons, ButtonPanel,
  fpspreadsheetgrid, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    btnFind: TButton;
    BtnNew: TButton;
    BtnOpen: TButton;
    BtnSave: TButton;
    btnStop: TButton;
    CbReadFormulas: TCheckBox;
    CbAutoCalc: TCheckBox;
    cbIgnoreFirst: TCheckBox;
    edtFindRet: TEdit;
    edtResult: TEdit;
    edtFindVal: TEdit;
    edtFindRange: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    lbState: TLabel;
    memLog: TMemo;
    Panel3: TPanel;
    Panel4: TPanel;
    SheetsCombo: TComboBox;
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    SaveDialog: TSaveDialog;
    Splitter1: TSplitter;
    WorksheetGrid: TsWorksheetGrid;
    procedure btnFindClick(Sender: TObject);
    procedure BtnNewClick(Sender: TObject);
    procedure BtnOpenClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CbAutoCalcChange(Sender: TObject);
    procedure CbReadFormulasChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure memLogDblClick(Sender: TObject);
    procedure SheetsComboSelect(Sender: TObject);
    procedure ToggleBox1Change(Sender: TObject);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
    procedure showMsg(msg:string);
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  fpcanvas, lazutf8,variants,StrUtils,
  fpstypes, fpsutils, fpsReaderWriter, fpspreadsheet;

var
  isStop:boolean;


{ TForm1 }

procedure TForm1.BtnNewClick(Sender: TObject);
var
  dlg: TForm;
  edCols, edRows: TSpinEdit;
begin
  dlg := TForm.Create(nil);
  try
    dlg.Width := 220;
    dlg.Height := 128;
    dlg.Position := poMainFormCenter;
    dlg.Caption := 'New workbook';
    edCols := TSpinEdit.Create(dlg);
    with edCols do begin
      Parent := dlg;
      Left := dlg.ClientWidth - Width - 24;
      Top := 16;
      Value := WorksheetGrid.ColCount - ord(WorksheetGrid.ShowHeaders);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edCols.Top + 3;
      Caption := 'Columns:';
      FocusControl := edCols;
    end;
    edRows := TSpinEdit.Create(dlg);
    with edRows do begin
      Parent := dlg;
      Left := edCols.Left;
      Top := edCols.Top + edCols.Height + 8;
      Value := WorksheetGrid.RowCount - ord(WorksheetGrid.ShowHeaders);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edRows.Top + 3;
      Caption := 'Rows:';
      FocusControl := edRows;
    end;
    with TButtonPanel.Create(dlg) do begin
      Parent := dlg;
      Align := alBottom;
      ShowButtons := [pbCancel, pbOK];
    end;
    if dlg.ShowModal = mrOK then begin
      WorksheetGrid.NewWorkbook(edCols.Value, edRows.Value);
      SheetsCombo.Items.Clear;
      SheetsCombo.Items.Add('Sheet 1');
      SheetsCombo.ItemIndex := 0;
    end;
  finally
    dlg.Free;
  end;
end;

var
  isDebug:boolean=true;
procedure TForm1.btnFindClick(Sender: TObject);
var
  i,j,k:integer;
  s,lines:String;
  list:TStringList;
  findStart,findEnd,findRet:integer;//查找范围和返回列，序号从1开始
  findValue:integer;//要查找的列
  findResult:integer;
  //fr:String;
  ch:char;
  arr_str:array of string;
  isFind:Boolean;
begin
  //showmessage(IntToStr(ord('a')-97));
  s:=edtFindRange.Text;
  if pos(':',s)<1 then begin
     MessageDlg('提示,请输入正确的查找范围！',mtError,[mbOk],0);
     exit;
  end;

  isStop:=false;

  ch:=lowercase(pchar(copy(s,1,pos(':',s)-1))^);
  findStart:=ord(ch)-97;
  showMsg('findStart:'+IntToStr(findStart));
  findEnd:=ord(lowercase(pchar(copy(s,pos(':',s)+1,length(s)-pos(':',s)))^))-97;
  showMsg('findEnd:'+IntToStr(findEnd));

  s:=edtFindVal.text;
  findValue:=ord(lowercase(pchar(s)^))-97;
  showMsg('findValue:'+IntToStr(findValue));

  s:=edtResult.text;
  findResult:=ord(lowercase(pchar(s)^))-97;
  showMsg('findResult:'+IntToStr(findResult));

  s:=edtFindRet.text;
  findRet:=ord(lowercase(pchar(s)^))-97;
  showMsg('findRet:'+IntToStr(findRet));

  list:=TStringList.Create();
  try
    for i:=0 to worksheetGrid.RowCount-1 do begin
      Application.ProcessMessages;
      if isStop=true then
         exit;

      lines:='';
      for j:=1 to worksheetgrid.colCount-1 do begin
        s:=vartoStr(worksheetgrid.Cells[j,i]);
        lines:=lines+','+ s;

      end;
      delete(lines,1,1);
      //memLog.Lines.Add(lines);
      list.Add(lines);
    end;
    showMsg(list.Text);

    for i:=0 to list.Count-1 do begin
      Application.processMessages;
      if isStop=true then
         exit;

      if i=0 then begin
         if cbIgnoreFirst.Checked=true then
            continue;
      end;
      lbState.caption:=IntToStr(i)+'-->'+IntTostr(list.count-1);
      showMsg(list.Strings[i]);
      lines:=list.Strings[i];
      //arr_str:=lines.Split(',');
      //s:=arr_str[findValue];
      s:=lines.split(',')[findValue];
      if s<>'' then begin //查找的值不为空
         isFind:=false;
         for j:=1 to list.count-1 do begin
           lines:=list.strings[j];
           arr_str:=lines.split(',');
           //for k:=0 to arr_str.count-1 do begin
           for k:=findStart to findEnd do begin
             if s=arr_str[k] then begin
               worksheetGrid.Cells[findResult+1,i]:=arr_str[findRet];
               isFind:=true;
               break;
             end;
           end;
           if isFind=true then //只查找第一个匹配的行
              break;
         end;

         if isFind=false then
            worksheetgrid.cells[findResult+1,i]:='';
      end;
    end;

    MessageDlg('查找完毕！', mtInformation, [mbOK], 0);
  finally
    list.Free;
  end;

end;

procedure TForm1.BtnOpenClick(Sender: TObject);
begin
  if OpenDialog.FileName <> '' then begin
    OpenDialog.InitialDir := ExtractFileDir(OpenDialog.FileName);
    OpenDialog.FileName := ChangeFileExt(ExtractFileName(OpenDialog.FileName), '');
  end;
  if OpenDialog.Execute then begin
    LoadFile(OpenDialog.FileName);
  end;
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.BtnSaveClick(Sender: TObject);
var
  err, fn: String;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  if WorksheetGrid.Workbook.Filename <>'' then begin
    fn := AnsiToUTF8(WorksheetGrid.Workbook.Filename);
    SaveDialog.InitialDir := ExtractFileDir(fn);
    SaveDialog.FileName := ChangeFileExt(ExtractFileName(fn), '');
  end;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      WorksheetGrid.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));
      //WorksheetGrid.WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));     // works as well
    finally
      Screen.Cursor := crDefault;
      // Show a message in case of error(s)
      err := WorksheetGrid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TForm1.btnStopClick(Sender: TObject);
begin
  isStop:=true;
end;

procedure TForm1.Button1Click(Sender: TObject);

begin
end;

procedure TForm1.CbAutoCalcChange(Sender: TObject);
begin
  WorksheetGrid.AutoCalc := CbAutoCalc.Checked;
end;

procedure TForm1.CbReadFormulasChange(Sender: TObject);
begin
  WorksheetGrid.ReadFormulas := CbReadFormulas.Checked;
end;

procedure TForm1.FormCreate(Sender: TObject);
const
  THICK_BORDER: TsCellBorderStyle = (LineStyle: lsThick; Color: clNavy);
  MEDIUM_BORDER: TsCellBorderSTyle = (LineStyle: lsMedium; Color: clRed);
  DOTTED_BORDER: TsCellBorderSTyle = (LineStyle: lsDotted; Color: clRed);
begin
  // Add some cells and formats
  {WorksheetGrid.ColWidths[1] := 180;
  WorksheetGrid.ColWidths[2] := 100;

  WorksheetGrid.Cells[1,1] := 'This is a demo';
  WorksheetGrid.MergeCells(1,1, 2,1);
  WorksheetGrid.HorAlignment[1,1] := haCenter;
  WorksheetGrid.CellBorders[1,1, 2,1] := [cbSouth];
  WorksheetGrid.CellBorderStyles[1,1, 2,1, cbSouth] := THICK_BORDER;
  WorksheetGrid.BackgroundColors[1,1, 2,1] := RGBToColor(220, 220, 220);
  WorksheetGrid.CellFontColor[1,1] := clNavy;
  WorksheetGrid.CellFontStyle[1,1] := [fssBold];

  WorksheetGrid.Cells[1,2] := 'Number:';
  WorksheetGrid.HorAlignment[1,2] := haRight;
  WorksheetGrid.CellFontStyle[1,2] := [fssItalic];
  WorksheetGrid.CellFontColor[1,2] := clNavy;
  WorksheetGrid.Cells[2,2] := 1.234;

  WorksheetGrid.Cells[1,3] := 'Date:';
  WorksheetGrid.HorAlignment[1,3] := haRight;
  WorksheetGrid.CellFontStyle[1,3] := [fssItalic];
  WorksheetGrid.CellFontColor[1,3] := clNavy;
  WorksheetGrid.NumberFormat[2,3] := 'mmm dd, yyyy';
  WorksheetGrid.Cells[2,3] := date;

  WorksheetGrid.Cells[1,4] := 'Time:';
  WorksheetGrid.HorAlignment[1,4] := haRight;
  WorksheetGrid.CellFontStyle[1,4] := [fssItalic];
  WorksheetGrid.CellFontColor[1,4] := clNavy;
  WorksheetGrid.NumberFormat[2,4] := 'hh:nn';
  WorksheetGrid.Cells[2,4] := now();

  WorksheetGrid.Cells[1,5] := 'Rich text:';
  WorksheetGrid.HorAlignment[1,5] := haRight;
  WorksheetGrid.CellFontStyle[1,5] := [fssItalic];
  WorksheetGrid.CellFontColor[1,5] := clNavy;
  WorksheetGrid.Cells[2,5] := '100 cm<sup>2</sup>';

  WorksheetGrid.Cells[1,6] := 'Formula:';
  WorksheetGrid.HorAlignment[1,6] := haRight;
  WorksheetGrid.CellFontStyle[1,6] := [fssItalic];
  WorksheetGrid.CellFontColor[1,6] := clNavy;
  WorksheetGrid.Cells[2,6] := '=B2^2*PI()';
  WorksheetGrid.CellComment[2,6] := 'Area of the circle with radius given in cell B2';
  WorksheetGrid.NumberFormat[2,6] := '0.000';

  CbAutoCalc.Checked := WorksheetGrid.AutoCalc;
  CbReadFormulas.Checked := WorksheetGrid.ReadFormulas;  }
end;

procedure TForm1.memLogDblClick(Sender: TObject);
begin
  memLog.clear;
end;

procedure TForm1.SheetsComboSelect(Sender: TObject);
begin
  WorksheetGrid.SelectSheetByIndex(SheetsCombo.ItemIndex);
end;

procedure TForm1.ToggleBox1Change(Sender: TObject);
begin

end;

// Loads first worksheet from file into grid
procedure TForm1.LoadFile(const AFileName: String);
var
  err: String;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

      // Update user interface
      Caption := Format('ExcelUtils - %s (%s)', [
        AFilename,
        GetSpreadTechnicalName(WorksheetGrid.Workbook.FileFormatID)
      ]);

      // Collect the sheet names in the combobox for switching sheets.
      WorksheetGrid.GetSheets(SheetsCombo.Items);
      SheetsCombo.ItemIndex := 0;
    except
      on E:Exception do begin
        // Empty worksheet instead of the loaded one
        WorksheetGrid.NewWorkbook(26, 100);
        Caption := 'fpsGrid - no name';
        SheetsCombo.Items.Clear;
        // Grab the error message
        WorksheetGrid.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := WorksheetGrid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.showMsg(msg: string);
begin
  if isDebug then
     memLog.lines.add(msg);
end;


initialization
  {$I mainform.lrs}

end.

