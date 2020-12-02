unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Vcl.StdCtrls, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.ExtCtrls, Vcl.WinXPickers, Vcl.CheckLst,
  FireDAC.VCLUI.Async, FireDAC.Comp.UI, FireDAC.Phys.IBWrapper, StrUtils, DateUtils, ComObj;

type

  TSclad = record
    ID: Integer;
    Name: String;
  end;

  TForm1 = class(TForm)
    DatePicker1: TDatePicker;
    FDGUIxAsyncExecuteDialog1: TFDGUIxAsyncExecuteDialog;
    GridPanel1: TGridPanel;
    CheckBox2: TCheckBox;
    GridPanel2: TGridPanel;
    CheckBox1: TCheckBox;
    Button2: TButton;
    Button3: TButton;
    GridPanel3: TGridPanel;
    RadioButton5: TRadioButton;
    RadioButton6: TRadioButton;
    RadioButton7: TRadioButton;
    RadioButton8: TRadioButton;
    GridPanel5: TGridPanel;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    DatePicker2: TDatePicker;
    CheckBox8: TCheckBox;
    CheckBox9: TCheckBox;
    CheckBox10: TCheckBox;
    GridPanel4: TGridPanel;
    ComboBox1: TComboBox;
    RadioButton2: TRadioButton;
    RadioButton1: TRadioButton;
    GridPanel6: TGridPanel;
    CheckBox11: TCheckBox;
    Button4: TButton;
    GridPanel7: TGridPanel;
    ComboBox2: TComboBox;
    CheckBox13: TCheckBox;
    CheckBox12: TCheckBox;
    GridPanel8: TGridPanel;
    Button1: TButton;
    Label1: TLabel;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    CheckBox14: TCheckBox;
    procedure ExportData(DateStart, DateEnd: TDate);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure RadioButton7Click(Sender: TObject);
    procedure RadioButton8Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure CheckBox11Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBox13Click(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Connection: TFDConnection;
  Query: TFDQuery;
  Sclads: array of TSclad;

implementation

{$R *.dfm}

uses Unit2, Unit3, Unit4;

Function GetScladIDByName(Str: String): Integer;
  Var
  I: Integer;
begin
  Result := -1;
  for I := Low(Sclads) to High(Sclads) do
    if AnsiUpperCase(Str)=AnsiUpperCase(Sclads[I].Name) then
  Result := Sclads[I].ID;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  ExportData(Datepicker1.Date,Datepicker2.Date);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  Form2.Show;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  Form3.Show;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  Form4.Show;
end;

procedure TForm1.CheckBox11Click(Sender: TObject);
begin
  if CheckBox11.Checked then
    Button4.Enabled := true
  else
    Button4.Enabled := false;
end;

procedure TForm1.CheckBox13Click(Sender: TObject);
begin
  if CheckBox13.Checked then
    Combobox2.Enabled := true
  else
    Combobox2.Enabled := false;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
  begin
    Button2.Enabled := true;
    Button3.Enabled := true;
  end
  else
  begin
    Button2.Enabled := false;
    Button3.Enabled := false;
  end;
end;

procedure TForm1.ExportData(DateStart, DateEnd: TDate);
  Var
  I, J: Integer;
  Result: TStringlist;
  ResultPath: String;
  OrderStatus: String;
  SortStr: String;
  PayStr: String;
  Statuses: String;
  Address: String;
  NeedToContinue: Boolean;
  XL: OleVariant;
begin

  Try

    Button1.Enabled := false;
    GridPanel1.Enabled := false;
    GridPanel2.Enabled := false;
    GridPanel3.Enabled := false;
    GridPanel5.Enabled := false;
    GridPanel7.Enabled := false;

    NeedToContinue := False;

    Statuses := '';

    ResultPath := '';

    Address := '';

    PayStr := '';

    SortStr := '';


    if CheckBox3.Checked=True then
    if Statuses='' then
    Statuses := '(dor.status_id='+#39+'1'+#39+')' else Statuses := Statuses + ' or (dor.status_id='+#39+'1'+#39+')';

    if CheckBox4.Checked=True then
    if Statuses='' then
    Statuses := '(dor.status_id='+#39+'3'+#39+')' else Statuses := Statuses + ' or (dor.status_id='+#39+'3'+#39+')';

    if CheckBox5.Checked=True then
    if Statuses='' then
    Statuses := '(dor.status_id='+#39+'4'+#39+')' else Statuses := Statuses + ' or (dor.status_id='+#39+'4'+#39+')';

    if CheckBox6.Checked=True then
    if Statuses='' then
    Statuses := '(dor.status_id='+#39+'5'+#39+')' else Statuses := Statuses + ' or (dor.status_id='+#39+'5'+#39+')';

    if CheckBox7.Checked=True then
    if Statuses='' then
    Statuses := '(dor.status_id='+#39+'7'+#39+')' else Statuses := Statuses + ' or (dor.status_id='+#39+'7'+#39+')';

    Statuses := '('+Statuses+')';

    if Statuses='()' then
    begin
        ShowMessage('Пожалуйста, выберите статусы готовности заказов!');
        Exit;
    end;

    if CheckBox1.Checked then
    begin
      if (Form2.Memo1.Lines.Count<>Form3.Memo1.Lines.Count) or
         (Form2.Memo1.Lines.Count=0) or
         (Form3.Memo1.Lines.Count=0) then
      begin
        ShowMessage('Кол-во телефонов и адресов не совпадает!'+#13#10+#13#10+
                    'Кол-во телефонов: '+IntToStr(Form2.Memo1.Lines.Count)+#13#10+#13#10+
                    'Кол-во адресов: '+IntToStr(Form3.Memo1.Lines.Count));
        exit;
      end;
    end;

    Try
      Connection := TFDConnection.Create(nil);
      Try
      Connection.DriverName := 'FB';
      with Connection.Params as TFDPhysFBConnectionDefParams do
        begin
          //Protocol := ipTCPIP;
          //Server := 'mail.apetta.ru';
          Server := '192.168.0.50';
          Database := 'E:\12345\DB\ARM.fdb';
          UserName := 'sysdba';
          Password := 'masterkey';
          IBAdvanced := 'config=WireCompression=false';
        end;
      Connection.Connected := True;
      Except
      on E:Exception do
        begin
          ShowMessage('Произошла ошибка при подключении к БД.'+#13#10+E.ClassName+' '+E.Message+#13#10+'Обратитесь к тех. поддержке!');
          Exit;
        end;
      End;

      Query := TFDQuery.Create(nil);
      Query.ResourceOptions.CmdExecMode := amCancelDialog;
      Try
        Query.Connection := Connection;
        if (RadioButton7.Checked=true) or (RadioButton3.Checked=true) or (RadioButton4.Checked=true) then
          Query.SQL.Add('select distinct(d.doc_num), '+
                        'coalesce(cast(d.date_cr as date),''Отсутствует'') as date_cr, '+
                        'coalesce(cast(dor.date_out as date),''Отсутствует'') as date_out, '+
                        'coalesce(cast(dor.date_out_fact as date),''Отсутствует'') as date_out_fact, '+
                        'dor.status_id, dor.kredit, dor.debet, c.name, c.teleph_cell, (select first 1 descript from doc_order_others doo where doo.doc_order_id=dor.id order by doo.dt DESC) as comment, s.ext_info ')
        else
          Query.SQL.Add('select distinct(d.doc_num), '+
                        'coalesce(cast(d.date_cr as date),''Отсутствует'') as date_cr, '+
                        'coalesce(cast(dor.date_out as date),''Отсутствует'') as date_out, '+
                        'coalesce(cast(dor.date_out_fact as date),''Отсутствует'') as date_out_fact, '+
                        'dor.status_id, dor.kredit, dor.debet, c.name, c.teleph_cell, (select first 1 descript from doc_order_others doo where doo.doc_order_id=dor.id order by doo.dt DESC) as comment ');
          Query.SQL.Add('from docs d ');
          Query.SQL.Add('left join deps dep on dep.dep_id=d.dep_id ');
          Query.SQL.Add('left join docs_order dor on dor.doc_id=d.doc_id ');
          Query.SQL.Add('left join contragents c on c.contr_id=d.contragent_id ');

        if (RadioButton7.Checked=true) or (RadioButton3.Checked=true) or (RadioButton4.Checked=true)  then
          Query.SQL.Add('left join doc_order_services s on s.doc_order_id=dor.id ');

        If RadioButton5.Checked=true then
          Query.SQL.Add('where  (cast(d.date_cr as date) >= '''+DateToStr(DateStart)+''') and (cast(d.date_cr as date) <= '''+DateToStr(DateEnd)+''') and ') else
          Query.SQL.Add('where  (cast(dor.date_out as date) >= '''+DateToStr(DateStart)+''') and (cast(dor.date_out as date) <= '''+DateToStr(DateEnd)+''') and ');

        if CheckBox13.Checked=True then
          Query.SQL.Add('(dor.sclad_kredit_id='+IntToStr(GetScladIDByName(ComboBox2.Text))+') and ');

        if CheckBox8.Checked=True then
          if PayStr='' then
            PayStr := '(dor.debet<dor.kredit)' else PayStr := PayStr + ' or (dor.debet<dor.kredit)';

        if CheckBox9.Checked=True then
          if PayStr='' then
            PayStr := '(dor.debet=dor.kredit)' else PayStr := PayStr + ' or (dor.debet=dor.kredit)';

        if CheckBox10.Checked=True then
          if PayStr='' then
            PayStr := '(dor.debet>dor.kredit)' else PayStr := PayStr + ' or (dor.debet>dor.kredit)';

        PayStr := '('+PayStr+')';

        if PayStr='()' then
        begin
            ShowMessage('Пожалуйста, выберите статусы оплаты заказов!');
            Exit;
        end;

        if RadioButton7.Checked=true then
        begin
          Query.SQL.Add('(s.tovar_id=1002241) and (s.status_id<>7) and ');
          if CheckBox2.State=cbChecked then
          Query.SQL.Add('(s.ext_info <> '''') and ')
          else
          if CheckBox2.State=cbUnchecked then
          Query.SQL.Add('(s.ext_info = '''') and ');
        end;

        if RadioButton3.Checked=true then
        begin
          Query.SQL.Add('((s.tovar_id=1001053) or (s.tovar_id=1001054) or (s.tovar_id=1001055)) and (s.status_id<>7) and ');
          if CheckBox12.State=cbChecked then
          Query.SQL.Add('(s.ext_info <> '''') and ')
          else
          if CheckBox12.State=cbUnchecked then
          Query.SQL.Add('(s.ext_info = '''') and ');
        end;

        if RadioButton4.Checked=true then
        begin
          Query.SQL.Add('(s.tovar_id=1002270) and ');
          if CheckBox14.State=cbChecked then
          Query.SQL.Add('(s.ext_info <> '''') and ')
          else
          if CheckBox14.State=cbUnchecked then
          Query.SQL.Add('(s.ext_info = '''') and ');
        end;

        Query.SQL.Add(PayStr+' and ');

        Query.SQL.Add(Statuses);

        case ComboBox1.ItemIndex of
          0: SortStr := 'Order by d.doc_num';
          1: SortStr := 'Order by dor.status_id';
          2: SortStr := 'Order by dor.kredit';
          3: SortStr := 'Order by dor.debet';
          4: SortStr := 'Order by date_cr';
          5: SortStr := 'Order by date_out';
          6: SortStr := 'Order by date_out_fact';
        end;

        if RadioButton1.Checked=true then
        SortStr := SortStr+' ASC' else SortStr := SortStr+' DESC';

        Query.SQL.Add(SortStr);

        Query.Open;
        Query.FetchAll;
      Except
      on E:Exception do
        begin
          ShowMessage('Произошла ошибка при запросе к БД.'+#13#10+E.ClassName+' '+E.Message+#13#10+'Обратитесь к тех. поддержке!');
          exit;
        end;
      End;

      XL := CreateOleObject('Excel.Application');
      XL.DisplayAlerts := true;
      XL.WorkBooks.Add;
      XL.WorkBooks[1].WorkSheets[1].Name := DateToStr(DateStart) + ' - ' + DateToStr(DateEnd);
      XL.WorkBooks[1].WorkSheets[1].Cells[1,1] := '# заказа';
      XL.Columns[1].ColumnWidth := 10;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,2] := 'Статус заказа';
      XL.Columns[2].ColumnWidth := 10;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,3] := 'Дата приема';
      XL.Columns[3].ColumnWidth := 10;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,4] := 'Пред. дата выдачи';
      XL.Columns[4].ColumnWidth := 20;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,5] := 'Факт. дата выдачи';
      XL.Columns[5].ColumnWidth := 20;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,6] := 'Комментарий';
      XL.Columns[6].ColumnWidth := 100;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,7] := 'Стоимость заказа';
      XL.Columns[7].ColumnWidth := 10;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,8] := 'Оплачено';
      XL.Columns[8].ColumnWidth := 10;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,9] := 'ФИО клиента';
      XL.Columns[9].ColumnWidth := 20;
      XL.WorkBooks[1].WorkSheets[1].Cells[1,10] := 'Телефон клиента';
      XL.Columns[10].ColumnWidth := 20;
      if (RadioButton7.Checked=true) or (RadioButton4.Checked=true) or (RadioButton3.Checked=true) or (CheckBox1.Checked=true) then
      begin
      XL.WorkBooks[1].WorkSheets[1].Cells[1,11] := 'Адрес клиента';
      XL.Columns[11].ColumnWidth := 50;
      end;
      XL.Visible := true;

      Result := TStringList.Create;
      Result.Text := 'Sep=;';
      if (RadioButton4.Checked=true) or (RadioButton7.Checked=true) or (RadioButton3.Checked=true) or (CheckBox1.Checked=true) then
        Result.Add('# заказа;Статус заказа;Дата приема;Пред. дата выдачи;Факт. дата выдачи;Комментарий;Стоимость заказа;Оплачено;ФИО клиента;Телефон клиента;Адрес клиента')
      else
        Result.Add('# заказа;Статус заказа;Дата приема;Пред. дата выдачи;Факт. дата выдачи;Комментарий;Стоимость заказа;Оплачено;ФИО клиента;Телефон клиента');
      if Query.RecordCount>0 then
      begin
        Query.First;
        For I:=0 to Query.RecordCount-1 do
        begin

          if CheckBox11.Checked then
          begin
            NeedToContinue := false;
            for J := 0 to Form4.Memo1.Lines.Count-1 do
            if Query.FieldByName('doc_num').AsString=Form4.Memo1.Lines.Strings[J] then
            begin
              NeedToContinue := true;
              Break;
            end;
            if NeedToContinue=True then
            begin
              Query.Next;
              Continue;
            end;
          end;

          OrderStatus := '';
          if Query.FieldByName('status_id').AsString='1' then
            OrderStatus := 'Новый';
          if Query.FieldByName('status_id').AsString='3' then
            OrderStatus := 'В исполнении';
          if Query.FieldByName('status_id').AsString='4' then
            OrderStatus := 'Исполненный';
          if Query.FieldByName('status_id').AsString='5' then
            OrderStatus := 'Выданный';
          if Query.FieldByName('status_id').AsString='7' then
            OrderStatus := 'Отмененный';
          if CheckBox1.Checked=false then
          begin
            if (RadioButton7.Checked=true) or (RadioButton3.Checked=true) or (RadioButton4.Checked=true) then
            begin
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1].NumberFormat := '@';
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1] := Query.FieldByName('doc_num').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,2] := OrderStatus;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,3] := Query.FieldByName('date_cr').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,4] := Query.FieldByName('date_out').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,5] := Query.FieldByName('date_out_fact').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,6] := StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase]);
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,7] := Query.FieldByName('kredit').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,8] := Query.FieldByName('debet').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,9] := Query.FieldByName('name').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,10] := Query.FieldByName('teleph_cell').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Query.FieldByName('ext_info').AsString;
            //Result.Add(Query.FieldByName('doc_num').AsString+';'+OrderStatus+';'+Query.FieldByName('date_cr').AsString+';'+Query.FieldByName('date_out').AsString+';'+Query.FieldByName('date_out_fact').AsString+';'+StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase])+';'+Query.FieldByName('kredit').AsString+';'+Query.FieldByName('debet').AsString+';'+Query.FieldByName('name').AsString+';'+Query.FieldByName('teleph_cell').AsString+';'+Query.FieldByName('ext_info').AsString)
            end
            else
            begin
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1].NumberFormat := '@';
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1] := Query.FieldByName('doc_num').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,2] := OrderStatus;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,3] := Query.FieldByName('date_cr').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,4] := Query.FieldByName('date_out').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,5] := Query.FieldByName('date_out_fact').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,6] := StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase]);
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,7] := Query.FieldByName('kredit').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,8] := Query.FieldByName('debet').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,9] := Query.FieldByName('name').AsString;
            XL.WorkBooks[1].WorkSheets[1].Cells[I+2,10] := Query.FieldByName('teleph_cell').AsString;
            //Result.Add(Query.FieldByName('doc_num').AsString+';'+OrderStatus+';'+Query.FieldByName('date_cr').AsString+';'+Query.FieldByName('date_out').AsString+';'+Query.FieldByName('date_out_fact').AsString+';'+StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase])+';'+Query.FieldByName('kredit').AsString+';'+Query.FieldByName('debet').AsString+';'+Query.FieldByName('name').AsString+';'+Query.FieldByName('teleph_cell').AsString);
            end;
          End else
          begin
            Address := '';
            for J := 0 to Form2.Memo1.Lines.Count-1 do
            if Pos(AnsiReplaceStr(Query.FieldByName('teleph_cell').AsString,'+7',''),Form2.Memo1.Lines.Strings[J])<>0 then
            Address := AnsiReplaceStr(Form3.Memo1.Lines.Strings[J],#9,' ');
            if (RadioButton7.Checked=true) or (RadioButton3.Checked=true) or (RadioButton4.Checked=true) then
            begin
              if Address<>'' then
              begin
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1].NumberFormat := '@';
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1] := Query.FieldByName('doc_num').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,2] := OrderStatus;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,3] := Query.FieldByName('date_cr').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,4] := Query.FieldByName('date_out').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,5] := Query.FieldByName('date_out_fact').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,6] := StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase]);
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,7] := Query.FieldByName('kredit').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,8] := Query.FieldByName('debet').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,9] := Query.FieldByName('name').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,10] := Query.FieldByName('teleph_cell').AsString;
              if (Query.FieldByName('ext_info').AsString<>'') and (Address<>'') then
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Query.FieldByName('ext_info').AsString+' или '+Address
              else
                if (Query.FieldByName('ext_info').AsString<>'') then
                  XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Query.FieldByName('ext_info').AsString
                else
                  if (Address<>'') then
                    XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Address;
              end
              else
              begin
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1].NumberFormat := '@';
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1] := Query.FieldByName('doc_num').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,2] := OrderStatus;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,3] := Query.FieldByName('date_cr').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,4] := Query.FieldByName('date_out').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,5] := Query.FieldByName('date_out_fact').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,6] := StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase]);
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,7] := Query.FieldByName('kredit').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,8] := Query.FieldByName('debet').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,9] := Query.FieldByName('name').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,10] := Query.FieldByName('teleph_cell').AsString;
                XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Query.FieldByName('ext_info').AsString;
              end;
            end
            else
            begin
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1].NumberFormat := '@';
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,1] := Query.FieldByName('doc_num').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,2] := OrderStatus;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,3] := Query.FieldByName('date_cr').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,4] := Query.FieldByName('date_out').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,5] := Query.FieldByName('date_out_fact').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,6] := StringReplace(Query.FieldByName('comment').AsString, #13#10, ' ',[rfReplaceAll, rfIgnoreCase]);
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,7] := Query.FieldByName('kredit').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,8] := Query.FieldByName('debet').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,9] := Query.FieldByName('name').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,10] := Query.FieldByName('teleph_cell').AsString;
              XL.WorkBooks[1].WorkSheets[1].Cells[I+2,11] := Address;
            end;
          end;
          Query.Next;
        end;

        XL.Columns.AutoFit;
      end else ShowMessage('Найти заказы по заданным условиям не удалось!');
    Finally
      If Assigned(Query) then FreeAndNil(Query);
      If Assigned(Connection) then FreeAndNil(Connection);
      If Assigned(Result) then FreeAndNil(Result);
      XL := null;
    End;
  Finally
    Button1.Enabled := true;
    GridPanel1.Enabled := true;
    GridPanel2.Enabled := true;
    GridPanel3.Enabled := true;
    GridPanel5.Enabled := true;
    GridPanel7.Enabled := true;
  End;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  DatePicker1.Date := IncWeek(Now,-1);
  DatePicker2.Date := Now;
end;

procedure TForm1.FormShow(Sender: TObject);
Var
I: Integer;
begin
    Try
      Connection := TFDConnection.Create(nil);
      Try
      Connection.DriverName := 'FB';
      with Connection.Params as TFDPhysFBConnectionDefParams do
        begin
          Server := '192.168.0.50';
          Database := 'E:\12345\DB\ARM.fdb';
          UserName := 'sysdba';
          Password := 'masterkey';
          IBAdvanced := 'config=WireCompression=false';
        end;
      Connection.Connected := True;
      Except
      on E:Exception do
        begin
          ShowMessage('Произошла ошибка при подключении к БД.'+#13#10+#13#10+E.ClassName+' '+E.Message+#13#10+#13#10+'Обратитесь к тех. поддержке!');
          Exit;
        end;
      End;

      Query := TFDQuery.Create(nil);
      Query.ResourceOptions.CmdExecMode := amCancelDialog;
      Try
        Query.Connection := Connection;
        try
          Query.Open('select * from sclads order by name');
          Query.FetchAll;
        except
          ShowMessage('Не удалось загрузить названия складов!');
          Application.Terminate;
        end;

        if Query.RecordCount<=0 then
        begin
          ShowMessage('Не удалось загрузить названия складов!');
          Application.Terminate;
        end else
        begin
          Combobox2.Items.Clear;
          SetLength(Sclads,0);
          SetLength(Sclads,Query.RecordCount);
          for I := 0 to Query.RecordCount-1 do
            begin
              Sclads[I].ID := Query.FieldByName('ID').AsInteger;
              Sclads[I].Name := Query.FieldByName('NAME').AsString;
              Combobox2.Items.Add(Sclads[I].Name);
              Query.Next;
            end;
          Combobox2.ItemIndex:=0;
        end;
      Finally
        FreeAndNil(Query);
      End;
    Finally
      FreeAndNil(Connection);
    End;
end;

procedure TForm1.RadioButton3Click(Sender: TObject);
begin
  CheckBox2.Checked := false;
  CheckBox2.Enabled := false;
  CheckBox14.Checked := false;
  CheckBox14.Enabled := false;
  CheckBox12.Enabled := true;
end;

procedure TForm1.RadioButton4Click(Sender: TObject);
begin
  CheckBox2.Checked := false;
  CheckBox2.Enabled := false;
  CheckBox12.Checked := false;
  CheckBox12.Enabled := false;
  CheckBox14.Enabled := true;
end;

procedure TForm1.RadioButton7Click(Sender: TObject);
begin
  CheckBox12.Checked := false;
  CheckBox12.Enabled := false;
  CheckBox14.Checked := false;
  CheckBox14.Enabled := false;
  CheckBox2.Enabled := true;
end;

procedure TForm1.RadioButton8Click(Sender: TObject);
begin
  CheckBox2.Checked := false;
  CheckBox2.Enabled := false;
  CheckBox12.Checked := false;
  CheckBox12.Enabled := false;
  CheckBox14.Checked := false;
  CheckBox14.Enabled := false;
end;

end.
