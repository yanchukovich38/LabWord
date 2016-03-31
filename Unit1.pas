unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, VBIDE_TLB,Word_TLB,Office_TLB;

type
  TForm1 = class(TForm)
    Edit6: TEdit;
    Edit7: TEdit;
    Label7: TLabel;
    Label9: TLabel;
    Label1: TLabel;
    Label5: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit5: TEdit;
    Edit4: TEdit;
    Edit3: TEdit;
    Label2: TLabel;
    Label8: TLabel;
    Label4: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    ComboBox1: TComboBox;
    Edit13: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Button1: TButton;
    Edit14: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  word: WordApplication;
  Doc: WordDocument;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  WordApp: WordApplication;
  Docs: Documents;
  Doc: WordDocument;
  Pars: Paragraphs;
  Par: Paragraph;
  D: OleVariant;
begin
  WordApp := CoWordApplication.Create;
  WordApp.Visible := True;
  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, True);
  Doc := (WordApp.Documents.Item(1) as WordDocument);
  Doc.Paragraphs.Item(1).Format.LeftIndent:=WordApp.CentimetersToPoints(4) ;
  Doc.Paragraphs.Item(1).Format.SpaceAfter:=10;
  Doc.Paragraphs.Item(1).Range.Text := 'Доверенность на управление автомобилем'
  +#13+ edit4.Text + #32 + 'дата'
  +#13+ 'Я, гр.' + #32 + edit1.Text + #32 + 'проживающий  в ' + #32
  + edit4.Text+ ',' + #32 + 'паспорт' + #32 + edit2.text + #32 + 'выданный' + edit3.Text
  + ',' + #32 + 'имея в собственности' + #32 + edit11.Text + #32 + 'государственный регистрационный знак'
  + #32 + edit7.Text + #32 + 'идентификационный номер (VIN)' + #32 + edit8.text + #32 + 'кузов'
  + #32 + combobox1.Text + #32 + 'цвет' + #32 + edit13.Text + #32 + 'паспорт транспортного средства' + #32 + edit9.Text
  + #32 + 'свидетельство о регистрации ТС' + #32 + edit10.Text + #32 + 'выдано ГИБДД' + #32 + 'ДОБАВИТЬ!!!'
  + #32 + 'стоит на учете в ГИБДД' + #32 + 'настоящей доверенностью уполномачиваю гр.' + #32 + edit5.Text + #32
  + 'проживающего в гор.' + #32 + edit6.Text + #32 + 'управлять указанным автомобилем, следить за его техническим состоянием(право ремонта), быть моим представителем в ГИБДД с правом совершения регистрационных действий ,'+ #32 +  'снятия с учета в ГИБДД, изменения регистрационных данных, замены номерных узлов и агрегатов, кузова, государственных номерных знаков, прохождения ГТО, получения дубликатов регистрационных документов на автомобиль с правом расписываться за меня'
  + #32 + 'и выполнять все действия, связанные с этим поручением'
  + #13 + 'Доверенность выдана без права передоверия до' + #32 + edit14.text + #32 + 'года'
  + #13 + 'Подпись __________________________________________________________________' ;
Doc.Paragraphs.Item(1).Range.Font.Color:=clblue;
Doc.Paragraphs.Item(2).Format.LeftIndent:=WordApp.CentimetersToPoints(0) ;
Doc.Paragraphs.Item(2).Format.SpaceAfter:=5;
Doc.Paragraphs.Item(3).Format.LeftIndent:=WordApp.CentimetersToPoints(0) ;
Doc.Paragraphs.Item(3).Format.SpaceAfter:=5;
Doc.Paragraphs.Item(4).Format.LeftIndent:=WordApp.CentimetersToPoints(0) ;
Doc.Paragraphs.Item(4).Format.SpaceAfter:=10;
Doc.Paragraphs.Item(5).Format.LeftIndent:=WordApp.CentimetersToPoints(0) ;
Doc.Paragraphs.Item(5).Format.SpaceAfter:=10;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
Combobox1.Items.Add('Седан');
Combobox1.Items.Add('Купэ');
Combobox1.Items.Add('Кабриолет');
Combobox1.Items.Add('Внедорожник');
Combobox1.Items.Add('Универсал');
Combobox1.Items.Add('Пикап');
end;

end.
