unit uSplashScreen;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,System.Math,
  Vcl.Imaging.jpeg, bsSkinCtrls, bsSkinExCtrls, IniFiles, AdvWiiProgressBar,
  Vcl.ComCtrls;

type
  TfrmBoasVindas = class(TForm)
    Button1: TButton;
    Image1: TImage;
    Shape1: TShape;
    bsSkinDivider1: TbsSkinDivider;
    Image2: TImage;
    Label1: TLabel;
    Label3: TLabel;
    CheckBox1: TCheckBox;
    Button2: TButton;
    Label4: TLabel;
    cb_Idioma: TComboBox;
    Button3: TButton;
    Timer1: TTimer;
    ProgressBar1: TProgressBar;
    Termos: TRichEdit;
    procedure Button1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure cb_IdiomaChange(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Shape1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image2MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CarrearIdioma(TrocarIdioma : Boolean);
    function LerConfiguracao : String;
  end;

var
  frmBoasVindas: TfrmBoasVindas;

implementation

{$R *.dfm}

uses uPrincipal;

procedure TfrmBoasVindas.Button1Click(Sender: TObject);
begin
  frmPrincipal.SalvarConfiguracao;
  frmPrincipal.Idiomaatual := cb_Idioma.Text;
  close;
end;

procedure TfrmBoasVindas.Button2Click(Sender: TObject);
begin
  if termos.Visible then
    termos.Visible := false
  else
    termos.Visible := true;
end;

procedure TfrmBoasVindas.Button3Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmBoasVindas.CarrearIdioma(TrocarIdioma : Boolean);
begin
  //True = Le do arquivo ini
  //False = Mostra os controles na tela inicialç para o usuario selecionar o idioma

  if TrocarIdioma = True then
    begin
      frmPrincipal.SelecionarIdioma(FrmPrincipal.IdiomaAtual);
      Label4.Visible := false;
      cb_Idioma.Visible := false;
      Button1.Visible := false;
      Button2.Visible := false;
      CheckBox1.Visible := false;

      if FrmPrincipal.IdiomaAtual = 'Português do Brasil (PT-BR)' then
        begin
          Label1.Caption := 'Criei o MailMerge para sanar uma necessidade própria, por padrãoo word não salva arquivos de mala direta em arquivos separados. Com o MailMerge é possivel criar uma mala direta sendo que os arquivos são gravados separadamente em doc e pdf.';
          frmPrincipal.Button4.width := 100;
          frmPrincipal.Button5.left := 114;
          frmPrincipal.cb_SalvarComo.Items.Clear;
          frmPrincipal.cb_SalvarComo.Items.Add('Documento do Word (*.DOC)');
          frmPrincipal.cb_SalvarComo.Items.Add('Documento PDF (*.PDF)');
          frmPrincipal.cb_SalvarComo.Items.Add('Salvar nos Formatos (*.DOC e *.PDF)');
          frmPrincipal.cb_SalvarComo.ItemIndex := 0;
        end
      else
      if FrmPrincipal.IdiomaAtual = 'United States English (ENG)' then
        begin
          Label1.Caption := 'I created MailMerge to meet my own need, by default Word does not save mail merge files in separate files. With MailMerge it is possible to create a mail merge and the files are saved separately in doc and pdf.';
          frmPrincipal.cb_SalvarComo.Items.Clear;
          frmPrincipal.cb_SalvarComo.Items.Add('Word document (*.DOC)');
          frmPrincipal.cb_SalvarComo.Items.Add('PDF document (*.PDF)');
          frmPrincipal.cb_SalvarComo.Items.Add('Save to Formats (*.DOC e *.PDF)');
          frmPrincipal.cb_SalvarComo.ItemIndex := 0;
          frmPrincipal.Button4.width := 100;
          frmPrincipal.Button5.left := 114;
        end
      else
      if FrmPrincipal.IdiomaAtual = 'Español Europeo (ESP)' then
        begin
          Label1.Caption := 'Creé MailMerge por mi propia necesidad; de forma predeterminada, Word no guarda archivos de combinación de correspondencia en archivos separados. MailMerge permite crear combinaciones de correspondencia guardando archivos por separado como doc y pdf.';
          frmPrincipal.cb_SalvarComo.Items.Clear;
          frmPrincipal.cb_SalvarComo.Items.Add('Documento de Word (*.DOC)');
          frmPrincipal.cb_SalvarComo.Items.Add('Documento PDF (*.PDF)');
          frmPrincipal.cb_SalvarComo.Items.Add('Guardar en formatos (*.DOC e *.PDF)');
          frmPrincipal.cb_SalvarComo.ItemIndex := 0;
          frmPrincipal.Button4.width := 200;
          frmPrincipal.Button5.left := 214;
        end;
    end
  else
  begin
    frmPrincipal.SelecionarIdioma(cb_Idioma.Text);
    Label4.Visible := True;
    cb_Idioma.Visible := True;
    Button1.Visible := True;
    Button2.Visible := True;
    CheckBox1.Visible := True;

    if cb_Idioma.Text = 'Português do Brasil (PT-BR)' then
      begin
        CheckBox1.Caption := 'Declaro que lí e concordo com os termos.';
        Label1.Caption := 'Criei o MailMerge para sanar uma necessidade própria, por padrãoo word não salva arquivos de mala direta em arquivos separados. Com o MailMerge é possivel criar uma mala direta sendo que os arquivos são gravados separadamente em doc e pdf.';
        Button2.Width := 75;
        Button2.Left := 481;
        frmPrincipal.Button4.width := 100;
        frmPrincipal.Button5.left := 114;

        frmPrincipal.cb_SalvarComo.Items.Clear;
        frmPrincipal.cb_SalvarComo.Items.Add('Documento do Word (*.DOC)');
        frmPrincipal.cb_SalvarComo.Items.Add('Documento PDF (*.PDF)');
        frmPrincipal.cb_SalvarComo.Items.Add('Salvar nos Formatos (*.DOC e *.PDF)');
        frmPrincipal.cb_SalvarComo.ItemIndex := 0;

        Termos.Lines.LoadFromFile(ExtractFilePath(ParamStr(0)) + '\docs\' + 'Termos de Uso do MailMerge (PT-BR).rtf');
      end
    else
    if cb_Idioma.Text = 'United States English (ENG)' then
      begin
        frmBoasVindas.CheckBox1.Caption := 'I declare that I have read and agree to the terms.';
        Label1.Caption := 'I created MailMerge to meet my own need, by default Word does not save mail merge files in separate files. With MailMerge it is possible to create a mail merge and the files are saved separately in doc and pdf.';
        Button2.Width := 85;
        Button2.Left := 471;
        frmPrincipal.cb_SalvarComo.Items.Clear;
        frmPrincipal.cb_SalvarComo.Items.Add('Word document (*.DOC)');
        frmPrincipal.cb_SalvarComo.Items.Add('PDF document (*.PDF)');
        frmPrincipal.cb_SalvarComo.Items.Add('Save to Formats (*.DOC e *.PDF)');
        frmPrincipal.cb_SalvarComo.ItemIndex := 0;
        frmPrincipal.Button4.width := 100;
        frmPrincipal.Button5.left := 114;

        Termos.Lines.LoadFromFile(ExtractFilePath(ParamStr(0)) + '\docs\' + 'MailMerge Terms of Use (ING).rtf');
      end
    else
    if cb_Idioma.Text = 'Español Europeo (ESP)' then
      begin
        frmBoasVindas.CheckBox1.Caption := 'Declaro que he leído y acepto los términos.';
        Label1.Caption := 'Creé MailMerge por mi propia necesidad; de forma predeterminada, Word no guarda archivos de combinación de correspondencia en archivos separados. MailMerge permite crear combinaciones de correspondencia guardando archivos por separado como doc y pdf.';
        Button2.Width := 90;
        Button2.Left := 465;
        frmPrincipal.cb_SalvarComo.Items.Clear;
        frmPrincipal.cb_SalvarComo.Items.Add('Documento de Word (*.DOC)');
        frmPrincipal.cb_SalvarComo.Items.Add('Documento PDF (*.PDF)');
        frmPrincipal.cb_SalvarComo.Items.Add('Guardar en formatos (*.DOC e *.PDF)');
        frmPrincipal.cb_SalvarComo.ItemIndex := 0;
        frmPrincipal.Button4.width := 200;
        frmPrincipal.Button5.left := 214;
        Termos.Lines.LoadFromFile(ExtractFilePath(ParamStr(0)) + '\docs\' + 'Términos de uso de MailMerge (ES).rtf');
      end;
  end;
end;

procedure TfrmBoasVindas.cb_IdiomaChange(Sender: TObject);
begin
  CarrearIdioma(false); //True para para carergar do arquivo INI e False para o usuario escolher
end;

procedure TfrmBoasVindas.CheckBox1Click(Sender: TObject);
begin
  Button1.Enabled := CheckBox1.Checked;
end;

procedure TfrmBoasVindas.FormCreate(Sender: TObject);
var i : Integer;
begin
if SysLocale.PriLangID = LANG_SPANISH then
  cb_Idioma.ItemIndex := 0
else
if SysLocale.PriLangID = LANG_ENGLISH then
  cb_Idioma.ItemIndex := 1
else
if SysLocale.PriLangID = LANG_PORTUGUESE then
  cb_Idioma.ItemIndex := 2;


  Randomize;
  try
    if LerConfiguracao <> 'Selecionar'  then
      begin
        FrmPrincipal.IdiomaAtual := LerConfiguracao;
        CarrearIdioma(true);
        ProgressBar1.Max :=  RandomRange(100,500);
        ProgressBar1.Visible := True;
        timer1.Enabled := True;
      end;
    except
    begin
      CarrearIdioma(false);
    end;
  end;
end;

procedure TfrmBoasVindas.FormMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
   sc_DragMove = $f012;
begin
  ReleaseCapture;
  Perform(wm_SysCommand, sc_DragMove, 0);
end;

procedure TfrmBoasVindas.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
   sc_DragMove = $f012;
begin
  ReleaseCapture;
  Perform(wm_SysCommand, sc_DragMove, 0);
end;

procedure TfrmBoasVindas.Image2MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
   sc_DragMove = $f012;
begin
  ReleaseCapture;
  Perform(wm_SysCommand, sc_DragMove, 0);
end;

function TfrmBoasVindas.LerConfiguracao : String;
var
  ArquivoINI: TIniFile;
begin
try
  ArquivoINI := TIniFile.Create(ExtractFilePath(ParamStr(0)) + '\' + 'Config.ini');
  Result := ArquivoINI.ReadString('Config', 'Idioma','');
  ArquivoINI.Free;
except
  Result := 'Selecionar';
end;
end;
procedure TfrmBoasVindas.Shape1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
   sc_DragMove = $f012;
begin
  ReleaseCapture;
  Perform(wm_SysCommand, sc_DragMove, 0);
end;

procedure TfrmBoasVindas.Timer1Timer(Sender: TObject);
begin
  if ProgressBar1.Max > ProgressBar1.Position then
    ProgressBar1.Position := ProgressBar1.Position + RandomRange(1,20)
  else
  Close;
end;

end.
