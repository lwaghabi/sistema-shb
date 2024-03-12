Attribute VB_Name = "Geral"
Option Explicit

Public rs As New ADODB.Recordset
Public usu As New ADODB.Recordset
Public db As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public glb As New ADODB.Recordset
Public neg As New ADODB.Recordset
Public hneg As New ADODB.Recordset
Public dneg As New ADODB.Recordset
Public hdneg As New ADODB.Recordset
Public pes As New ADODB.Recordset
Public ctp As New ADODB.Recordset
Public hctp As New ADODB.Recordset
Public ctr As New ADODB.Recordset
Public hctr As New ADODB.Recordset
Public nfe As New ADODB.Recordset
Public hnfe As New ADODB.Recordset
Public dnfe As New ADODB.Recordset     'Detalhe Nota Fiscal
Public hdnfe As New ADODB.Recordset
Public nfd As New ADODB.Recordset      'nfd = Nota Fiscal Desdobramento
Public hnfd As New ADODB.Recordset
Public CartRep As New ADODB.Recordset
Public CartPromot As New ADODB.Recordset
Public Contato As New ADODB.Recordset
Public ICM As New ADODB.Recordset
Public UfRegiao As New ADODB.Recordset
Public CondProc As New ADODB.Recordset
Public FreteCobranca As New ADODB.Recordset
Public Bco As New ADODB.Recordset
Public NatuOper As New ADODB.Recordset
Public Prod As New ADODB.Recordset
Public ProdPco As New ADODB.Recordset
Public Emp As New ADODB.Recordset
Public TabPreco As New ADODB.Recordset
Public ProdPreco As New ADODB.Recordset
Public UnidEmb As New ADODB.Recordset
Public ProdTerc As New ADODB.Recordset
Public ProdEntrada As New ADODB.Recordset
Public ProdFornec As New ADODB.Recordset
Public tlanc As New ADODB.Recordset
Public fpag As New ADODB.Recordset
Public uoper As New ADODB.Recordset
Public gge As New ADODB.Recordset
Public gdet As New ADODB.Recordset
Public unid As New ADODB.Recordset
Public Ativ As New ADODB.Recordset
Public Rmb As New ADODB.Recordset
Public RmbDet As New ADODB.Recordset
Public ext As New ADODB.Recordset
Public evt As New ADODB.Recordset
Public lgt As New ADODB.Recordset
Public asoe As New ADODB.Recordset
Public asoa As New ADODB.Recordset
Public cto As New ADODB.Recordset
Public agcto As New ADODB.Recordset
Public eqpt As New ADODB.Recordset
Public eqh As New ADODB.Recordset
Public teq As New ADODB.Recordset
Public leq As New ADODB.Recordset
Public contab As New ADODB.Recordset
Public ccc As New ADODB.Recordset

Public acDb As Integer  'acDb é o acumulado de rotinas usando o banco de dados

Public acGlb As Integer
Public acUsu As Integer
Public acNeg As Integer
Public achNeg As Integer
Public acdNeg As Integer
Public achdNeg As Integer

Public acPes As Integer
Public acContab As Integer

Public Path As String

Public MesPedido As Integer

Public DataComAtraso As Date
Public DiadaSemana As Integer
Public DataHojeInvertida As String


Public UltPessoa As String

Global glbMaquina As String
Global glbEnderecoIP As String
Global glbTipoAcesso As String
Global glbStatusUsuario As Byte
Global glbStusSistema As Byte
Global GlbStatus As String
Global glbFuncao As String
Global glbNumPedido As String
Global glbCompPedido As String
Global glbEmpresa As Integer
Global glbEmissorNF As String
Global glbTransporte As String
Global glbNotaFiscal As String
Global glbBanco As Integer
Global glbPlaca As String
Global glbOrdemDeCarga As String
Global glbMotorista As String
Global glbCFOP As String
Global Incluir As Byte
Global Aba_Pessoa As Byte
Global Aba_Endereco As Byte
Global Aba_Detalhes As Byte
Global Erro_Critica As Integer
Global flagRegAnt As Byte
Global indPedido As Integer
Global Tabela_Pedido(500) As String

Global Foto_Pedida As String
Global Caminho As String
Global Extensao As String

Global Data_Hoje As Date

Global Ano_Hoje As Integer
Global Mes_Hoje As Integer
Global Dia_Hoje As Integer

'Definição de parâmetros de trabalho para atualização do Estoque

Global Funcao As Byte '1=Atualiza producao
                      '2=Atualiza Devolucao
                      '3=Atualiza Vendas
Global ano As Integer
Global mes As Integer
Global Dia As Integer
Global Produto As String
Global Entra As Currency
Global Sai As Currency
Global TracoIn As Integer
Global TracoOut As Integer
Global Mes_Pedido As Integer

Global OrdemDeCarga As String
Global EmissorOrdemDeCarga As String

Global fornecedor As String

Global GeradorCntrl As Byte

Global producao As Currency
Global QtdInicial As Currency
Global Devolucao As Currency
Global Venda As Currency
Global AssistTec As Currency
Global Promocao As Currency

Global Inclui_Estoque As Byte
Global Altera_Estoque As Byte

Global glbUsuario As String
Global StatusSistema As Byte

Global Transportadora As String
Global ultimoRegistro As String

Global NDias As Integer
Global DataInformada As Date
Global DataRetorno As Date
Global Verifica As String

'Fim da definição de parâmetros para atualização de Estoque

Global Status_Atualiza_Estoque

Public CON As New ADODB.Connection
Public rsl As New ADODB.Connection

'Public Function Compilando() As Boolean
'Compilando = App.Path Like "*Meus Documentos\SISTEMA*"
'End Function

Public Sub AbrirRelatorio(sql As String, Rel As Object)
Dim tentou As Boolean
Dim server As String
Dim senha As String
Dim usuario As String
Dim banco As String
Dim erroDriver As Boolean

On Error GoTo Erro
Inicio:
server = "mysql.sistemaos.com.br;"
banco = "sistemaos03;"
usuario = banco
senha = "zinholui47"

'         If Compilando Then
'            server = "localhost;"
'            banco = "Local;"
'            usuario = "root;"
'            senha = ""
'         End If

         Set CON = CreateObject("ADODB.Connection")
         'Set rel = CreateObject("ADODB.Recordset")
         Dim sConn As String

30       sConn = "Driver={MySQL ODBC 3.51 Driver};"
40       If erroDriver Then sConn = "Driver={MySQL ODBC 5.2 ANSI Driver};": tentou = True
50       sConn = sConn & "Server=" & server
60       sConn = sConn & "Database=" & banco
70       sConn = sConn & "User=" & usuario
80       sConn = sConn & "Password=" & senha

'MsgBox sConn
100      CON.Open sConn
         'rsl.CursorLocation = adUseClient
   With Rel
      .DataControl1.ConnectionString = sConn
      .DataControl1.Source = sql
      .Show 1
      'MsgBox ("Mostrei")
   End With

Exit Sub
Erro:

130   If Not tentou Then
         MsgBox ("Vou tentar maiis uma vez")
         erroDriver = True: GoTo Inicio
      Else
         MsgBox Err.Description
      End If
End Sub

Public Function Rotina_AbrirBanco() As Boolean

      Dim server As String
      Dim BDados As String
      Dim NomeUs As String
      Dim PassWD As String
      Dim Driver As String
      Dim erroDriver As Boolean
      Dim tentou As Boolean
      
10 Mouse: On Error GoTo ConnectMQ_Error
Inicio:
20    Driver = "Driver={MySQL ODBC 3.51 Driver};"

30    If erroDriver Then
40       tentou = True
50       Driver = "Driver={MySQL ODBC 5.2 ANSI Driver};"
60    End If
70    server = "Server=mysql.sistemaos.com.br"
80    BDados = "sistemaos03"
90    NomeUs = "sistemaos03"
100   PassWD = "zinholui47"
'   If Compilando Then
'      server = "localhost;"
'      BDados = "Local;"
'      NomeUs = "root;"
'      PassWD = ""
'   End If
110  If db.State = 1 Then db.Close: Set db = Nothing
120   db.Open Driver & server & _
              ";Database= " & BDados & _
              ";User= " & NomeUs & _
              ";Password= " & PassWD & _
              ";Option=3;"
130   Rotina_AbrirBanco = True
140 MouseOff:    Exit Function
ConnectMQ_Error:
'Se não encontrou o Driver 3.51 tenta com o 5.1 (apenas uma vez)
150   If Not tentou Then erroDriver = True: GoTo Inicio
160   Rotina_AbrirBanco = False
End Function
Public Sub FechaDB()

    If glb.State = 1 Then
       glb.Close: Set glb = Nothing
       acGlb = 0
    End If
    If contab.State = 1 Then
       contab.Close: Set contab = Nothing
       acContab = 0
    End If
    If usu.State = 1 Then
       usu.Close: Set usu = Nothing
       acUsu = 0
    End If
    If neg.State = 1 Then
       neg.Close: Set neg = Nothing
       acNeg = 0
    End If
    If hneg.State = 1 Then
       hneg.Close: Set hneg = Nothing
       achNeg = 0
    End If
    If dneg.State = 1 Then
       dneg.Close: Set dneg = Nothing
       acdNeg = 0
    End If
    If hdneg.State = 1 Then
       hdneg.Close: Set hdneg = Nothing
       achdNeg = 0
    End If
    If pes.State = 1 Then
       pes.Close: Set pes = Nothing
       acPes = 0
    End If
    If ctp.State = 1 Then
       ctp.Close: Set ctp = Nothing
    End If
    If hctp.State = 1 Then
       hctp.Close: Set hctp = Nothing
    End If
    If ctr.State = 1 Then
       ctr.Close: Set ctr = Nothing
    End If
    If hctr.State = 1 Then
       hctr.Close: Set hctr = Nothing
    End If
    If nfe.State = 1 Then
       nfe.Close: Set nfe = Nothing
    End If
    If hnfe.State = 1 Then
       hnfe.Close: Set hnfe = Nothing
    End If
    If nfd.State = 1 Then
       nfd.Close: Set nfd = Nothing
    End If
    If hnfd.State = 1 Then
       hnfd.Close: Set hnfd = Nothing
    End If
    If dnfe.State = 1 Then
       dnfe.Close: Set dnfe = Nothing
    End If
    If hdnfe.State = 1 Then
       hdnfe.Close: Set hnfd = Nothing
    End If
    If CartRep.State = 1 Then
       CartRep.Close: Set CartRep = Nothing
    End If
    If CartPromot.State = 1 Then
       CartPromot.Close: Set CartPromot = Nothing
    End If
    If Contato.State = 1 Then
       Contato.Close: Set Contato = Nothing
    End If
    If ICM.State = 1 Then
        ICM.Close: Set ICM = Nothing
     End If
     If UfRegiao.State = 1 Then
        UfRegiao.Close: Set UfRegiao = Nothing
     End If
     If Prod.State = 1 Then
        Prod.Close: Set Prod = Nothing
     End If
     If ProdPco.State = 1 Then
        ProdPco.Close: Set ProdPco = Nothing
     End If
      If Emp.State = 1 Then
        Emp.Close: Set Emp = Nothing
     End If
     If NatuOper.State = 1 Then
        NatuOper.Close: Set NatuOper = Nothing
     End If

     If TabPreco.State = 1 Then
        TabPreco.Close: Set TabPreco = Nothing
     End If
     
     If ProdPreco.State = 1 Then
        ProdPreco.Close: Set ProdPreco = Nothing
     End If
     If UnidEmb.State = 1 Then
        UnidEmb.Close: Set UnidEmb = Nothing
     End If
     If ProdTerc.State = 1 Then
        ProdTerc.Close: Set ProdTerc = Nothing
     End If
     If ProdEntrada.State = 1 Then
        ProdEntrada.Close: Set ProdEntrada = Nothing
     End If
     If ProdFornec.State = 1 Then
        ProdFornec.Close: Set ProdFornec = Nothing
     End If
     If tlanc.State = 1 Then
        tlanc.Close: Set tlanc = Nothing
     End If
     If fpag.State = 1 Then
        fpag.Close: Set fpag = Nothing
     End If
     If uoper.State = 1 Then
        uoper.Close: Set uoper = Nothing
     End If
     If gge.State = 1 Then
        gge.Close: Set gge = Nothing
     End If
     If gdet.State = 1 Then
        gdet.Close: Set gdet = Nothing
     End If
     If unid.State = 1 Then
        unid.Close: Set unid = Nothing
     End If
     If Ativ.State = 1 Then
        Ativ.Close: Set Ativ = Nothing
     End If
     If Rmb.State = 1 Then
        Rmb.Close: Set Rmb = Nothing
     End If
     If RmbDet.State = 1 Then
        RmbDet.Close: Set RmbDet = Nothing
     End If
     If ext.State = 1 Then
        ext.Close: Set ext = Nothing
     End If
     If evt.State = 1 Then
        evt.Close: Set evt = Nothing
     End If
     If lgt.State = 1 Then
        lgt.Close: Set lgt = Nothing
     End If
     If asoe.State = 1 Then
        asoe.Close: Set asoe = Nothing
     End If
     If asoa.State = 1 Then
        asoa.Close: Set asoa = Nothing
     End If
     If cto.State = 1 Then
        cto.Close: Set cto = Nothing
     End If
     If agcto.State = 1 Then
        agcto.Close: Set agcto = Nothing
     End If
     If eqpt.State = 1 Then
        eqpt.Close: Set eqpt = Nothing
     End If
     If eqh.State = 1 Then
        eqh.Close: Set eqh = Nothing
     End If
     If teq.State = 1 Then
        teq.Close: Set teq = Nothing
     End If
     If leq.State = 1 Then
        leq.Close: Set leq = Nothing
     End If
    If ccc.State = 1 Then
       ccc.Close: Set ccc = Nothing
     End If
    If db.State = 1 Then
       db.Close: Set db = Nothing
    End If

End Sub

