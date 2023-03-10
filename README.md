<p align="center">
  <img alt="Repository size" src="https://img.shields.io/github/directory-file-count/marelps/excel-agendamento-de-consulta?style=flat-square">
  <a href="https://twitter.com/piterparquinho">
    <img alt="Siga no Twitter" src="https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2Fpiterparquinho">
  </a>
  <img alt="Github last commit" src="https://img.shields.io/github/last-commit/marelps/excel-agendamento-de-consulta?style=flat-square">
   <img alt="License" src="https://img.shields.io/badge/license-MIT-brightgreen">
  <a href="https://rocketseat.com.br">
    <img alt="Feito por Vitória" src="https://img.shields.io/badge/feito%20por-Vitória-%237519C1">
  </a>

# Planilha de agendamento de consulta em Excel Macro
<h4 align="center"> 
	✅ Planilha Concluída ✅
</h4>

##  Planilha em excel com botão macro que insere dados para agendamento de consulta como especialidades, exames, local da realização do exame, data e horário. 

<p align="center">
 <a href="#objetivo">Objetivo</a> •
 <a href="#como-usar">Como Usar</a> •  
 <a href="#autor">Autor</a> • 
  <a href="#licença">Licença</a> • 
 <a href="#readme">Versões do README</a>
</p>

## Objetivo
Essa planilha foi criada na época em que eu estagiava em um centro de infecctologia e senti a necessidade de facilitar algumas das minhas principais funções lá dentro: o cadastro dos agendamentos em uma planilha para que eu pudesse ter controle do que estava fazendo e ciência de todos os agendamentos que chegava até mim. 

A ideia foi muito bem abraçada por meus supervisores pois todo o processo do agendamento ainda tinha muitas falhas que não só dependiam da nossa unidade e sim de todas as unidades de saúde da prefeitura da cidade. 

Além disso, era um documento simples que era possível que outras pessoas com diferentes tipos de conhecimento com excel ou computadores no geral, pudessem utilizar o mesmo arquivo sem trazer nenhum problema com a formatação ou coisas escritas em lugares errados.
  
As especialidades, exames e local da realização foram já pré-cadastradas para uso pessoal na época em que eu fazia estágio em um centro clínico, porém é possível adicionar ou remover dados através das macros feitas.

 ## Como Usar
 Através de um botão visto no topo da planilha, é possível abrir esse formulário e é aqui que é possível cadastrar os agendamentos que chegava para mim no estágio.

<p align="center">
   <img src="imgs/form.png" alt="Form">
</p>

 ### Macros
 Macros utilizadas no formulário. Textos escritos com um ' no começo da linha, são alguns comentários que fiz para me localizar.
 ```
'Identifica o tipo do objeto e insere se for um dos tipos definidos
Private Sub lsInserir(ByRef lTextBox As Variant, ByVal Plan1 As String, ByVal lColunaCodigo As Long, ByVal lUltimaLinha As Long)
    If (TypeOf lTextBox Is MSForms.TextBox) Or (TypeOf lTextBox Is MSForms.ComboBox) Then
        Sheets(Plan1).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Text
    Else
        If TypeOf lTextBox Is MSForms.OptionButton Then
            If lTextBox.Value = True Then
                Sheets(Plan1).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Caption
            End If
        End If
    End If
End Sub

'Loop por todos os componentes da tela
'frmProntuario = Nome do UserForm atual
'Plan1 = Nome da planilha aonde irão ser inseridos os valores
'lColunaCodigo = Coluna de referência para a inserção dos dados
Public Function lsInserirTextBox(frmProntuario As UserForm, ByVal Plan1 As String, ByVal lColunaCodigo As Long)
    Dim controle            As Control
    Dim lUltimaLinhaAtiva   As Long
    
    lUltimaLinhaAtiva = Worksheets(Plan1).Cells(Worksheets(Plan1).Rows.Count, lColunaCodigo).End(xlUp).Row + 1
    
    For Each controle In frmProntuario.Controls
        lsInserir controle, Plan1, lColunaCodigo, lUltimaLinhaAtiva
    Next
End Function

'Limpa todos os objetos TextBox da tela
Public Function lsLimparTextBox(frmProntuario As UserForm)
    Dim controle            As Control
    
    For Each controle In frmProntuario.Controls
        If TypeOf controle Is MSForms.TextBox Then
            controle.Text = ""
        End If
    Next
End Function

'Aciona o botão de limpar
Private Sub CommandButton1_Click()
    lsLimparTextBox frmProntuario
    
    TextBox1.SetFocus
End Sub

'Aciona o botão de inserir
Private Sub CommandButton2_Click()
    lsInserirTextBox frmProntuario, "PRONTUARIO", 2
    
    lsLimparTextBox frmProntuario
    
    TextBox1.SetFocus
End Sub

'Textos em caixa alta
Private Sub TextBox1_Change()
    TextBox1 = UCase(TextBox1)
    'Ucase = Upper case
End Sub
Private Sub TextBox2_Change()
    TextBox2 = UCase(TextBox2) 'Ucase = Upper case
End Sub

Private Sub TextBox3_Change()
    TextBox3 = UCase(TextBox3)
    'Ucase = Upper case
End Sub

 ```
 ***
Macro utilizada para chamar o formulário no botão localizado no topo da planilha

<p align="center">
   <img src="imgs/button.jpeg" alt="Button in the top of the spreadsheet">
</p>


```
Sub ChamarFormCadastro()
    frmCadastro.Show
End Sub
```

## Autor
<p align="center">
 <img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/48718646?v=4" width="100px;" alt="Autora do projeto"/>
 <br />
 <sub><b>Vitória Garrucho</b></br> Feito com ❤️</sub></p>

<p align="center">Entre em contato através das minhas redes sociais!<br>
<a href="https://twitter.com/piterparquinho" target="_blank"><img src="https://img.shields.io/badge/-@piterparquinho-1ca0f1?style=flat-square&labelColor=1ca0f1&logo=twitter&logoColor=white&link=https://twitter.com/piterparquinho" alt="Twitter Badge"></a>
<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank"><img src="https://img.shields.io/badge/-Vitória-blue?style=flat-square&logo=Linkedin&logoColor=white&link=https://www.linkedin.com/in/vitoriagarrucho/" alt="Linkedin Badge"></a>
<a href="mailto:vitoriagarrucho@gmail.com" target="_blank"><img src="https://img.shields.io/badge/-vitoriagarrucho@gmail.com-c14438?style=flat-square&logo=Gmail&logoColor=white&link=mailto:vitoriagarrucho@gmail.com" alt="Gmail Badge"></a>
 </p>

## Licença

Este projeto esta sobe a licença [MIT](./LICENSE).

Feito com ❤️ por Vitória Garrucho

<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank">Entre em contato!</a>

## README
[Português](./README.md)  |  [English](./README-en.md)
