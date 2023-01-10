<p align="center">
  <img alt="Repository size" src="https://img.shields.io/github/directory-file-count/marelps/excel-agendamento-de-consulta?style=flat-square">
  <a href="https://twitter.com/piterparquinho">
    <img alt="Siga no Twitter" src="https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2Fpiterparquinho">
  </a>
  <img alt="Github last commit" src="https://img.shields.io/github/last-commit/marelps/excel-agendamento-de-consulta?style=flat-square">
   <img alt="License" src="https://img.shields.io/badge/license-MIT-brightgreen">
  <a href="https://rocketseat.com.br">
    <img alt="Made by vitória" src="https://img.shields.io/badge/made%20by-Vitória-%237519C1">
  </a>

# Appointment scheduling spreadsheet in Excel Macro
<h4 align="center"> 
	✅ Spreadsheet Completed ✅
</h4>

##  Excel spreadsheet with macro button that inserts data for scheduling appointments such as specialties, exams, exam location, date and time.

<p align="center">
 <a href="#about">About</a> •
 <a href="#how-to-use">How to Use</a> •  
 <a href="#Author">Author</a> • 
 <a href="#license">License</a> • 
 <a href="#readme">Versions of README</a>
</p>

## About
This spreadsheet was created when I was an intern at an infectious disease center and felt the need to facilitate some of my main functions there: the registration of appointments in a spreadsheet so that I could have control of what I was doing and be aware of all the appointments that came to me.

The idea was very well embraced by my supervisors because the whole scheduling process still had many flaws that depended not only on our unit but on all the health units in the city. 

In addition, it was a simple document that was possible that other people with different kinds of knowledge with excel or computers in general, could use the same file without bringing any problems with formatting or things written in the wrong places.
  
The specialties, exams and place of performance were already pre-registered for personal use at the time I was an intern in a clinical center, but it is possible to add or remove data through the macros made.

 ## How to Use
 Through a button seen at the top of the spreadsheet, you can open this form and this is where you can register the appointments that came in for me at the internship.

<p align="center">
   <img src="imgs/form.png" alt="Form">
</p>

### Macros
 Macros used in the form. Text written with a ' at the beginning of the line, are some comments I made to localize myself.
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
Macro used to call the form on the button located at the top of the spreadsheet

<p align="center">
   <img src="imgs/button.jpeg" alt="Button in the top of the spreadsheet">
</p>



```
Sub ChamarFormCadastro()
    frmCadastro.Show
End Sub
```
## Author
<p align="center">
 <img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/48718646?v=4" width="100px;" alt="Autora do projeto"/>
 <br />
 <sub><b>Vitória Garrucho</b></br> Made with ❤️</sub></p>

<p align="center">Contact me through my social!<br>
<a href="https://twitter.com/piterparquinho" target="_blank"><img src="https://img.shields.io/badge/-@piterparquinho-1ca0f1?style=flat-square&labelColor=1ca0f1&logo=twitter&logoColor=white&link=https://twitter.com/piterparquinho" alt="Twitter Badge"></a>
<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank"><img src="https://img.shields.io/badge/-Vitória-blue?style=flat-square&logo=Linkedin&logoColor=white&link=https://www.linkedin.com/in/vitoriagarrucho/" alt="Linkedin Badge"></a>
<a href="mailto:vitoriagarrucho@gmail.com" target="_blank"><img src="https://img.shields.io/badge/-vitoriagarrucho@gmail.com-c14438?style=flat-square&logo=Gmail&logoColor=white&link=mailto:vitoriagarrucho@gmail.com" alt="Gmail Badge"></a>
 </p>

## License

This project is under license [MIT](./LICENSE).

Made with ❤️ by Vitória Garrucho

<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank">Contact me!</a>

## README
[Português](./README.md)  |  [English](./README-en.md)
