Attribute VB_Name = "NewMacros"
Option Explicit
'module abr�viation avec doublons 18.03.2012
Option Compare Text
Public sOption As Integer
'my repeat sert dans le form MyInputBox pour indiquer la saisie � r�p�tition
'd'abr�viation. Sa valeur est de 10 si on veut cr�er et rester dans le formulaire,
'de 20 dans le cas o� l'on veut sortir imm�diatement
'fonctionencours sert � savoir dans quelle fonction on part (abr�ger ou d�velopper)
'mypbkmsg est la valeur qu'on r�cup�re quand on saisit un mot ou son abr�viation
Public Schemin, sR�f�rence, sR�f�rence2, sTitre, MyLettresNonSon, MyDestFolder, MyTable
Dim MyHeureD�but, MyHeureFin, MyD�butExclusion, MyFinExclusion, IsLettreSansSon, MyPosition, MyPonctuation
Public ResultAddMot, ForcerCr�ation
Public MyPbkMsg, sAcronyme, MyAbr�viation, MyRepeat, MySaisie, FonctionEnCours, MyApostrophe, UsageRechercheMot, MyOldFile As Integer
Public LienPrincipal, LienSecondaire, MyIndexAutocorrect, myab, MyNewWord, zLettresDuMilieuBrutes, MyConjug, MyInfinitif, EndIsConjug, EndIsAccord, Filter
Public terminaisons(), NombreTerminaisons
Public SingleLetter, ChoixAbr�viation, MyAbDansDot, MyId

Public dbNorthwind As DAO.Database
Dim myFootNote As Boolean, MyNomDoc, MyWindow As Window
Public MySameConsonnes, MySelectionInPrevious


Public Sub r�f�rences()



Dim sR�cup As Variant, pdc As Integer, pdr As Integer
Dim Schemin, sR�fs, sTitre As Variant, sPr�R�f�rence As Variant
Dim pnX As Integer, pnY As Integer, sTaille As Integer

Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject

MyDataObject.GetFromClipboard

sR�cup = MyDataObject.GetText
'je r�cup�re tout
'on contr�le d'abord les deux �l�ments qui peuvent manquer, et qui sont signal�s
'par la cha�ne pass�e par Access
'on contr�le d'abord l'existence du chemin
pdc = InStr(1, sR�cup, "pasdechemin")
If pdc <> 0 Then 'cela veut dire qu'il a trouv� ma phrase


'est qu'est-ce qu'on fait ?

End If

'puis l'existence de la r�f�rence

pdr = InStr(1, sR�cup, "pasderef")
If pdr <> 0 Then 'cela veut dire qu'il a trouv� ma phrase


'est qu'est-ce qu'on fait ?

End If

'extraction du chemin
sTaille = Len(sR�cup)
pnX = InStr(1, sR�cup, "xxxx")
Schemin = Left(sR�cup, pnX - 1)




'extraction du titre

pnY = InStr(1, sR�cup, "zzzz")
sTitre = Left(sR�cup, pnY - 1)

'sTaille = sTaille - pnX
sTitre = Replace(sTitre, Schemin & "xxxx", "")


'extraction de la r�f�rence

sR�f�rence = Replace(sR�cup, Schemin & "xxxx", "")

sR�f�rence = Replace(sR�f�rence, sTitre, "")
sR�f�rence = Replace(sR�f�rence, "zzzz", "")


r�cup�ration.Show 'on demande � l'utilisateur ce qu'il veut faire avec les r�f�rences

Select Case sOption

Case 1 'titre complet dans le texte avec hyperlien + nbp avec la r�f�rence

ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        Schemin, SubAddress:="", ScreenTip:="", TextToDisplay:=sTitre



  With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End)
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
    Selection.TypeText Text:=sR�f�rence

Case 2 'titre complet et r�f�rence en note de bas de page

ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        Schemin, SubAddress:="", ScreenTip:="", TextToDisplay:=""


  With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End)
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
    Selection.TypeText Text:=sTitre & " " & sR�f�rence


Case 3 'seulement la r�f�rence en note de bas de page

    With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End)
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
    Selection.TypeText Text:=sR�f�rence
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend

   ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        Schemin, SubAddress:="", ScreenTip:="", TextToDisplay:=""

    ActiveWindow.ActivePane.Close

Case 4

Case 5

ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        Schemin, _
        SubAddress:="", ScreenTip:="", TextToDisplay:=""



End Select

End Sub








Sub sauver_document()

Dim sR�cup As Variant
Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject
Dim MyNumberDocs, i

Dim adoc As Document
MyDataObject.GetFromClipboard
sR�cup = MyDataObject.GetText

MyNumberDocs = Documents.Count

Select Case MyNumberDocs

Case 1
'mettre un message de confirmation

Rename.nom_fichier = sR�cup
Rename.Caption = "Confirmer nom du fichier svp !"
Rename.Show

ActiveDocument.SaveAs filename:=MySaisie, FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False

Case Else

For i = 1 To MyNumberDocs

Set adoc = Documents(i)


SaveFile.MyFiles.AddItem adoc.FullName
   SaveFile.MyFiles.List(i - 1, 1) = i
Next i

'For Each adoc In Documents
 '   MsgBox adoc.
    
'Next adoc
SaveFile.Texte_message = MyNumberDocs & " fichiers sont ouverts. Choisir le fichier � renommer (l'ancien sera sauvegard�)"






SaveFile.bouton_confirmer.Enabled = False
SaveFile.Show
 
Rename.nom_fichier = Trim(sR�cup)
Rename.Caption = "Confirmer nom du fichier svp !"
Rename.Show
Set adoc = Documents(MyOldFile)
adoc.Activate

ActiveDocument.SaveAs filename:=MySaisie

'''''''''''''''''''''''''''''''''''''


End Select
''''''''''''''''''''''''''''''''''''''''''''''

End Sub





Sub supprimer_paragraphe()
'
' supprimer_paragraphe Macro
' Macro enregistr�e le 23/07/2006 par Emmanuel BARBE
Dim MyDataObject As MSForms.DataObject
Dim MyMsg As Integer
Set MyDataObject = New MSForms.DataObject '

MyMsg = MsgBox("appliquer � tout le texte (oui) ou seulement � la s�lection (non) ?", 4, "domaine")


If MyMsg = 6 Then Selection.WholeStory '

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .CorrectHangulEndings = True
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     CommandBars("Stop Recording").Visible = False
    
MyMsg = MsgBox("Changer la casse et contr�ler l'orthographe ? (oui ou non) ou sortir (annule)", 3, "casse")

If MyMsg = 2 Then Exit Sub

If MyMsg = 1 Then '
Selection.Range.Case = wdLowerCase
    If Options.CheckGrammarWithSpelling = True Then
        ActiveDocument.CheckGrammar
    Else
        ActiveDocument.CheckSpelling
    End If
 End If
 
    
    
Selection.WholeStory
toto:
If Selection.Characters.Count > 249 Then
  Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Union europ�enne"
        .Replacement.Text = "UE"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute

End If
Selection.WholeStory
If Selection.Characters.Count > 249 Then

MsgBox "trop de caract�res", vbAbortRetryIgnore, Selection.Characters.Count

End If


MyDataObject.SetText Selection


MyDataObject.PutInClipboard
    
End Sub
Sub nbp()
'
' nbp Macro
' Macro enregistr�e le 25/12/2006 par Sony
'
    With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End)
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
End Sub


Sub mail_sauvegarde()
'
' mail_sauvegarde Macro
' Macro enregistr�e le 25/12/2006 par Sony
'
    ActiveDocument.SendMail



    
    
End Sub
Sub changer_casse()
'
' changer_casse Macro
' Macro enregistr�e le 30/12/2006 par Sony
'
    Selection.Range.Case = wdLowerCase
    If Options.CheckGrammarWithSpelling = True Then
        ActiveDocument.CheckGrammar
    Else
        ActiveDocument.CheckSpelling
    End If
End Sub
Sub mettre_subrillance_en_nbp()
'
' Macro5 Macro
' Macro enregistr�e le 31/12/2006 par Sony
'
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "("
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    CommandBars("Stop Recording").Visible = False

Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ")"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    CommandBars("Stop Recording").Visible = False




    Selection.Cut
    With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End)
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
    Selection.PasteAndFormat (wdPasteDefault)
    ActiveWindow.ActivePane.Close
End Sub

Public Sub mes_macros()
macros.Show
End Sub
Sub entr�e_glossaire()
'
' Macro6 Macro
' Macro enregistr�e le 01/01/2007 par Sony
'
    Selection.Copy
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Selection
        .Replacement.Text = Selection & "*"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    Documents.Open filename:="C:\EJE\livre\glossaire.doc"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.PasteAndFormat (wdPasteDefault)
    ActiveDocument.Save
    ActiveDocument.Close
End Sub
Sub imposer_le_fran�ais()
'
' imposer_le_fran�ais Macro
' Macro enregistr�e le 02/01/2007 par Sony
'
    Selection.LanguageID = wdFrench
    Selection.NoProofing = False
    Application.CheckLanguage = True
End Sub

Public Sub pac()

NormalTemplate.AutoTextEntries("pac").Insert WHERE:=Selection.Range, _
        RichText:=True

Selection.MoveLeft Unit:=wdCharacter, Count:=1


End Sub
Sub chercher_pac()
'
' chercher_pac Macro
' Macro enregistr�e le 03/01/2007 par Sony
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ChrW(61654)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub bleu()
'
' bleu Macro
' Macro enregistr�e le 19/02/2007 par Sony
'
    Application.DisplayStatusBar = True
    Application.ShowWindowsInTaskbar = True
    Application.ShowStartupDialog = True
    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayLeftScrollBar = False
        .StyleAreaWidth = 0
        .DisplayRightRuler = False
        .DisplayScreenTips = True
        With .View
            .ShowAnimation = True
            .Draft = False
            .WrapToWindow = False
            .ShowPicturePlaceHolders = False
            .ShowFieldCodes = False
            .ShowBookmarks = False
            .FieldShading = wdFieldShadingWhenSelected
            .ShowTabs = False
            .ShowSpaces = False
            .ShowParagraphs = False
            .ShowHyphens = False
            .ShowHiddenText = False
            .ShowAll = False
            .ShowDrawings = True
            .ShowObjectAnchors = False
            .ShowTextBoundaries = False
            .ShowHighlight = True
            .DisplayPageBoundaries = True
            .DisplaySmartTags = True
        End With
    End With
    With Options
        .Pagination = True
        .WPHelp = False
        .WPDocNavKeys = False
        .ShortMenuNames = False
        .RTFInClipboard = True
        .BlueScreen = True
        .EnableSound = True
        .ConfirmConversions = False
        .UpdateLinksAtOpen = True
        .SendMailAttach = True
        .MeasurementUnit = wdPoints
        .AllowPixelUnits = False
        .AllowReadingMode = True
        .AnimateScreenMovements = True
        .VirusProtection = False
        .ApplyFarEastFontsToAscii = False
        .InterpretHighAnsi = wdHighAnsiIsHighAnsi
        .BackgroundOpen = False
        .AutoCreateNewDrawings = True
    End With
    Application.DisplayRecentFiles = True
    RecentFiles.Maximum = 9
End Sub

Public Sub r�cup_nomcomplet()
Dim MyDataObject As MSForms.DataObject
Dim MyMsg As Integer
Set MyDataObject = New MSForms.DataObject
Dim MyPath As String



MyDataObject.SetText ActiveDocument.FullName



MyDataObject.PutInClipboard

End Sub


Public Sub Abr�ger2()
'd�but sub
'm�thode_ab.Hide

FonctionEnCours = "abr�ger"
UsageRechercheMot = ""

'Hypoth�se de d�part : on a tap� un mot � l'�cran et on veut l'abr�ger. Ce programme marche pour un mot.
'il faut encore traiter la question des acronymes. Appel�e par "control f"





Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject
Dim Myentry, MyWord, MyMsg, MySuggestion As String
Dim mySpelling, MySpellingAb As Boolean
Dim MyAutoCorrects, i, MyReplaceEntry, j, sNombre, sPasDeSuggestion As Integer, k
Dim MySpell As Dictionary
Dim MyValue, MyApos, MyLong, MyCar, MyTxtForInput, MySpace, MyWord1
Dim sTailleNom, MyWord2, MyAbr�viation2, MyWord3
Dim LastAb, ThisComputer

'Dim DicCustom As Dictionaries
'Set DicCustom = Application.CustomDictionaries.ActiveCustomDictionary
'MsgBox DicCustom.Name


Dim MyDate

'MyDate = GetSetting("fasttype", section:="param�tres", Key:="date_usage")
ThisComputer = GetSetting("fasttype", section:="param�tres", Key:="cet ordinateur")
LastAb = get_settings_from_bdd(3)
If ThisComputer <> LastAb Then



extraire_abr�viations





End If


MyReplaceEntry = 0
Dim MyActiveDocument As Document, MyExistingAb
MyExistingAb = 0

Set MyActiveDocument = Application.ActiveDocument
' la phase ci-dessous permet de r�cup�rer la valeur de l'abr�viation.

''''''''''''''''

MyWord = Selection.Text 'c'est � dire un mot complet, cad encore la valeur de l'abr�viation

If Len(MyWord) <= 1 Then


Selection.MoveLeft Unit:=wdCharacter, Count:=1
Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
MyWord = Selection.Text




MyWord = LCase(MyWord)
MyWord = Replace(MyWord, " ", "")

Else
MyWord = LCase(Selection.Text)

MyWord = Trim(MyWord)


End If

'Dim y
MyWord = Trim(MyWord)



If Len(MyWord) < 2 Then


        Select Case MayBeAlone(MyWord)
        
        
            Case -1
            cr�er_lettres_seules
            Case 0
            sMessage "la lettre :" & Chr(10) & Chr(10) & MyWord & Chr(10) & Chr(10) & " ne peut servir d'abr�viation car est elle est signifiante", "annuler", "rien", "rien", "rien", "Lettre signifiante", 255, 0
            Exit Sub
        
        
        End Select 'MayBeAlone(MyWord)
        
        
        
        Exit Sub
        
        Else 'on recherche si le mot n'est pas un groupe d'expression
        
        MySpace = InStrRev(MyWord, " ")
        
        Do While MySpace > 0
        MySpace = InStrRev(MyWord, " ", MySpace)
        
        Loop
        
        
            If MySpace > 0 Then
                MyLong = Len(MyWord)
                MyWord1 = Left(MyWord, MySpace - 1)
                MyWord = Right(MyWord, MyLong - MySpace)
            
            
            End If 'MySpace > 0
            
        
        
        
End If 'Len(MyWord) < 2
recommencer:

    
       
 chercher_abr�viation_existante (MyWord)

'II : on v�rifie l'orthographe du mot � abr�ger



MySpellingAb = Application.CheckSpelling(MyWord)


       Select Case MySpellingAb
        
          Case False 'l'orthographe est mauvaise. On va proposer les corrections possibles,
        ' et par d�faut la premi�re fournie par le correcteur
   
            chercher_suggestion_mot (MyWord)
   

        
   
        
        
        End Select 'MySpellingAb

'chercher_utilisation_valeur Myword

'II : on demande l'abr�viation.
askab:
'MyValue = InputBox("abr�viation pour " & MyWord, "abr�viation de " & MyWord)
    
'If MyValue = "" Then Exit Sub 'l'utilisateur a choisi d'annuler
'recommencer:
MyInputBox.zone_mot = MyWord

If MyExistingAb > 0 Then
MyTxtForInput = MyWord & " est d�j� abr�g� " & MyExistingAb & " fois. Vous pouvez supprimer certaines ou toutes les abr�viations existantes ou en cr�er une nouvelle"
MyInputBox.bouton_annuler.TabIndex = 1
MyInputBox.bouton_ok_et_sortir.TabIndex = 2
'MyInputBox.zone_abr�viation_existantes.Enabled = True
'MyInputBox.zone_abr�viation_existantes.Visible = True
'MyInputBox.bouton_supprimer_abr�viation.Visible = True

Else

'pour l'instant, on pose l'hypoth�se que si le mot � abr�ger n'est pas dans le dictionnaire, il
'n'a pas de de nom d'abr�viation.



MyTxtForInput = "abr�viation pour " & MyWord
MyInputBox.bouton_ok_et_sortir.TabIndex = 1
MyInputBox.zone_abr�viation.TabIndex = 2

If MySpellingAb = False Then
MyTxtForInput = "abr�viation pour " & MyWord & ". Attention, ce mot n'est pas dans le dictionnaire."
MyInputBox.texte.ForeColor = 255

End If 'MySpellingAb = False
End If

OpenMyInputBox MyTxtForInput, MyWord 'en l'occurrence le mot � abr�ger, cad la valeur de l'abr�viation.

'le formulaire retourne les valeurs

MyValue = MyPbkMsg 'myvalue repr�sente le nom de l'abr�viation '

If MyRepeat = 30 Then 'cette valeur est pass�e par le champ zone_suggestion_orthographe
'cela veut dire qu'on annulle le processus d'abr�viation

   MyActiveDocument.Activate
   ' Selection.TypeText Text:=MyPbkMsg
    
   ' Selection.MoveRight Unit:=wdCharacter, Count:=1
    Exit Sub

End If 'MyRepeat = 30

MyPbkMsg = MyAbr�viation

MyWord = MyAbr�viation '

If MyValue = 0 Then Exit Sub
MyId = extraire_id(MyWord, "table_m�re")
' AutoCorrect.Entries.Add MyValue, MyWord
 stocker_abr�viations MyValue, MyWord, False, False, MyId
enseigner_abr�viations MyValue, MyWord
 
 Select Case MyRepeat
 
 Case 10 'on souhaite cr�er un nouveau couple abr�viation/mot
MyInputBox.zone_abr�viation_existantes.Clear
MyExistingAb = 0

chercher_utilisation_valeur MyWord
'chercher_utilisation_abr�viation (MyValue)

GoTo recommencer

 End Select
 

End Sub

Public Sub d�velopper2()
ForcerCr�ation = "oui"
d�velopper_espace

End Sub

Public Sub sMessage(message, bouton1, bouton2, bouton3, bouton4, titre, couleurtexte, boutond�faut)

'le forme renvoie une valeur MyPbkMsg qui contient 1, 2, 3 ou 4 selon le bouton objet du clic

Dim sbouton As String


mymsgbox.texte = message


If bouton1 <> "rien" Then
mymsgbox.bouton1.Caption = bouton1
mymsgbox.bouton1.Visible = True
Else
mymsgbox.bouton1.Visible = False
End If

If bouton2 <> "rien" Then
mymsgbox.bouton2.Caption = bouton2
mymsgbox.bouton2.Visible = True
Else
mymsgbox.bouton2.Visible = False
End If

If bouton3 <> "rien" Then
mymsgbox.bouton3.Caption = bouton3
mymsgbox.bouton3.Visible = True
Else
mymsgbox.bouton3.Visible = False
End If

If bouton4 <> "rien" Then
mymsgbox.bouton4.Caption = bouton4
mymsgbox.bouton4.Visible = True
Else
mymsgbox.bouton4.Visible = False
End If

mymsgbox.Show


End Sub

Public Sub OpenMyInputBox(texte, sAbr�viation)
'ReadSetup



MyInputBox.texte = texte
MyInputBox.bouton_acronyme.Caption = "acronyme " & get_accord("acronyme")



Select Case FonctionEnCours

Case "abr�ger"

MyInputBox.zone_mot = sAbr�viation
'suggestion (sAbr�viation)'!!!!! fonction de suggestion d'abr�viation � revoir
'la valeur myrepeat vient des boutons d'inputbox, et sert dans l'hypoth�se
'o� l'on veut entrer plusieurs abr�viations � la suite
MyInputBox.zone_abr�viation.TabIndex = 0


If MyRepeat = 10 Then MyInputBox.zone_abr�viation = MySaisie

Case "d�velopper"

MyInputBox.zone_abr�viation = sAbr�viation
MyInputBox.zone_mot.TabIndex = 0
If MyRepeat = 10 Then MyInputBox.zone_mot = MySaisie

End Select 'FonctionEnCours


MyInputBox.Caption = "Entrez une valeur"
peupler_ab_similaires myab, "myinputbox"

MyInputBox.Show





End Sub




Public Sub ReadSetup()



 Dim fso          ' As Scripting.FileSystemObject
 Dim ts           ' As Scripting.TextStream

 Dim strline, bret, strDest As String, mysettings, intsettings
 Dim s, MyEnd, MyBegin, MyAutoCorrects, j, MyName, MyValue, k, kk, l
 
 
 
  
  

MyInputBox.terminaisons.Clear
MyInputBox.sons.Clear
'MyInputBox.zone_abr�viation = ""

k = 0


 'MyAutoCorrects = AutoCorrect.Entries.Count
 l = 1
 mysettings = GetAllSettings(appname:="fasttype", section:="sons")
    For intsettings = LBound(mysettings, 1) To UBound(mysettings, 1)
    
    MyValue = mysettings(intsettings, 0)
    MyName = mysettings(intsettings, 1)
       ' Debug.Print mysettings(intsettings, 0), mysettings(intsettings, 1)
        
        MyInputBox.sons.AddItem MyValue
        MyInputBox.sons.List(l - 1, 1) = MyName
        
    l = l + 1
    Next intsettings
       
        
        
        
        
     '   For j = 1 To MyAutoCorrects
    l = 1
  mysettings = GetAllSettings(appname:="fasttype", section:="terminaisons")
          For intsettings = LBound(mysettings, 1) To UBound(mysettings, 1)
        

            MyValue = mysettings(intsettings, 0)
            MyName = mysettings(intsettings, 1)
          
            MyInputBox.terminaisons.AddItem MyValue
            MyInputBox.terminaisons.List(l - 1, 1) = MyName
                  
                 
       
         
                
    l = l + 1
         Next intsettings




End Sub


Public Sub cr�er_lettres_seules()


'cette fonction ouvre le formulaire "lettres_seules".
'il extrait les valeurs d'abr�viation de chaque lettre pouvant servir d'abr�viation
'� partir du fichier d'abr�viation.
'il identifie les lettres � partir des contr�les du formulaires.

Dim fso          ' As Scripting.FileSystemObject
 Dim ts           ' As Scripting.TextStream

Dim strline, bret, strDest As String
Dim s, MyEnd, MyBegin, snom, sLongueur, sValeur
Dim myfile, MyControls, i, mynamecontrol
  



'''''''''''''''''''''

  MyControls = lettres_seules.Controls.Count
   For i = 0 To MyControls - 1
   mynamecontrol = lettres_seules.Controls(i).Name
   If Len(mynamecontrol) = 1 Then
    snom = lettres_seules(i).Name
    If check_existence_valeur_pour_abr�viation(snom) Then
    sValeur = AutoCorrect.Entries(snom).Value
    lettres_seules.Controls(i) = sValeur
    End If
    End If
    Next
    
   
lettres_seules.Show

   
End Sub



Public Sub ouvrir_accueil()


 Dim rdshippers As Recordset
 
 
 'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
 Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")
 
 
accueil.nombre_ab = rdshippers.RecordCount
 
 



accueil.Show
End Sub

Public Function MayBeAlone(sLetter)
'cette fonction extrait des champs pr�sents dans le form "lettres_seules"
'les lettres pouvant constituer seules une abr�viation, cad n'ayant pas de
'signifiant dans la langue fran�aise.

Dim MyControls, i, mynamecontrol, snom

 MyControls = lettres_seules.Controls.Count
   For i = 0 To MyControls - 1
   mynamecontrol = lettres_seules.Controls(i).Name
   If Len(mynamecontrol) = 1 Then
    snom = lettres_seules(i).Name
        If sLetter = snom Then
        MayBeAlone = -1
        Exit Function
        End If
        
    
    End If
    
    Next

MayBeAlone = 0

End Function
Sub affecter_touches()
'
' affecter_touches Macro
'
'
Dim MyKey

    
    
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyK, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="abr�ger2"
    
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="d�velopper2"
        
         
          CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="cr�er_lettres_seules"
        
        
     CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyM, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="load_m�thode"
        
        
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyM, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ouvrir_marqueur"
   
        
        
End Sub



Public Sub stocker_abr�viations(myab, MyValue, cr�ation_auto As Boolean, JamaisDansRegistre As Boolean, Id)


Dim fso ' As Scripting.FileSystemObject

 Dim ts, fd, fsp   ' As Scripting.TextStream

 Dim str, sFileName, sFileName2, sFile, sFichier, sExiste, sFichier2 ' As String
 
 Dim snom, sValeur, sLigne, mycontrolsn, i, MyControls, mynamecontrol, test
 
 Dim docNew As Document
'Dim dbNorthwind As DAO.Database
Dim rdshippers As Recordset
Dim MyParam�tres As Recordset
Dim intRecords 'As Integer

 'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
      Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")
    Set MyParam�tres = dbNorthwind.OpenRecordset("param�tres")
  
  With MyParam�tres
  .Index = "PrimaryKey"
  .Seek "=", 1
  .Edit
  !date_usage = Date
  !LastAb = get_param�tres("cet ordinateur")
  !MyHeure = Time
  .Update
  End With
    
test = abr�viation_existe(myab, MyValue)

 Select Case abr�viation_existe(myab, MyValue) ' l'abr�viation n'existe pas
'si abr�viation_existe renvoie -11, cela veut dire que l'abr�viation et sa valeur existent d�j�. Elles ne devraient pas �tre enregistr�es en doublon dans
'la base ni dans le registre �videmment
    
 Case 0 'l'abr�viation n'existe pas :
    
    
    With rdshippers
    
   .AddNew
   !nom = Trim(myab)
    !valeur = Trim(MyValue)
    !cr�ation_auto = cr�ation_auto
        If JamaisDansRegistre = False Then
        !registre = -1
        Else
        !registre = 0
        End If
    !jamais_dans_registre = JamaisDansRegistre
    !r�f�rence = Id
    !valeur_lettres_ab = valeur_lettres_abr�viations(myab)
    !taille = Len(myab)
    .Update

    End With
    
    
    'on cr�� l'abr�viation dans le registre
    
    If JamaisDansRegistre = False Then AutoCorrect.Entries.Add myab, MyValue
    
    
    Case -1  'l'abr�viation existe
    
    'alors il faut mettre le registre de l'ancienne � 0 et cr�er une nouvelle abr�viation dont le registre sera aussi sur la valeur 0
    
    With rdshippers
    .Index = "nom"
    .Seek "=", myab
    .Edit
    !registre = 0
    .Update
    
    End With
    
    With rdshippers
    
    .AddNew
     !nom = Trim(myab)
     !valeur = Trim(MyValue)
    !cr�ation_auto = cr�ation_auto
    !registre = 0
    !jamais_dans_registre = JamaisDansRegistre
    !r�f�rence = Id
    !valeur_lettres_ab = valeur_lettres_abr�viations(myab)
    !taille = Len(myab)
    .Update
    
    End With
    
    
    'on supprime l'abr�viation dans le registre
    'on teste d'abord si l'abr�viation existe pour �tre certain de ne pas provoquer une erreur
    
    If check_existence_valeur_pour_abr�viation(myab) = True Then
    
    AutoCorrect.Entries(myab).Delete
    End If
    
    
    
    
    
  End Select
  
 
    
  'rdShippers.Close
  'dbNorthwind.Close
   





'UpDateLastAb


End Sub


Public Sub open_acroynymes()
modif_son_terminaisons.Caption = "acronymes"
modif_son_terminaisons.Label_valeur.Caption = "nom complet"
modif_son_terminaisons.Label_nom.Caption = "abr�viation"
modif_son_terminaisons.bouton_supprimer.Visible = False
modif_son_terminaisons.Show
End Sub

Public Sub calcul_�conomies()
Dim fso ' As Scripting.FileSystemObject

 Dim ts, fd, fsp   ' As Scripting.TextStream

 Dim str, sFileName, MyAverage ' As String
 
 Dim snom, sValeur, sLigne, mycontrolsn, i, j, MyControls, mynamecontrol, MyTotal, MyLenNom, MyLenValeur, mytotallen, cr��es, strsql, MyConjugaisons_rares
 Dim MyCr�ation_scripteur As Recordset
 Dim rdshippers As Recordset
 
 
 
 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
 Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")
 
 
Statistiques.nombre_total_abr�viations = rdshippers.RecordCount
 
 
 
    MyTotal = 0
   MyControls = AutoCorrect.Entries.Count
   For i = 1 To MyControls - 1

    MyLenNom = Len(AutoCorrect.Entries(i).Name)
    MyLenValeur = Len(AutoCorrect.Entries(i).Value)
    mytotallen = MyLenValeur - MyLenNom
    MyTotal = mytotallen + MyTotal
    MyAverage = Round(MyTotal / i, 1)
    
   
    Statistiques.Nombre_frappes = MyTotal
    Statistiques.Moyenne_�conomie = MyAverage
    
    Next
    
    
    Set rdshippers = dbNorthwind.OpenRecordset("Dates_cr�ation_abr�viations")
  
              
           If rdshippers.BOF = True Then Exit Sub
                        
                        
        rdshippers.MoveFirst
        i = 0
        While rdshippers.EOF = False
        Statistiques.Abr�viations_jour.AddItem rdshippers.Fields(0)
        Statistiques.Abr�viations_jour.List(i, 1) = rdshippers.Fields(1)
        i = i + 1
        rdshippers.MoveNext
        Wend
      
    
 strsql = "SELECT Count(abr�viations.cr�ation_auto) AS CompteDecr�ation_auto, abr�viations.cr�ation_auto FROM abr�viations GROUP BY abr�viations.cr�ation_auto HAVING (((abr�viations.cr�ation_auto)=-1));"
 Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 
 Statistiques.abr�viations_automatiques = rdshippers.Fields("CompteDecr�ation_auto")
 Statistiques.nombre_abr�viations_scripteur = Statistiques.nombre_total_abr�viations - Statistiques.abr�viations_automatiques
 Statistiques.multiplicateur = Round(Statistiques.abr�viations_automatiques / Statistiques.nombre_abr�viations_scripteur, 1)
 strsql = "SELECT Count(abr�viations.registre) AS CompteDeregistre FROM abr�viations GROUP BY abr�viations.registre HAVING (((abr�viations.registre)=0));"

 
 
 
  Set rdshippers = dbNorthwind.OpenRecordset(strsql)
Statistiques.conjugaisons_rares = rdshippers.Fields("compteDeregistre")
 
 
                            
Statistiques.Show
                                       
                        
                                
      

   
   
' MsgBox "l'ensemble de vos " & i & "  abr�viation repr�sente " & MyTotal & " frappes �conomis�es, soit un gain moyen de " & MyAverage & " lettres par mot", vbCritical, "total"
End Sub

Public Sub extraire_abr�viations()



 
 Dim strline, bret, strDest As String
 Dim s, MyEnd, MyBegin, MyAutoCorrects, j, MyName, MyValue, filename, sDelete, CompteR�cup, i
 Dim MyIndex
Dim folder, subflds, fld, fl As file, MyLen, MyInternalIndex, strsql
 ' Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
 
 
 sDelete = 0
 CompteR�cup = 0
 'il faut d'abord effacer toutes les abr�viations du fichier des abr�viations,
 'car sinon cela doublonne
 
 MyAutoCorrects = AutoCorrect.Entries.Count
' MsgBox myautocorrects
 
 If IsEmpty(MyAutoCorrects) Then GoTo skip:
encore_une_fois:
 For j = 1 To MyAutoCorrects

'Debug.Print AutoCorrect.Entries(1).Name & " " & AutoCorrect.Entries(1).Value

  AutoCorrect.Entries(1).Delete 'en fait, il faut toujours effacer l'entr�e qui porte l'index 1,
  'car les index sont renum�rot�s apr�s chaque effacement, de sorte qu'il y en a toujours 1 qui porte le num�ro 1.

 'MyAutocorrects = AutoCorrect.Entries.Count
'If MyAutocorrects < 1 Then GoTo skip

 'sDelete = MyAutocorrects - j
Next j
skip:

 MyAutoCorrects = AutoCorrect.Entries.Count 'on recompte le nombre d'autocorrect
 
 If MyAutoCorrects > 0 Then GoTo encore_une_fois 'si ce n'est pas vide, on recommence.


'Debug.Print "nombre entr�es effac�es " & sDelete
'Debug.Print "nombre d'entr�es restantes " & AutoCorrect.Entries.Count
'Debug.Print "nombre d'entrees compt�es " & myautocorrects
'Debug.Print "valeur de j " & j


'''''''''''''''''''''''''''''''''''

accueil.nombre_ab = MyAutoCorrects

'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")

strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.registre FROM abr�viations WHERE (((abr�viations.registre)=-1));"

    
    
    
    MyIndex = MyIndex + 1
        

 
   
    
  i = 0
    
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

    With rdshippers
    .MoveFirst
      Do While Not .EOF
  
        If rdshippers.Fields("nom") <> "" And rdshippers.Fields("valeur") <> "" Then
         
        
        AutoCorrect.Entries.Add rdshippers.Fields("nom"), rdshippers.Fields("valeur")
        accueil.nombre_ab = i
        i = i + 1
        
        End If
        
         .MoveNext
      Loop
   End With


''''''''''''''''''''''''''''''''''''''''


' Set rdshippers = dbNorthwind.OpenRecordset("param�tres")
'
'
'  With rdshippers
'  .Index = "PrimaryKey"
'  .Seek "=", 1
'  .Edit
'  !date_usage = Date
'  !LastAb = get_param�tres("cet ordinateur")
'  !MyHeure = Time
'  .Update
'  End With
'


MsgBox i
MsgBox i & " abr�viations r�cup�r�es du fichier de stockage", vbOKOnly, "FasType"


'SaveSetting appname:="fasttype", section:="param�tres", Key:="date_usage", setting:=Date
'MsgBox CompteR�cup & "  abr�vations ont �t� r�cup�r�es", vbOKOnly, "r�cup�ration des abr�viations"





Exit Sub




End Sub

Public Sub doc_local()
Dim Schemin, sString, sString2, sString3, sString4, sdocument As Document


sString2 = "http://www.ue.espacejudiciaire.net/docs/"
sString3 = "http://www.ue.espacejudiciaire.net/docsprives/"
sString4 = "f:\eje\docsprives\"
'schemin = sdocument.
Schemin = Selection.Hyperlinks.Item(1).Address
If InStr(1, Schemin, sString2) Then
Schemin = Replace(Schemin, sString2, sString4, 1)
Else
Schemin = Replace(Schemin, sString3, sString4, 1)
End If
ActiveDocument.FollowHyperlink (Schemin)

Exit Sub


End Sub
Public Sub save_file()
Dim sR�cup As Variant, pdc As Integer, pdr As Integer
Dim Schemin, sR�fs, sTitre As Variant, sPr�R�f�rence As Variant
Dim pnX As Integer, pnY As Integer, sTaille As Integer

Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject

MyDataObject.GetFromClipboard

sR�cup = MyDataObject.GetText
ActiveDocument.SaveAs filename:=sR�cup, FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False



End Sub



Public Sub document_keypress(SPACE)
MsgBox "salut"
End Sub
Sub dico()
'
' dico Macro
'
'
End Sub
Sub zoom()
'
' zoom Macro
' Macro enregistr�e le 13/04/2009 par Emmanuel BARBE
'
    ActiveWindow.ActivePane.View.zoom.PageFit = wdPageFitBestFit
End Sub
Sub poser_marque()
'
' poser_marque Macro
' Macro enregistr�e le 15/04/2009 par Emmanuel BARBE
'
    Selection.TypeText Text:="[Point � contr�ler]"
End Sub
Sub chercher_marque()
'
' chercher_marque Macro
' Macro enregistr�e le 15/04/2009 par Emmanuel BARBE
'

End Sub

Public Sub ouvrir_marqueur()
FonctionEnCours = "ouvrir_marqueur"
load_marqueurs

Load marqueurs
 marqueurs.Show
          
End Sub

Public Sub chercher_marqueur()
Dim j, MyName, MyValue, s, MyAutoCorrects
FonctionEnCours = "chercher_marqueur"


load_marqueurs
marqueurs.texte = "double-cliquer sur le marqueur � rechercher"
marqueurs.�l�ments.AddItem "rechercher tous les marqueurs"
marqueurs.bouton_ajouter.Visible = False
marqueurs.bouton_supprimer.Visible = False
   
  
   
   
marqueurs.Show
End Sub
Sub Macro1()
'
' Macro1 Macro
' Macro enregistr�e le 16/04/2009 par Emmanuel BARBE
'
    Selection.EndKey Unit:=wdStory
    Selection.MoveLeft Unit:=wdCharacter, Count:=45, Extend:=wdExtend
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorBlue
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub

Public Sub remplacements()
'cette macro permet de substituer la premi�re lettre du nom d'une abr�viation,
'pour le f�minmin pluriel etc.
Dim MyAutoCorrects, j, MyName, MyValue


 MyAutoCorrects = AutoCorrect.Entries.Count
        
        
        For j = 1 To MyAutoCorrects - 2
        
      

        If InStrRev(AutoCorrect.Entries(j).Name, ";") = 1 Then
        
            
            MyName = Replace(AutoCorrect.Entries(j).Name, ";", "y", , 1)
            
            MyValue = AutoCorrect.Entries(j).Value
            AutoCorrect.Entries.Add MyName, MyValue
            
            AutoCorrect.Entries.Item(j).Delete
            
            
          
         End If
         
          
          Next j
End Sub


Public Sub chercher_abr�viation_existante(MyWord)
Dim rdshippers As Recordset
Dim h, strsql As String
MyInputBox.zone_abr�viation_existantes.Clear

If MyWord = "" Then MyWord = MyInputBox.suggestions


strsql = "SELECT abr�viations.valeur, abr�viations.nom FROM abr�viations WHERE (((abr�viations.valeur) Like """ & MyWord & """));"







   
 
Set rdshippers = dbNorthwind.OpenRecordset(strsql)
                                                    
If rdshippers.BOF = True Then Exit Sub
         
         rdshippers.MoveFirst
        While rdshippers.EOF = False
                MyInputBox.zone_abr�viation_existantes.AddItem rdshippers.Fields("nom").Value
                rdshippers.MoveNext
                        
        Wend

                   
                     
                                      
                            
                               
                               
                               
                 
                                                             
                                                       

         
                     

 
 
 

End Sub

Public Sub chercher_utilisation_valeur(MyWord)
'cette fonction rechercher pour les afficher dans le champ MyInputBox.zone_mots_approchants
'les utilisations du mot qui est dans la zone_mot de ce m�me formulaire


Dim sTailleNom, MyWord1, MyWord2, i, MySuggestion, k, MyAutoCorrects, myab
 MyInputBox.zone_mots_approchants.Clear
MyAutoCorrects = AutoCorrect.Entries.Count
sTailleNom = Len(MyWord)


Select Case sTailleNom

Case 1
GoTo SkipRecherche

Case 2
MyWord1 = Left(MyWord, sTailleNom - 1)

Case Else
MyWord1 = Left(MyWord, sTailleNom - 2)
End Select 'sTailleNom




MyWord1 = MyWord1 & "*"

MyWord2 = "*" & MyWord & "*"
 
        
        For i = 1 To MyAutoCorrects
      
          MySuggestion = LCase(AutoCorrect.Entries(i).Value)
            'on recherche les corrections
            'puis on regarde si elles ont fait l'objet d'une abr�viation
            'et on les ins�re dans la zone_suggestion_orthographe du form myinputbox
            'on indique s'il y a une abr�viation ou non
            'If MySuggestion Like "11111*" Or MySuggestion Like "12345*" Or MySuggestion Like "56879*" Then GoTo SkipIf
            If MySuggestion Like MyWord1 Or MySuggestion Like MyWord2 Then
            k = k + 1
            
            'MySuggestion = AutoCorrect.Entries(i).Value
            myab = AutoCorrect.Entries(i).Name
            MyInputBox.zone_mots_approchants.AddItem MySuggestion
            MyInputBox.zone_mots_approchants.List(k - 1, 1) = myab
           
            
            End If
            
SkipIf:
        Next 'i
        



        
                
        
        
MyInputBox.listes_d�roulantes.Value = 0
        
       
SkipRecherche:
 
End Sub

Public Function abr�viation_existe(myab, MyPbkMsg)
'cette fonction retourne :
' 0 si l'abr�viation n'existe pas
' -11 si l'abr�viation existe avec la m�me valeur
'-1 si l'abr�viation existe avec une valeur diff�rente
Dim rdshippers As Recordset
Dim strsql
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")


strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """) AND ((abr�viations.valeur)=""" & MyPbkMsg & """));"

Set rdshippers = dbNorthwind.OpenRecordset(strsql)
              
If rdshippers.RecordCount >= 1 Then
        abr�viation_existe = -11
            '  "existe avec la m�me valeur"
        Exit Function
        
        
    Else

        strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """));"
        
        Set rdshippers = dbNorthwind.OpenRecordset(strsql)
            If rdshippers.RecordCount >= 1 Then
            abr�viation_existe = -1
            'existe avec une valeur diff�rente
            Else
            abr�viation_existe = 0
            'n'existe pas
            End If
            


End If

End Function

Public Sub chercher_suggestion_mot(MyWord)
Dim sNombre, i, MySuggestion, MyAutoCorrects, j
MyInputBox.zone_suggestion_orthographe.Clear

sNombre = Application.GetSpellingSuggestions(MyWord).Count
        
                If sNombre = 0 Then
      '
          '
                
                End If 'snombre = O
        
        For i = 1 To sNombre
         
            'on recherche les corrections
            'puis on regarde si elles ont fait l'objet d'une abr�viation
            'et on les ins�re dans la zone_suggestion_orthographe du form myinputbox
            'on indique s'il y a une abr�viation ou non
            
             MySuggestion = Application.GetSpellingSuggestions(MyWord).Item(i)
             MyInputBox.zone_suggestion_orthographe.AddItem MySuggestion
             MyInputBox.zone_suggestion_orthographe.List(i - 1, 1) = "pas d'abr�viation"
                
                 MyAutoCorrects = AutoCorrect.Entries.Count
                                  
                    For j = 1 To MyAutoCorrects
        
                    If AutoCorrect.Entries(j).Value = MySuggestion Then
                    MyInputBox.zone_suggestion_orthographe.List(i - 1, 1) = AutoCorrect.Entries(j).Name
              
                    End If 'AutoCorrect.Entries(j).Value = MySuggestion
        
        
          Next 'j
                
        Next 'i
MyInputBox.listes_d�roulantes.Value = 1
End Sub

Public Sub suggestion(MyWord)

'fonction d'abr�viation automatique des mots en fonctions des modalit�s d�finies par l'utilisateur
'quant � la mani�re d'abr�ger un mot


Dim ChercheEspace, MySuggestion
Dim sPresent, sPresent1, sTailleMot, MyAutoCorrects, i, SonFinale, LettreFinale, TailleFinale, sSon, LettreSon, MotSansFinale
Dim MonMotInitial, Pr�sentDansMotInitial, VoyelleDansSon, LettresDansSons, PositionVoyelle, k, l
Dim MyPropositionFinale, VoyelleAutourS, j, MyAutocorectsVoyelles, VoyelleDroite, VoyelleGauche, ChDansMot
Dim A, b, kk, kkk, kkkk, MyExistingAb, TailleMyWord, MyWordSansS, MyNewWord, TailleMyAb
Dim MyOrthographe

MySuggestion = ""
k = 0
MonMotInitial = MyWord





'il faut d'abord traiter la question des mots qui comportent des espaces, dont il faut retenir la premi�re lettre
sTailleMot = Len(MyWord)
ChercheEspace = InStrRev(MyWord, "  ", 1) ' cela ne marche pas car il ne d�tecte pas l'espace
    Do While ChercheEspace >= 1
    MySuggestion = MySuggestion & Left(MyWord, 1)
    MyWord = Right(MyWord, sTailleMot - ChercheEspace)
    ChercheEspace = InStrRev(MyWord, " ", 1)
    
    Loop
    


'dans les am�liorations
'
' il faut d'abord voir si un mot au pluriel a d�j� une abr�viation au singulier

 MyAutoCorrects = AutoCorrect.Entries.Count
  MyExistingAb = 0
        For j = 1 To MyAutoCorrects
        
        If AutoCorrect.Entries(j).Value = MyWord Then 'la valeur existe d�j�
        
        GoTo prochaine_v�rif:
            
        Else 'AutoCorrect.Entries(j).value = MyAb
        TailleMyWord = Len(MyWord)
        MyWordSansS = Right(MyWord, 1)
            
            If MyWordSansS = "s" Then
            
            If AutoCorrect.Entries(j).Value = Left(MyWord, TailleMyWord - 1) Then
            
                MyNewWord = AutoCorrect.Entries(j).Name & "h"
                MyInputBox.zone_abr�viation = MyNewWord
                Exit Sub
                
                    
            
            
            End If 'AutoCorrect.Entries(j).Name = Left(MyAb, TailleMyAb - 1)
            
            End If 'myAbSansS = Right(MyAb, 1) = "h"
            
        
         
         
              
         End If 'AutoCorrect.Entries(j).Name = myAb
        
       
        
        Next 'j
   
  

prochaine_v�rif:


' il faut ensuite voir si un mot au singulier a d�j� une abr�vation au pluriel

MyAutoCorrects = AutoCorrect.Entries.Count
  MyExistingAb = 0
        For j = 1 To MyAutoCorrects
        
        If AutoCorrect.Entries(j).Value = MyWord & "s" Then '
        
        TailleMyAb = Len(AutoCorrect.Entries(j).Value)
        MyWordSansS = Left(AutoCorrect.Entries(j).Value, TailleMyAb - 1)
               
                    
                    MyOrthographe = Application.CheckSpelling(MyWordSansS)
                    
                        Select Case MyOrthographe
                        
                            Case True
                            TailleMyAb = Len(AutoCorrect.Entries(j).Name)
                            
                            MyInputBox.zone_abr�viation = Left(AutoCorrect.Entries(j).Name, TailleMyAb - 1)
                            Exit Sub
                            Exit Sub
                        
                        
                        End Select 'myOrthographe
            
            
            'End If 'myAbSansS = Right(MyAb, 1) = "h"
            
        
         
         
              
         End If 'AutoCorrect.Entries(j).Name = myAb
        
       
        
        Next 'j
   
  
 

'
'I : traitement des mots qui ne comportent pas d'espace


LettresDansSons = ""
LettreFinale = ""
MonMotInitial = MyWord

'A. traitement des finales
'la m�thode consiste � rechercher toutes les finales d�finies par l'utilisateur pour voir
'si le mot la contient.
'dans l'affirmative, la finale est remplac�e par sa lettre.
'il faut encore traiter la question du pluriel.

MyAutoCorrects = AutoCorrect.Entries.Count

For i = 1 To MyAutoCorrects

If AutoCorrect.Entries(i).Name Like "56789*" Then
'les name des abr�viations sont class�s par ordre alphab�tique.
'donc, quand il a lu tous les name qui commence par 56789,
'c'est inutile de continuer jusqu'au bout car cela ralentit le processus
'Le but est de d�tecter quand, apr�s un name commen�ant par 56789,
'on passe � un name ne commen�ant pas par ce groupe de lettres;
'donc, tant qu'on n'a pas atteint le groupe de name 56789,
'la valeur kk reste � z�ro
'quand on lit le groupe de name 56789, la valeur k passe � 1 et s'incr�mente (pour permettre
'des v�rifications sur le nombre d'occurrences
'quand k = est � une valeur sup�rieure � 0 et KK reprend une valeur positive
'alors cela signifie qu'on repasse � des name ne commen�ant pas par 56789
'donc, on peut sortir le la boucle de lecture des autoroccerct.entries

k = k + 1
kk = 0

'ici

    SonFinale = Replace(AutoCorrect.Entries(i).Name, "56789", "")
    
 If SonFinale = "er" Then
 SonFinale = SonFinale
 
 End If 'SonFinale = "er"
 
    
'SonFinale = Replace(SonFinale, "12345", "")
'LettreFinale = Replace(AutoCorrect.Entries(i).Value, "56789", "")


TailleFinale = Len(SonFinale)


sPresent = InStrRev(MyWord, SonFinale)
If sTailleMot - TailleFinale < 1 Then GoTo skip
sTailleMot = Len(MyWord)
If sPresent = sTailleMot - (TailleFinale - 1) Then
MyWord = Left(MyWord, sTailleMot - TailleFinale)

'les finales sont stock�es dans le fichier des abr�viations avec la valeur 56789 avant le nom comme avant la valeur
 
    LettreFinale = Replace(AutoCorrect.Entries(i).Value, "56789", "")
  
  
GoTo FinaleTrait�e:
End If


    Else ' AutoCorrect.Entries(i).Name not Like "56789*

kk = 1
If k = 1 And kk = 1 Then GoTo FinTraitementFinales

End If 'AutoCorrect.Entries(i).Name Like "56789" Then


skip:


Next 'i = 1 To myautocorrects


FinaleTrait�e:

FinTraitementFinales:
k = 0

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'attention, il faut r�int�grer la finale qui s'appelle sFinale
'les traitements suivants se font hors finale
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


'I.B traitement des sons interm�diaires
'extraction du premier son
'on va chercher dans le mot sans sa finale si se trouvent les sons qui ont �t� d�finis par l'utilisateur.
'le pb est d'�viter que la substitution cr��e de nouveaux sons qui viennent parasiter le processus.
'pour ce faire, on v�rifier d'abord si le son �tait pr�vu dans le mot initial, qui est donc pass� dans une
'variable au tout d�but du sub. C'est seulement dans ce cas-l� qu'on proc�de � la substitution
'

For i = 1 To MyAutoCorrects

If AutoCorrect.Entries(i).Name Like "12345*" Then

k = k + 1
kk = 0

sSon = Replace(AutoCorrect.Entries(i).Name, "12345", "")

LettreSon = Replace(AutoCorrect.Entries(i).Value, "12345", "")
'ici, il faut travailler sur la partie sans la finale

Pr�sentDansMotInitial = InStrRev(MonMotInitial, sSon) 'il faut �viter une substitution li�e � la cr�ation de sons
'suite � une premi�re substitution, d'o� l'importance de v�rifier dans le mot initial si le mot � substituer
'est ou non dans le mot qu'on est en train d'abr�ger

If Pr�sentDansMotInitial > 0 Then

'il faut introduire une exception pour le son ch qui s'il n'est pas entour�
'de voyelles se prononce k et non pas che.
    
    Select Case sSon 'on prend cette technique pour rajouter, le cas �ch�ant, des case
    
    Case "ch"
    
    ChDansMot = InStrRev(MyWord, sSon)
    
    If ChDansMot = 0 Then GoTo remplace
    
    
    k = 0
        For l = 1 To MyAutoCorrects
        
        If AutoCorrect.Entries(l).Name Like "99999*" Then
        
        VoyelleDroite = AutoCorrect.Entries(l).Value
        
        'If VoyelleDroite = "o" Then
        'k = 0
        'GoTo lettreO
        'End If
        
        
        'MyAutocorectsVoyelles = AutoCorrect.Entries.Count
            
           ' For j = 1 To MyAutocorectsVoyelles
              '  If AutoCorrect.Entries(j).Name Like "99999*" Then
                ' VoyelleDroite = AutoCorrect.Entries(j).Value
            
            ChDansMot = InStrRev(MyWord, "ch" & VoyelleDroite)
            If ChDansMot = 0 Then
            'cela veut dire qu'il faut passer � la voyelle suivante
            
            GoTo nextl
            
            
            Else
            
            
            GoTo remplace:
            End If
            'MyWord = Replace(MyWord, VoyelleGauche & "s" & VoyelleDroite, VoyelleGauche & "z" & VoyelleDroite)
                'End If
                
            

            
           ' Next j 'For j = 1 To MyAutocorectsVoyelles
        
        
        
        End If 'AutoCorrect.Entries(i).Name Like "99999*"
        
nextl:
        Next l 'For i = 1 To myautocorrects
lettreO:
        
If k = 0 Then LettreSon = "k"
    
    End Select
    



remplace:

MyWord = Replace(MyWord, sSon, LettreSon)
LettresDansSons = LettresDansSons & LettreSon

End If 'Pr�sentDansMotInitial > 0
'ensuite, il faut virer toutes les voyelles sauf celles qui sont dans le son de remplacement

Else 'AutoCorrect.Entries(i).Name not Like "12345*"

kk = 1
If k = 1 And kk = 1 Then GoTo FinTraitementSons

End If 'AutoCorrect.Entries(i).Name Like "12345*"
SkipCh:
Next i '= 1 To myautocorrects
FinTraitementSons:
k = 0
'I.C : suppression des voyelles

'grosso modo, la m�thode d'abr�viation la plus simple, une fois qu'on a trait� les finales et les sons,
'consister � enlever les voyelles.
'toutefois, il faut faire attention � deux �cueils :
'a) ne pas enlever des voyelles qui ont �t� substitu�es � des sons.
'b) ne pas enlever la voyelle quand elle est en premi�re lettre dans le mot, cas souvent,
'on en a besoin pour abr�ger.
'pas trouv� d'autre moyen que de r�p�ter le truc pour chaque voyelle...


'� ce stade, il faudrait transformer les "voyelle / s / voyelle en z
'pour ce faire, les voyelles sont entr�es dans le fichier des abr�viations
'sous la forme nom = 99999 + la voyelle valeur = voyelle
'on passe en boucle les voyelles � gauche de s
'puis au sein d'une boucle voyelles � droite
'on recherche si le son voyelle & s & voyelle existe
'si c'est le cas, on la remplace par z.

For i = 1 To MyAutoCorrects

If AutoCorrect.Entries(i).Name Like "99999*" Then
k = k + 1
kk = 0


VoyelleGauche = AutoCorrect.Entries(i).Value

MyAutocorectsVoyelles = AutoCorrect.Entries.Count
 kkk = 0
 kkkk = 0
    For j = 1 To MyAutocorectsVoyelles
        If AutoCorrect.Entries(j).Name Like "99999*" Then
        kkk = kkk + 1
        kkkk = 0
        
        
        VoyelleDroite = AutoCorrect.Entries(j).Value
        MyWord = Replace(MyWord, VoyelleGauche & "s" & VoyelleDroite, VoyelleGauche & "z" & VoyelleDroite)
        
        Else 'AutoCorrect.Entries(j).Name  note Like "99999*" Then
        
        kkkk = 1
        If kkk = 1 And kkkk = 1 Then GoTo FinTraitementVOyelleDroite
       
        End If
        
    
    
    
    Next j 'For j = 1 To MyAutocorectsVoyelles
FinTraitementVOyelleDroite:



Else 'AutoCorrect.Entries(i).Name not Like "99999*"
 
kk = 1
If k = 1 And kk = 1 Then GoTo FinTraitementVoyellesGauches


End If 'AutoCorrect.Entries(i).Name Like "99999*"
Next i 'For i = 1 To myautocorrects

FinTraitementVoyellesGauches:

k = 0


VoyelleDansSon = InStrRev(LettresDansSons, "a")
PositionVoyelle = InStrRev(MyWord, "a", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "a" & Replace(MyWord, "a", "", 2)
         
         Else
         MyWord = Replace(MyWord, "a", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "e")
PositionVoyelle = InStrRev(MyWord, "e", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "e" & Replace(MyWord, "e", "", 2)
         
         Else
         MyWord = Replace(MyWord, "e", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "i")
PositionVoyelle = InStrRev(MyWord, "i", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "i" & Replace(MyWord, "i", "", 2)
         
         Else
         MyWord = Replace(MyWord, "i", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "o")
PositionVoyelle = InStrRev(MyWord, "o", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "o" & Replace(MyWord, "o", "", 2)
         
         Else
         MyWord = Replace(MyWord, "o", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "u")
PositionVoyelle = InStrRev(MyWord, "u", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "u" & Replace(MyWord, "u", "", 2)
         
         Else
         MyWord = Replace(MyWord, "u", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "y")
PositionVoyelle = InStrRev(MyWord, "y", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "y" & Replace(MyWord, "y", "", 2)
         
         Else
         MyWord = Replace(MyWord, "y", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1

End If 'VoyelleDansSon = 0

VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
    
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1
        

End If 'VoyelleDansSon = 0


VoyelleDansSon = InStrRev(LettresDansSons, "�")
PositionVoyelle = InStrRev(MyWord, "�", 1)
If VoyelleDansSon = 0 Then
    
        If PositionVoyelle = 1 Then

         MyWord = "�" & Replace(MyWord, "�", "", 2)
         
         Else
         MyWord = Replace(MyWord, "�", "")
         
         End If 'PositionVoyelle = 1
        

End If 'VoyelleDansSon = 0

'virer les doubles consonnes
MyPropositionFinale = MyWord & LettreFinale

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "tt")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "tt", "t")


Pr�sentDansMotInitial = InStrRev(MonMotInitial, "pp")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "pp", "p")

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "mm")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "mm", "m")

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "rr")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "rr", "r")

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "nn")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "nn", "n")

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "ff")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "ff", "f")

Pr�sentDansMotInitial = InStrRev(MonMotInitial, "ll")
If Pr�sentDansMotInitial <> 0 Then MyWord = Replace(MyPropositionFinale, "ll", "l")

 MyInputBox.zone_abr�viation = MyWord & LettreFinale
End Sub

Public Sub essai_tableau()

Dim MyAutoCorrects, i
MyAutoCorrects = AutoCorrect.Entries.Count
ReDim myarray(MyAutoCorrects, 1)
For i = 1 To MyAutoCorrects
If AutoCorrect.Entries(i).Name Like "99999*" Then
'myarray(i,1) = AutoCorrect.Entries(i).Value ; AutoCorrect.Entries(i).name



End If




Next i
End Sub

Public Sub aaa_liste_fonctions()

'chercher_utilisation_abr�viation MyAb
'chercher_abr�viation_existante MyValue

End Sub
Sub signet_automatique()
On Error GoTo erreur
'
' signet_automatique Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'



Dim MyWord, mytaille, MyNiveauTitre
Dim MyLetter, i, MySignets, j
Selection.HomeKey Unit:=wdStory

Selection.Bookmarks.ShowHidden = True
MySignets = ActiveDocument.Bookmarks.Count
For j = 1 To MySignets
'MsgBox ActiveDocument.Bookmarks(j).Name
ActiveDocument.Bookmarks(j).Delete


Next
Selection.Bookmarks.ShowHidden = False
 
sMessage "Indiquer le niveau de titres pour les indexation", "1", "2", "3", "4", "Niveau d'indexation", "bleu", 0

MyNiveauTitre = MyPbkMsg

 
For i = 1 To MyNiveauTitre


 
Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
       
        Select Case i
        
        
         Case 1
           .OutlineLevel = wdOutlineLevel1
         Case 2
         .OutlineLevel = wdOutlineLevel2
         Case 3
         .OutlineLevel = wdOutlineLevel3
         Case 4
         .OutlineLevel = wdOutlineLevel4
         
         Case 5
         
         .OutlineLevel = wdOutlineLevel5
         End Select
         
       With Selection.Find
       .Text = " "
       End With
      
    End With
    
    
    Selection.Find.Execute
    
While Selection.Find.Found
reprise:

Selection.Expand Unit:=wdSentence

 MyLetter = Chr(39)
'MyLetter = Left(MyWord, 1)
'MyLetter = Replace(MyLetter, """", "")
MyWord = Selection.Text
mytaille = Len(MyWord)
MyWord = Replace(MyWord, "  ", "")
MyWord = Replace(MyWord, " ", "_")
MyWord = Replace(MyWord, ".", "")
MyWord = Replace(MyWord, ";", "")
MyWord = Replace(MyWord, ",", "")
MyWord = Replace(MyWord, ":", "")
MyWord = Replace(MyWord, "/", "")
MyWord = Replace(MyWord, "]", "")
'MyWord = Replace(MyWord, myletter, "")
MyWord = Replace(MyWord, " ", "")
MyWord = Replace(MyWord, "?", "")
MyWord = Replace(MyWord, " ", "")
MyWord = Replace(MyWord, """", "")
MyWord = Replace(MyWord, "'", "")
MyWord = Replace(MyWord, "!", "")



MyWord = Left(MyWord, mytaille - 1)
MyWord = Trim(MyWord)


    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=MyWord
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With



Selection.MoveRight Unit:=wdWord, Count:=5

Selection.Find.Execute
Wend
Selection.HomeKey Unit:=wdStory
Next



erreur:
If err = 5828 Then

Clean_index.zone_texte = MyWord

Load Clean_index
Clean_index.Show

Select Case MyPbkMsg

Case 1 'skip
Selection.Find.Execute
Resume

Case 2 ' stop
Exit Sub
Case 3 'renvoi le signet nettoy�
MyWord = MyPbkMsg
Resume
End Select



'GoTo reprend:

    
Else

Resume Next

 End If

End Sub
Sub apostrophe()
'
' apostrophe Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .OutlineLevel = wdOutlineLevel2
    End With
    Selection.Find.ParagraphFormat.Borders.Shadow = False
    With Selection.Find
        .Text = "'"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute
End Sub
Sub chercher_titres()
'
' chercher_titres Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .OutlineLevel = wdOutlineLevel2
    End With
    With Selection.Find.ParagraphFormat
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 0
            .DistanceFromLeft = 0
            .DistanceFromBottom = 0
            .DistanceFromRight = 0
            .Shadow = False
        End With
    End With
    With Selection.Find
        .Text = " "
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub chercher_titre()
'
' chercher_titre Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .OutlineLevel = wdOutlineLevel2
    End With
    With Selection.Find.ParagraphFormat
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 0
            .DistanceFromLeft = 0
            .DistanceFromBottom = 0
            .DistanceFromRight = 0
            .Shadow = False
        End With
    End With
    With Selection.Find
        .Text = " "
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection.Find
        .Text = " "
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub trouver_titre()
'
' trouver_titre Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .OutlineLevel = wdOutlineLevel2
    End With
    With Selection.Find.ParagraphFormat
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 0
            .DistanceFromLeft = 0
            .DistanceFromBottom = 0
            .DistanceFromRight = 0
            .Shadow = False
        End With
    End With
    With Selection.Find
        .Text = " "
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub s�lectionner_paraphagraphe_essai()
'
' s�lectionner_paraphagraphe_essai Macro
' Macro enregistr�e le 03/07/2009 par Emmanuel BARBE
'
    Selection.TypeBackspace
End Sub
Sub sendkey()
'
' sendkey Macro
'
'
    Selection.HomeKey Unit:=wdStory
End Sub
Sub Macro2()
'
' Macro2 Macro
' Macro enregistr�e le 06/07/2009 par Emmanuel BARBE
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .OutlineLevel = wdOutlineLevel1
    End With
    Selection.Find.ParagraphFormat.Borders.Shadow = False
    With Selection.Find
        .Text = " "
        .Replacement.Text = "'"
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub cr�er_lien_hypertexte()
'
' Macro3 Macro
' Macro enregistr�e le 06/07/2009 par Emmanuel BARBE
'
 On Error GoTo erreur:
Dim sR�cup, pdc As Integer, pdr As Integer
Dim Schemin, sR�fs, sTitre As Variant, sPr�R�f�rence As Variant
Dim pnX As Integer, pnY As Integer, sTaille As Integer, i

Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject

MyDataObject.GetFromClipboard

sR�cup = MyDataObject.GetText


Dim MyTarget As Document


Set MyTarget = Application.Documents(sR�cup)
Select Case MyTarget.Bookmarks.Count
Case 0
   ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=sR�cup
        
        MsgBox "lien cr�� ; aucun lien secondaire n'existe dans le document ", vbInformation, "Cr�ation d'un lien hypertexte"
        Exit Sub
    Case Else
For i = 1 To MyTarget.Bookmarks.Count
'signets.signet.AddItem MyTarget.Bookmarks(i).Name

Next

End Select

'Load signets
'signets.Show









    
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
       sR�cup, _
        SubAddress:=LienSecondaire, ScreenTip:=""
        
        MsgBox "lien cr��", vbInformation, "Cr�ation d'un lien hypertexte"
        
        
erreur:
If err = 4160 Then

sMessage "Vous n'avez pas s�lectionn� de fichier cible. Collecter le nom dans le fichier cible", "ok", "ok", "ok", "ok", "Pas de fichier", "bleu", 1
Exit Sub


End If
End Sub

Public Sub essai_base_registre()
Dim mysettings, intsettings

'  SaveSetting appname:="fasttype", section:="param�tres", Key:="Nom du disque", setting:="f"
'  SaveSetting appname:="fasttype", section:="accords", Key:="singulier_pluriel", setting:="h"
'  SaveSetting appname:="fasttype", section:="accords", Key:="f�minin", setting:="f"
'  SaveSetting appname:="fasttype", section:="accords", Key:="f�minin_pluriel", setting:="w"
'  SaveSetting appname:="fasttype", section:="accords", Key:="acronyme", setting:="!"
'  SaveSetting appname:="fasttype", section:="param�tres", Key:="d�velopper", setting:="g"
'  SaveSetting appname:="fasttype", section:="param�tres", Key:="abr�ger", setting:="j"
SaveSetting appname:="fasttype", section:="param�tres", Key:="date_usage", setting:=Date



  
  
MsgBox GetSetting(appname:="fasttype", section:="param�tres", Key:="Nom du disque")
End Sub








Public Function get_accord(section)
Select Case section

Case "singulier_pluriel"
get_accord = GetSetting(appname:="fasttype", section:="accords", Key:=section)
Case "f�minin"
get_accord = GetSetting(appname:="fasttype", section:="accords", Key:=section)
Case "f�minin_pluriel"
get_accord = GetSetting(appname:="fasttype", section:="accords", Key:=section)
Case "acronyme"
get_accord = GetSetting(appname:="fasttype", section:="accords", Key:=section)

End Select

End Function



Public Sub open_settings()
settings.acronyme = get_accord("acronyme")
settings.singulier_pluriel = get_accord("singulier_pluriel")
settings.f�minin = get_accord("f�minin")
settings.f�minin_pluriel = get_accord("f�minin_pluriel")
settings.disque = get_hd
settings.d�velopper = get_param�tres("d�velopper")
settings.abr�ger = get_param�tres("abr�ger")
settings.m�thode = get_param�tres("m�thode")


If get_param�tres("AddMot") = "vrai" Then
settings.ajouter_mot_�_dictionnaire = True
Else
settings.ajouter_mot_�_dictionnaire = False
End If

settings.cet_ordinateur = get_param�tres("cet ordinateur")
settings.ordinateur_last_save = get_param�tres("ordinateur last saving")
settings.date_last_save = get_param�tres("date_usage")

Load settings
settings.Show


End Sub

Public Function get_hd()

get_hd = GetSetting(appname:="fasttype", section:="param�tres", Key:="Nom du disque")
 
End Function

Public Function get_param�tres(fonction)

get_param�tres = GetSetting(appname:="fasttype", section:="param�tres", Key:=fonction)







End Function

Public Sub recherche_mot_depuis_abr�viation(myab)
'On Error GoTo error

Dim MyIndex, filename, fso, ts, s, MyEnd, MyName, MyValue, dicCustom, MyTrouv�Premier, SearchFile, j, NumberLettresDuMilieu, NombreSonLettresDuMilieu
Dim LastLetter, TailleMyAb, FirstLetter, TwoLastLetters, ThreeLastLetters, NombreLignesFichier, myaccord
Dim MySettingAccords, i, MyTerminaison1, MyTerminaison2, MyFirstLetterVoyelle, MyAbExistante, MyNumberVerbe
Dim ComparaisonString, LastLetterComparaison, LettresDuMilieu, TailleLettresDuMilieu, AbDeuxLettres, EndIsTerminaison, LettersConjug, strsql
Dim LettresDuMilieuBrutes
Dim LettreDuMilieu()
ReDim LettreDuMilieu(1 To 10)
'Dim myab '� virer quand on branchera la fonction
Dim z
ReDim terminaisons(1 To 20)
Dim SonFirstLetter, l
Dim SonMiddleLetter
Dim FilterTerminaisons, MyTable1
ReDim SonMiddleLetter(1 To 15)
ReDim SonFirstLetter(1)
Dim MyFirstLetterIsSound
MyFirstLetterIsSound = 0
'Dim LettresDuMilieu(100)
Dim NombreValeurLettreDuMilieu(20)
EndIsConjug = 0
EndIsTerminaison = 0
EndIsAccord = 0
myaccord = 0
MyIndex = 0
myaccord = 0
NombreSonLettresDuMilieu = 0
'myab = InputBox("abr�viation ?")
TailleMyAb = Len(myab)
LastLetter = Right(myab, 1)
TwoLastLetters = Right(myab, 2)
ThreeLastLetters = Right(myab, 3)
FirstLetter = Left(myab, 1)
MyInputBox.sons_examin�s.Clear
MyInputBox.terminaisons_examin�es.Clear
MyInputBox.lettresMilieu.Clear
MyInputBox.rejet�s.Clear

MyInputBox.fichiers_examin�s = ""
MyInputBox.temps_recherche = ""




If UsageRechercheMot = "chercher_�_nouveau" Then
MyInputBox.suggestions.Clear
End If

'on identifie si la premi�re lettre est un son
  ' Dim dbNorthwind As Database
Dim rdshippers As Recordset
   SonFirstLetter(1) = FirstLetter 'la premi�re lettre doit �tre cherch�e �galement pour sa valeur propre et pas seulement pour ce qu'elle abr�ge
 
 'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
 
 strsql = "SELECT m�thode_ab.Valeur, m�thode_ab.Abr�viation FROM m�thode_ab WHERE (((m�thode_ab.Abr�viation)=""" & FirstLetter & """) AND ((m�thode_ab.d�but_mot)=Yes));"

 
 
 Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If rdshippers.BOF = True Then GoTo SkipSonD�but
    ReDim Preserve SonFirstLetter(rdshippers.RecordCount + 1)
       rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                               
                                      MyFirstLetterIsSound = MyFirstLetterIsSound + 1
                                     SonFirstLetter(MyFirstLetterIsSound + 1) = Trim(rdshippers.Fields("valeur").Value)
                                    
                                                              
                                    rdshippers.MoveNext
                        
                                Wend
 
'Ancienne version � partir de la base de registre ; en cours de suppression
 
'MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") '
'
'    For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
'       SonFirstLetter(1) = FirstLetter
'                If Trim((MySettingAccords((i), 1))) = FirstLetter Then
'                MyFirstLetterIsSound = MyFirstLetterIsSound + 1
'                SonFirstLetter(MyFirstLetterIsSound + 1) = Trim((MySettingAccords((i), 0)))
'   '
'                End If
'
'
'      Next i
 
SkipSonD�but:
MyFirstLetterIsSound = MyFirstLetterIsSound + 1


'I : d�termination de la lettre finale. Elle peut �tre soit dans les terminaisons, soit dans les accords.
'si elle est dans les accords, alors elle il faut prendre en compte la deuxi�me finale. Donc, on commence d'abord par v�rifier
'dans les accords.
Dim MyBox
MyBox = MyPbkMsg
'la constante mypbmsg contient une valeur utilis�e ensuite
'donc, je la stocke provisoirement pour la remettre en l'�tat dans quelques lignes,
'afin de pouvoir utiliser la fonction smessage.






         MySettingAccords = GetAllSettings(appname:="fasttype", section:="accords") '
        
         For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                If Trim((MySettingAccords((i), 1))) = LastLetter Then
                'virer ce message
               ' sMessage "voulez-vous le " & Replace(MySettingAccords((i), 0), "_", " ") & " de ce mot ou cherchez-vous un mot finissant par " & LastLetter & " ?", "rien", Replace(MySettingAccords((i), 0), "_", " "), LastLetter & " ou ce qu'elle abr�ge", "rien", "Quel sens donner � la lettre " & LastLetter & " ?", "bleu", 3
                    
                  ' MyHeureD�but = Timer
 
                    
'                    Select Case MyPbkMsg
'                        Case 2 ''on veut chercher le mot accord� en genre et en nombre
'
                  
                            EndIsAccord = -1
                            LastLetterComparaison = "*" 'dans le cas o� il y a un accord, il faut laisser un ast�risque pour r�cup�rer tous les mots qui ont une autre forme
                            'verbale derri�re
                            
                            LastLetter = Mid(myab, TailleMyAb - 1, 1)
                            GoTo skip:
                                    
'                        Case Else 'on veut chercher un mot finissant par une lettre de genre
'
'
'                        GoTo skip
'
'                    End Select 'Case MyPbkMsg
                
                End If 'Trim((MySettingAccords((i), 1))) = LastLetter

        Next i


 
skip:
MyPbkMsg = MyBox
If EndIsAccord = -1 Or MyPbkMsg = 3 Then GoTo SkipConjugaison

'si myaccord = -1, la lettre finale est une terminaison
 


'''''''''''''''''''''''''''''''''' 712012

Set rdshippers = dbNorthwind.OpenRecordset("temps_combinaison")

With rdshippers
        .Index = "nom"
        .Seek "=", Right(myab, 2)

End With

If rdshippers.NoMatch = False Then


                EndIsConjug = -1
               
                MyConjug = Right(myab, 2)
                
                GoTo skip2

End If

With rdshippers
        .Index = "nom"
        .Seek "=", Right(myab, 3)

End With

If rdshippers.NoMatch = False Then


                EndIsConjug = -1
               
                MyConjug = Right(myab, 3)
                
                GoTo skip2

End If



skip2:
If MyConjug <> "" Then 'si on veut d�clencher aussi cela pour la lettre r finale, il faut voir ensuite au niveau des r�sultats
'on passe l'information qu'il y a une conjugaison
'MyTerminaison1 = GetSetting(appname:="fasttype", section:="conjugaisons_deuxi�me", Key:=MyConjug)
'MyTerminaison2 = GetSetting(appname:="fasttype", section:="conjugaisons_premier", Key:=MyConjug)


'EndIsConjug = -1 '

End If 'MyConjug <> ""
SkipConjugaison:


         
               
            'si on a une conjugaison, il faut enlever deux ou trois lettres � droite de myab
       
        If EndIsConjug = -1 And MyPbkMsg = 2 Then
        LettresDuMilieuBrutes = Left(myab, Len(myab) - Len(MyConjug))
        LettresDuMilieuBrutes = Right(LettresDuMilieuBrutes, Len(LettresDuMilieuBrutes) - 1) 'il faut enlever le r qui repr�sente l'infinitif
        GoTo skip9
        End If
         
            'si on a un accord, il faut enlever deux lettre � droite de myab (l'autre �tant la terminaison)
            
        If EndIsAccord = -1 Then '
        LettresDuMilieuBrutes = Left(myab, TailleMyAb - 2)
        LettresDuMilieuBrutes = Right(LettresDuMilieuBrutes, Len(LettresDuMilieuBrutes) - 1)
           GoTo skip9
         End If
         
        LettresDuMilieuBrutes = Left(myab, TailleMyAb - 1)
        LettresDuMilieuBrutes = Right(LettresDuMilieuBrutes, Len(LettresDuMilieuBrutes) - 1)
        '� enlever ensuite
        


 'd�termination de l'existence de lettres du milieu qui ne sont pas abr�viatives
         
skip9:
         
  TailleLettresDuMilieu = Len(LettresDuMilieuBrutes)

skip2094:
     

      
    If LettresDuMilieuBrutes = "" Then
    AbDeuxLettres = -1
    
    
    
    GoTo skip_lettres_du_milieu

    
    End If
    
    If InStr(LettresDuMilieuBrutes, "z") Then zLettresDuMilieuBrutes = -1 ' zLettresDuMilieuBrutes est une variable
    'globale
   
 
   
   ReDim ArrayMiddleLetters(20, 10)
   Dim k, kbis
    For i = 1 To TailleLettresDuMilieu 'on peuple la premi�re colonne avec les lettres du milieu
           
           'If kBis <> -1 Then
          
           k = k + 1
           'kBis = -1
           'End If
          
           'la variable kbis sert � stocker dans la ligne 0 de la colonne de chaque lettre
           'le nombre de possibilit� pour chaque lettre de la combinaison.
           'la multiplication de ces nombres donne le nombre de lignes n�cessaires pour les combinaisons
           'possibles ainsi que la mani�re de les remplir.
          
                    
                    If kbis = 0 Then
                         kbis = 1
                    Else
                         kbis = k
                    End If
          
          
             LettreDuMilieu(i) = Trim(Mid(LettresDuMilieuBrutes, i, 1))
             ArrayMiddleLetters(k, i) = Trim(Mid(LettresDuMilieuBrutes, i, 1))
             ArrayMiddleLetters(kbis, 0) = 1
             NombreValeurLettreDuMilieu(i) = 1
                
                
            strsql = "SELECT m�thode_ab.Valeur, m�thode_ab.Abr�viation FROM m�thode_ab WHERE (((m�thode_ab.Abr�viation)=""" & Trim(Mid(LettresDuMilieuBrutes, i, 1)) & """) AND ((m�thode_ab.milieu_mot)=Yes));"
            Set rdshippers = dbNorthwind.OpenRecordset(strsql)
            If rdshippers.BOF = True Then GoTo SkipTerminaison
              '
'                    For j = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                       
'                     If Trim(Mid(LettresDuMilieuBrutes, i, 1)) = Trim((MySettingAccords((j), 1))) Then
                          
                           rdshippers.MoveFirst
                          While rdshippers.EOF = False
                          ArrayMiddleLetters(k + 1, i) = Trim(rdshippers.Fields("valeur").Value)
                          
                          
                           ArrayMiddleLetters(kbis, 0) = ArrayMiddleLetters(kbis, 0) + 1
                           

                          
                           k = k + 1
                           NombreValeurLettreDuMilieu(i) = NombreValeurLettreDuMilieu(i) + 1 'repr�sente le nombre de combinaison pour chaque lettre
                           '1 si aucun son associ�
   
'                        End If
                    rdshippers.MoveNext
'                    Next j
                Wend
SkipTerminaison:
k = 0

    Next i
   
   

Dim NumberLinesFinalArray
 
 'on va d�terminer le nombre de lignes � composer
 
 
 
    '   NumberLinesFinalArray = 1
    Select Case TailleLettresDuMilieu
    
    Case 1 'une seule lettre au milieu
    
     NumberLinesFinalArray = NombreValeurLettreDuMilieu(1)
     
    Case 2 '
    
       NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2)
    
    Case 3
    
     NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3)
    
    Case 4
    
      NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4)
      
    
    
    Case 5
      NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5)
      
     
    
    Case 6
     NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6)
     
    
    Case 7
    
    NumberLinesFinalArray = NombreValeurLettreDuMilieu(1) * NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
     
    
    End Select 'NumberLinesFinalArray


'on dimensionne le tableau final qui contient, par ligne, toutes les combinaisons possibles

ReDim finalarray(NumberLinesFinalArray, TailleLettresDuMilieu) 'en l'occurrence, le tableau a autant de colonne que de lettres


Dim MyNombreOccurences(10), m, n, o

'peuplement du tableau des combinaisons possibles des lettres du milieu
' la r�gles est la suivante
'chaque lettre de la colonne est dupliqu�e par le produit du nombre de possibilit�es des colonnes suivantes
'pex, si premi�re colonne � 2 possibilit�s (pex : �, ch)
'deuxi�me colonne a 2 possibilit�s (pex : �, in)
'troisi�me colonne a 3 possibili�t� (pex : o, ou, hou)

'cela donne : premi�re colonne : 6 x � puis 6 x ch
'deuxi�me colonne : 6 x � puis 6 x in
'troisi�me colonne : o, ou, hou jusqu'� la compl�tion du tableau



Select Case TailleLettresDuMilieu 'cad le nombre de colonne


    Case 1
    
     
            For m = 1 To NumberLinesFinalArray
                For n = 1 To NombreValeurLettreDuMilieu(1)
                    finalarray(n, 1) = ArrayMiddleLetters(n, 1)
                Next
            Next
    
    
    
    Case 2
    
        'peuplement de la premi�re colonne case 2
        'chaque lettre de le premi�re colonne doit �tre r�p�t�e le nombre de fois qu'il y de combinaison dans la colonne suivante,
        'en l'occurence dans la deuxi�me colonne
 o = 1
             
         For m = 1 To NumberLinesFinalArray / NombreValeurLettreDuMilieu(2)
             
             For n = 1 To NombreValeurLettreDuMilieu(2)
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        
        Next
        'peuplement de la seconde colonne case 2
 o = 1
         While o <= NumberLinesFinalArray
                For n = 1 To NombreValeurLettreDuMilieu(2)
                    finalarray(o, 2) = ArrayMiddleLetters(n, 2)
                    o = o + 1
                Next
            Wend
    
    
    
    Case 3
    
    
    'peuplement de la premi�re colonne '  Case 3
    
    o = 1
    m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3)
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        m = m + 1
        Wend
    
    'peuplement de la deuxi�me colonne '  Case 3
    
      o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3)
        
                finalarray(o, 2) = ArrayMiddleLetters(m, 2)
                o = o + 1
                 
             Next
                If m < NombreValeurLettreDuMilieu(2) And NombreValeurLettreDuMilieu(2) <> 1 Then
             m = m + 1
        Else
            m = 1
        End If
        
        
        Wend
    
    
    'peuplement de la troisi�me colonne '  Case 3
    
     o = 1
          While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3)
        
                finalarray(o, 3) = ArrayMiddleLetters(n, 3)
                o = o + 1
               
             Next
       
        
        Wend
    
    Case 4
    
    'peuplement de la premi�re colonne   Case 4
    
    o = 1
    m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4)
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        m = m + 1
        Wend
    
    
    
      'peuplement de la deuxi�me colonne   Case 4
      
         o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4)
        
                finalarray(o, 2) = ArrayMiddleLetters(m, 2)
                o = o + 1
                 
            Next
        If m < NombreValeurLettreDuMilieu(2) And NombreValeurLettreDuMilieu(2) <> 1 Then
             m = m + 1
        Else
            m = 1
        End If
         
        
        Wend
      
      
        'peuplement de la troisi�me colonne   Case 4
        
        
         
        o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(4)
           
                   finalarray(o, 3) = ArrayMiddleLetters(m, 3)
                   o = o + 1
                   Next
               
                    If m < NombreValeurLettreDuMilieu(3) And NombreValeurLettreDuMilieu(3) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
           
           Wend
    
        
          'peuplement de la quatri�me colonne   Case 4
    
    
               
            o = 1
               m = 1
                While o <= NumberLinesFinalArray
                    
                    For n = 1 To NombreValeurLettreDuMilieu(4)
               
                       finalarray(o, 4) = ArrayMiddleLetters(n, 4)
                       o = o + 1
                      
                    Next
              
               
               Wend
    
    Case 5
    
     'peuplement de la premi�re colonne case 5
    
    
    o = 1
    m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5)
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        m = m + 1
        Wend
    
    
    
      'peuplement de la deuxi�me colonne case 5
      
         o = 1
        m = 1
         
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5)
              
               finalarray(o, 2) = ArrayMiddleLetters(m, 2)
                o = o + 1
                
                Next
          
          
                    If m < NombreValeurLettreDuMilieu(2) And NombreValeurLettreDuMilieu(2) <> 1 Then
                    m = m + 1
                    Else
                    m = 1
                    End If
          
          
        Wend
      
      
        'peuplement de la troisi�me colonne case 5
                
        o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5)
        
        
                
               finalarray(o, 3) = ArrayMiddleLetters(m, 3)
               o = o + 1
             Next
                
                
              If m < NombreValeurLettreDuMilieu(3) And NombreValeurLettreDuMilieu(3) <> 1 Then
                    m = m + 1
                    Else
                    m = 1
                    End If
          
       
        
        
          
        Wend
      
      
        
          'peuplement de la quatri�me colonne case 5
    
             o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(5)
        
        
                
               finalarray(o, 4) = ArrayMiddleLetters(m, 4)
                o = o + 1
                Next
                
              If m < NombreValeurLettreDuMilieu(4) And NombreValeurLettreDuMilieu(4) <> 1 Then
                    m = m + 1
                    Else
                    m = 1
              End If
          
          
        Wend
      
      

       'peuplement de la cinqui�me  colonne case 5
    
        o = 1
               m = 1
                While o <= NumberLinesFinalArray
                    
                    For n = 1 To NombreValeurLettreDuMilieu(5)
               
                       finalarray(o, 5) = ArrayMiddleLetters(n, 5)
                       o = o + 1
                      
                    Next
              
               
               Wend
    
    
    
    Case 6
    
        
     'peuplement de la premi�re colonne case 6
    
    
    o = 1
    m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6)
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        m = m + 1
        Wend
    
    
    
      'peuplement de la deuxi�me colonne case 6
      
         o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6)
        
                finalarray(o, 2) = ArrayMiddleLetters(m, 2)
                o = o + 1
              Next
             
        If m < NombreValeurLettreDuMilieu(2) And NombreValeurLettreDuMilieu(2) <> 1 Then
             m = m + 1
        Else
            m = 1
        End If
        
        
        Wend
      
      
        'peuplement de la troisi�me colonne case 6
        
        
         
        o = 1 'repr�sente la ligne du tableau
        m = 1 'repr�sente l'occurrence de r�p�tition de la s�quence
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6)
           
                   finalarray(o, 3) = ArrayMiddleLetters(m, 3)
                   o = o + 1
                   Next
               
                    If m < NombreValeurLettreDuMilieu(3) And NombreValeurLettreDuMilieu(3) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
            
           Wend
    
        
          'peuplement de la quatri�me colonne case 6
    
    
        o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6)
           
                   finalarray(o, 4) = ArrayMiddleLetters(m, 4)
                   o = o + 1
                   Next
                
                    If m < NombreValeurLettreDuMilieu(4) And NombreValeurLettreDuMilieu(4) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
          
           Wend
               
            
          'peuplement de la cinqui�me colonne case 6
    
    
        o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(6)
           
                   finalarray(o, 5) = ArrayMiddleLetters(m, 5)
                   o = o + 1
                   Next
                
                    If m < NombreValeurLettreDuMilieu(5) And NombreValeurLettreDuMilieu(5) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
          
           Wend
               
   
       'peuplement de la sixi�me  colonne case 6
    
        o = 1
               m = 1
                While o <= NumberLinesFinalArray
                    
                    For n = 1 To NombreValeurLettreDuMilieu(6)
               
                       finalarray(o, 6) = ArrayMiddleLetters(n, 6)
                       o = o + 1
                      
                    Next
              
               
               Wend
    
    
    
    
    Case 7
    
        'peuplement de la premi�re colonne case 7
        
           
    o = 1
    m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(2) * NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
        
        
                finalarray(o, 1) = ArrayMiddleLetters(m, 1)
                o = o + 1
             Next
        m = m + 1
        Wend
        
        'peuplement de la deuxi�me colonne case 7
    
             
         o = 1
        m = 1
         While o <= NumberLinesFinalArray
             
             For n = 1 To NombreValeurLettreDuMilieu(3) * NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
        
                finalarray(o, 2) = ArrayMiddleLetters(m, 2)
                o = o + 1
              Next
             
        If m < NombreValeurLettreDuMilieu(2) And NombreValeurLettreDuMilieu(2) <> 1 Then
             m = m + 1
        Else
            m = 1
        End If
        
        
        Wend
    
        'peuplement de la troisi�me colonne case 7
        
         o = 1 'repr�sente la ligne du tableau
        m = 1 'repr�sente l'occurrence de r�p�tition de la s�quence
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
           
                   finalarray(o, 3) = ArrayMiddleLetters(m, 3)
                   o = o + 1
                   Next
               
                    If m < NombreValeurLettreDuMilieu(3) And NombreValeurLettreDuMilieu(3) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
            
           Wend
    
        'peuplement de la quatri�me colonne case 7
        
        
        o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(4) * NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
           
                   finalarray(o, 4) = ArrayMiddleLetters(m, 4)
                   o = o + 1
                   Next
                
                    If m < NombreValeurLettreDuMilieu(4) And NombreValeurLettreDuMilieu(4) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
          
           Wend
            
        'peuplement de la cinqui�me colonne case 7
        
            o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(5) * NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
           
                   finalarray(o, 5) = ArrayMiddleLetters(m, 5)
                   o = o + 1
                   Next
                
                    If m < NombreValeurLettreDuMilieu(5) And NombreValeurLettreDuMilieu(5) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
          
           Wend
        
        'peuplement de sixi�me colonne case 7
        
         o = 1
           m = 1
            While o <= NumberLinesFinalArray
                
                For n = 1 To NombreValeurLettreDuMilieu(6) * NombreValeurLettreDuMilieu(7)
           
                   finalarray(o, 6) = ArrayMiddleLetters(m, 6)
                   o = o + 1
                   Next
                
                    If m < NombreValeurLettreDuMilieu(6) And NombreValeurLettreDuMilieu(6) <> 1 Then
                            m = m + 1
                            Else
                            m = 1
                    End If
          
           Wend
        
        'peuplement de la 7�me colonne case 7
            
            o = 1
               m = 1
                While o <= NumberLinesFinalArray
                    
                    For n = 1 To NombreValeurLettreDuMilieu(7)
               
                       finalarray(o, 7) = ArrayMiddleLetters(n, 7)
                       o = o + 1
                      
                    Next
              
               
               Wend
        
    
    Case 8
    
    
End Select 'TailleLettresDuMilieu
  
 ReDim LettresDuMilieu(NumberLinesFinalArray)

    For i = 1 To NumberLinesFinalArray
 'middleletters = "o" & "*" & "c"
       
            Select Case TailleLettresDuMilieu
        
                 Case 1
                 
                LettresDuMilieu(i) = finalarray(i, 1)
                MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                 
                 Case 2
                 
                LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                ' Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2)
                 MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2)
                 
                 Case 3
                 
                 LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                ' Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3)
                 MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3)
                 Case 4
                 
                   LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                ' Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4)
                 MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4)
                 
                 Case 5
                 
                  LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                 ' Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5)
                  MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5)
                 
                 Case 6
                 
                  LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                  'Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6)
                 MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6)
                
                 Case 7
            
                  LettresDuMilieu(i) = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6) & "*" & finalarray(i, 7)
                  MyInputBox.lettresMilieu.AddItem LettresDuMilieu(i)
                  'Debug.Print finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6) & "*" & finalarray(i, 7)
                  MyInputBox.lettresMilieu = finalarray(i, 1) & "*" & finalarray(i, 2) & "*" & finalarray(i, 3) & "*" & finalarray(i, 4) & "*" & finalarray(i, 5) & "*" & finalarray(i, 6) & "*" & finalarray(i, 7)
                  
        
        
          End Select 'TailleLettresDuMilieu
        
   

    Next i
 
 'MsgBox NumberLinesFinalArray * NombreTerminaisons * MyFirstLetterIsSound
 
 
 ReDim finalarray(0, 0) 'on essaye de lib�rer de l'espace m�moire
 
 
'If EndIsConjug = -1 Then GoTo skip8 'si la terminaison est une conjugaison, on ne doit pas chercher dans les terminaisons

 
'TRAITEMENT DE LA TERMINAISON
 
' MySettingAccords = GetAllSettings(appname:="fasttype", section:="terminaisons") '
 MyIndex = 1
 NombreTerminaisons = 1

            If myaccord = -1 Then 'hypoth�se o� la derni�re lettre �tait un accord (pluriel, f�minin etc).
                LastLetter = Left(TwoLastLetters, 1)
                            
            End If
            'iciici
            If EndIsConjug = -1 And MyPbkMsg = 2 Then
                LastLetter = "r"
            
            End If 'endisconjug = -1
'terminaisons(1) = Right(MyAb, 1)
terminaisons(1) = LastLetter

 
  'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
 
             strsql = "SELECT m�thode_ab.Valeur, m�thode_ab.Abr�viation FROM m�thode_ab WHERE (((m�thode_ab.Abr�viation)=""" & LastLetter & """) AND ((m�thode_ab.terminaison)=Yes));"
             Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

 If rdshippers.BOF = True Then GoTo SkipTerminaison_deux
    
                            rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                        terminaisons(MyIndex + 1) = Trim(rdshippers.Fields("valeur").Value)
                                        MyIndex = MyIndex + 1
                                        NombreTerminaisons = MyIndex '
                                                              
                                    rdshippers.MoveNext
                        
                                Wend


skip8:
 
 
SkipTerminaison_deux:
skip_lettres_du_milieu:
          
                                
                       Dim docNew As Document
                       ' Dim dbNorthwind As Database
                        'Dim rdShippers As Recordset
                        Dim SizeMot
                        Dim intRecords 'As Integer
                        Dim Filter, TailleD�butFile, TailleFinFile, AvantMiddle, Apr�sMiddle
 '                       Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
                        Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
                        Dim AbWithZ
                        If InStr(1, myab, "z") Then AbWithZ = -1
                        
                        l = 0
                        SizeMot = TailleMyAb + 2
                            'If InStr(LettresDuMilieuBrutes, "z") > 0 And EndIsConjug = "" Then

                            'MyTable = "les_mots_sans_les_conjugaisons_avec_z"
                            
                            'Else
                            
                                      If EndIsConjug = True Then
                            
                            ''''''''''''''''''''''''''''''''''''''
 's�lection de la table dans laquelle va se faire la recherche du mot brut
                                              Select Case AbWithZ '
                                        
                                         
                                                        Case -1
                                                           
                                                        MyTable = "infinitifs_avec_z"
                                                          
                                                        Case Else
                                                        
                                                       MyTable = "infinitifs_sans_z"
                                                     
                                                End Select
                                              
                                              
                                    Else
                                       
                                        Select Case EndIsAccord
                                        
                                            Case -1  'on termine par un accord
                                        ''''''''''''''''''''''''''
                                        
                                                Select Case Right(myab, 1)
        
                                                                Case get_accord("f�minin")
                                                                
                                                                     Select Case AbWithZ '
                                                                     
                                                                        Case -1
                                                                        
                                                                        MyTable = "mots_avec_f�minin_et_f�minin_pluriel_avec_z"
                                                                        MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_avec_z"
                                                                        
                                                                        Case Else
                                                                        
                                                                        MyTable = "mots_avec_f�minin_et_f�minin_pluriel_sans_z"
                                                                        MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_sans_z"
                                                                     
                                                                    End Select 'zLettresDuMilieuBrutes
        
        
                                                                Case get_accord("f�minin_pluriel")
                                                                
                                                                    Select Case AbWithZ '
                                                                     
                                                                        Case -1
                                                                        
                                                                        MyTable = "mots_avec_f�minin_et_f�minin_pluriel_avec_z"
                                                                        MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_avec_z"
                                                                        
                                                                        Case Else
                                                                        
                                                                        MyTable = "mots_avec_f�minin_et_f�minin_pluriel_sans_z"
                                                                        MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_sans_z"
                                                                     
                                                                    End Select 'zLettresDuMilieuBrutes
                    
        
                                                                Case get_accord("singulier_pluriel")
                                                                
                                                                    Select Case AbWithZ '
                                                                     
                                                                        Case -1
                                                                        
                                                                         MyTable = "mots_avec_pluriel_avec_z"
                                                                         MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_avec_z"
                                                                        
                                                                        Case Else
                                                                        
                                                                         MyTable = "mots_avec_pluriel_sans_z"
                                                                         MyTable1 = "tables_mots_finissant_par_abr�viation_des_accords_sans_z"
                                                                     
                                                                    
                                                                     
                                                                    End Select 'zLettresDuMilieuBrutes
        
        
        
                                                End Select 'Right(myab, 1)
                                        
                                        
                                        
                                        
                                       Case 0 'on termine pas par un accord
                            
                                             Select Case AbWithZ '

                                                        Case -1

                                                            MyTable = "seulement_les_mots_Z"
                                                            
                                                        Case 0
                                                        
                                                            MyTable = "tout_sans_les_z"

                                                        Case Else
                                                            
                                                            If Right(myab, 1) = "z" Then
                                                             MyTable = "seulement_les_mots_Z"
                                                             End If
                                                  
                                                End Select 'zLettresDuMilieuBrutes
                                                
                                                
                                  End Select 'EndIsAccord
                            
                            End If 'endisconjug is true
                            
                 
                            
              
   'MsgBox MyTable
                          
                            
                          
            
            
               For i = 1 To NombreTerminaisons
                    For j = 1 To MyFirstLetterIsSound
                        For z = 1 To NumberLinesFinalArray
                                        
                            Filter = SonFirstLetter(j) & "*" & LettresDuMilieu(z) & "*" & terminaisons(i)
                            
                            
                          strsql = "SELECT " & MyTable & ".forme," & MyTable & ".indice FROM " & MyTable & " WHERE (((" & MyTable & ".forme) Like """ & Filter & """));"
'                             StrSql = "SELECT " & MyTable & ".forme FROM " & MyTable & " WHERE (((" & MyTable & ".forme) Like """ & Filter & """));"
                        
                                Set rdshippers = dbNorthwind.OpenRecordset(strsql)
                              
                                                
                                If rdshippers.BOF = True Then GoTo fin
                        
                        
                                    rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                               
                                        'ins�rer ici un contr�le de doublon
                                        l = 1
                                        If MyInputBox.suggestions.ListCount = 0 Then
                                        If IsVerb(rdshippers.Fields(0).Value) = True Or EndIsConjug = -1 Then GoTo goto12345 'exclusion des verbes � l'infitinitf
                                        'ou des verbes conjug�s
                                            If contr�le_coh�rence_abr�viative(myab, rdshippers.Fields(0).Value) <> -1 Then
                                        
                                        
goto12345:
                                                    MyInputBox.suggestions.AddItem rdshippers.Fields(0).Value
                                                    MyInputBox.suggestions.List(l - 1, 4) = rdshippers.Fields(1).Value
                                                    l = l + 1
                                                    GoTo skip3243
                                            End If ' contr�le_coh�rence_abr�viative(MyAb, rdShippers.Fields(0).Value) <> -1
                                        Else
                                            For l = 1 To MyInputBox.suggestions.ListCount
                                                If MyInputBox.suggestions.List(l - 1) = rdshippers.Fields(0).Value Then
                                                GoTo skip984545
                                                
                                                End If

                                            
                                            Next l
                                                 If IsVerb(rdshippers.Fields(0).Value) = True Or EndIsConjug = -1 Then GoTo goto1234567 ''exclusion des verbes � l'infitinitf
                                        'ou des verbes conjug�s
                                                 If contr�le_coh�rence_abr�viative(myab, rdshippers.Fields(0).Value) <> -1 Then
                                                 
goto1234567:
                                                 
                                                    MyInputBox.suggestions.AddItem rdshippers.Fields(0).Value
                                                    MyInputBox.suggestions.List(l - 1, 4) = rdshippers.Fields(1).Value
                                                    l = l + 1
                                                  End If ' contr�le_coh�rence_abr�viative(MyAb, rdShippers.Fields(0).Value) <> -1
                                        End If
                                        
skip3243:
                                    
                                    
                       
skip984545:

                            
                            
                                       rdshippers.MoveNext
                        
                                Wend
                                                
fin:
                        
                        Next z
                    Next j
                  Next i
                             
''''''''''''''''''''''''''' boucle pour les singuliers ayant une terminaison pluriel
'''''''''''''''''''''''' contr�le des terminaisons !!!!!!!!!!!!!!!!!!!!!!!!
Dim h
For h = 0 To MyNumberVerbe - 1

Dim MyTerminaison
Dim NbrMotsZ, NbrMotsAvecZ
 MySettingAccords = GetAllSettings(appname:="fasttype", section:="terminaisons") '


        For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)


            If InStr(Len(myab) - 1, myab, MySettingAccords((i), 1)) = 0 Then 'la lettre repr�sentant une finale ne se trouve pas comme finale de l'abr�viation
               If Len(MyInputBox.suggestions.List(h)) - Len(MySettingAccords((i), 0)) + 1 > 0 Then '�viter une valeur n�gative qui arrive parfois
               If InStr(Len(MyInputBox.suggestions.List(h)) - Len(MySettingAccords((i), 0)) + 1, MyInputBox.suggestions.List(h), MySettingAccords((i), 0)) > 0 Then

                    For j = 1 To NombreTerminaisons - 1 'il faut voir si la terminaison en question n'est pas contenue pas dans l'une des terminaisons possibles
                    'par exemple : "re" ne doit pas entra�ner l'exclusion de "ure" ou de "oire"
                    'on r�cup�re � cette fin les terminaisons possibles dans l'array "terminaisons"




                        If InStr(1, terminaisons(j + 1), MySettingAccords((i), 0)) > 0 Then GoTo skip33  'il faut ajouter 1 car la premi�re terminaison,

                        'dans l'array "terminaison" correspond � la lettre elle-m�me (cela sert pour prendre en compte aussi la lettre elle-m�me..



                    Next j



                NbrMotsAvecZ = NbrMotsAvecZ + 1
                MyInputBox.rejet�s.AddItem MyInputBox.suggestions.List(h)
                MyInputBox.rejet�s.List(NbrMotsAvecZ - 1, 1) = "terminaison < " & MySettingAccords((i), 0) & " > dans le mot mais pas dans l'abr�viation"


                End If 'InStr(1, myinputbox.suggestions.List(h - 1), MySettingAccords((i), 1)) > 0
                End If 'Len(MyInputBox.suggestions.List(h)) - Len(MySettingAccords((i), 0)) + 1 > 0

            End If

        Next i

skip33:
Next h


''''''''''''''''''''' fin de contr�les de l'usage des terminaisons

    If EndIsAccord = -1 Then
    Dim Strsql1
   
    
    
                        For i = 1 To 1
                            For j = 1 To MyFirstLetterIsSound
                                For z = 1 To NumberLinesFinalArray
                                        
                        
                                    FilterTerminaisons = SonFirstLetter(j) & "*" & LettresDuMilieu(z) & "*" & Right(myab, 1)
                            
                                Strsql1 = "SELECT " & MyTable1 & ".forme," & MyTable1 & ".indice FROM " & MyTable1 & " WHERE (((" & MyTable1 & ".forme) Like """ & FilterTerminaisons & """));"
'
                        
                                Set rdshippers = dbNorthwind.OpenRecordset(Strsql1)
                                 
                                                
                                If rdshippers.BOF = True Then GoTo fin25
                        
                        
                                    rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                               
                                        'ins�rer ici un contr�le de doublon
                                        l = 1
                                        If MyInputBox.MotsAccords.ListCount = 0 Then
                                        MyInputBox.MotsAccords.AddItem rdshippers.Fields(0).Value
                                        MyInputBox.MotsAccords.List(l - 1, 4) = rdshippers.Fields(1).Value
                                    


                                        l = l + 1
                                        GoTo skip324345
                                        Else
                                            For l = 1 To MyInputBox.MotsAccords.ListCount
                                                If MyInputBox.MotsAccords.List(l - 1) = rdshippers.Fields(0).Value Then
                                                GoTo skip98454522
                                                
                                                End If

                                            
                                            Next l
                                        MyInputBox.MotsAccords.AddItem rdshippers.Fields(0).Value
                                        MyInputBox.MotsAccords.List(l - 1, 4) = rdshippers.Fields(1).Value

                                     

                                        l = l + 1
                                        End If
                                        
skip324345:
                                    
                                    
                       
skip98454522:

                            
                            
                                       rdshippers.MoveNext
                        
                                Wend
                                                
fin25:
                        
                        Next z
                    Next j
                  Next i
                             
                   
      End If 'endisaccord = - 1
                       
 
 
 
 
                If EndIsAccord = -1 Then
                
                        Dim Mysearch, TailleMySearch, MyTwoFirstLetters
                        Dim MyNumberSingulier
                        MyNumberSingulier = MyInputBox.suggestions.ListCount
                        
                            If MyNumberSingulier > 0 Then
                                    
                                    Dim mysingulier()
                                    ReDim mysingulier(MyInputBox.suggestions.ListCount - 1, 1)
                  
                   
                   'd�termination de la colonne de recherche dans l'une ou l'autre des tables
                   
                   
                                    For h = 1 To MyNumberSingulier
                                        mysingulier(h - 1, 0) = MyInputBox.suggestions.List(h - 1)
                                        mysingulier(h - 1, 1) = MyInputBox.suggestions.List(h - 1, 4)
                                    
                                    Next h

   
   
                                    Select Case Right(myab, 1)
                                    
                                        Dim MyLookupField
                                    
                                         Case get_accord("f�minin")
                                            
                                                MyLookupField = 1
                                            
                                            Case get_accord("f�minin_pluriel")
                                                
                                                MyLookupField = 2
                                    
                                    
                                    Case get_accord("singulier_pluriel")
                                    
                                                MyLookupField = 3
                                            
                                    End Select
                                      
                                          Dim MyOrthographe
                                          MyInputBox.suggestions.Clear
                                          Set rdshippers = dbNorthwind.OpenRecordset(MyTable)
                                                    
                                                        For h = 1 To MyNumberSingulier
                                                        
                                      
                                                                      With rdshippers
                                                                      .Index = "forme"
                                                                      .Seek "=", mysingulier(h - 1, 0)
                                                                      
                                                                     
                                                                
                                                                      
                                                                      End With
                                                                    
                                                                    If rdshippers.NoMatch = True Then
                                                                              
                                                                                
                                                                                 
                                                                                
                                                                        Else
                                                                                     
                                                                                MyOrthographe = rdshippers.Fields(MyLookupField).Value
                                                                                
                                                                                If IsNull(MyOrthographe) = True Then GoTo skip980489
                                                                                
                                                                                MyInputBox.suggestions.AddItem MyOrthographe
                                                                                MyInputBox.suggestions.List(h - 1, 4) = mysingulier(h - 1, 1)
                                                                              
                                                                              
                                                                    End If 'rdShippers.NoMatch = True
                                                                                                            
                                                                                                      
skip980489:
                                                        
                                                                                                      
                                                       Next h
                                                                                            
                                                         
                                        End If 'MyNumberSingulier >0


End If 'EndIsAccord = -1


If EndIsConjug = -1 Then  'la fin est une conjugaison
MyNumberVerbe = MyInputBox.suggestions.ListCount
 
If MyPbkMsg = 3 Then GoTo ChoixDuForm
'Dim sRacineVerbe, sTailleVerbe, MyNumberVerbe


    MyNumberVerbe = MyInputBox.suggestions.ListCount

    If MyNumberVerbe > 0 Then
        Dim MyNumber
         MyNumber = GetSetting("fasttype", section:="temps_combinaison", Key:=MyConjug)
        ReDim myconjugaisons(MyNumberVerbe)
        ReDim MyInfinitifs(MyNumberVerbe, 2)
        Dim MyConst

   

                For h = 1 To MyNumberVerbe
                    MyInfinitifs(h, 0) = MyInputBox.suggestions.List(h - 1) ' la forme
                     MyInfinitifs(h, 2) = MyInputBox.suggestions.List(h - 1, 4) 'l'indice
                     
                    MyConst = acc�der_verbe_dans_table(MyInputBox.suggestions.List(h - 1), MyNumber)
                    If IsNull(MyConst) = False Then
                   ' myconjugaisons(h) = MyConst
                   MyInfinitifs(h, 1) = MyConst 'la conjugaison indiqu�e par l'abr�viation
                   Else
                    MyNumberVerbe = MyNumberVerbe - 1
                    End If
                
                'toto

                       ' mysingulier(h, 2) = Mid(mysearch, 1, 1) & "xxx" & Mid(mysearch, 2, 1)


                    Next h

    End If

    MyInputBox.suggestions.Clear

        ' For i = 1 To NumberLinesFinalArray
        
            For h = 1 To MyNumberVerbe
                  MyInputBox.suggestions.AddItem MyInfinitifs(h, 1) 'la forme
                  MyInputBox.suggestions.List(h - 1, 3) = MyInfinitifs(h, 0) 'la conjugaison indiqu�e par l'abr�viation
                  MyInputBox.suggestions.List(h - 1, 4) = MyInfinitifs(h, 2) 'l'indice
                   
            Next h
            
        '

 End If 'EndIsConjug = -1 Then  'l

ChoixDuForm:


MyNumberVerbe = MyInputBox.suggestions.ListCount

    
 Dim MyExistingAb, myLenMyAb, MyLenMyWord, MyHit, MyAbSubstitution, MySelection
    
'MyNumberMaxMots = GetSetting("fasttype", section:="param�tres", Key:="NombreMaxMots")
MyNumberVerbe = MyInputBox.suggestions.ListCount





   If MyNumberVerbe <> 0 Then


    
        If InStr(LettresDuMilieuBrutes, "z") = 0 Then 'il n'y a pas la lettre z dans les lettres du milieu
            ReDim motsz(MyNumberVerbe), MotsAvecZ(MyNumberVerbe)
            
            NbrMotsZ = 0
       ' NbrMotsAvecZ = 0
                For h = 1 To MyNumberVerbe 'on passe au contr�le tous les mots contenus dans le champ myinputbox.suggestions
                'si on les trouve, on les passe dans un array qu'on videra ensuite dans le m�me champ,
                'apr�s l'avoir remis � z�ro (clear)

                  If chercher_dans_Lettres_Z(MyInputBox.suggestions.List(h - 1)) = False Then 'si le mot ne comporte pas z

                  'NbrMotsZ = NbrMotsZ + 1
                  'motsz(NbrMotsZ) = myinputbox.suggestions.List(h - 1)

                    Else 'c'est juste pour contr�ler que des mots ont �t� enlev�s (ils contiennent "z")
                    '�ventuellement, on pourrait songer � la r�utiliser

                  NbrMotsAvecZ = NbrMotsAvecZ + 1
                '  MotsAvecZ(NbrMotsAvecZ) = myinputbox.suggestions.List(h - 1)
                  MyInputBox.rejet�s.AddItem MyInputBox.suggestions.List(h - 1)
                  MyInputBox.rejet�s.List(NbrMotsAvecZ - 1, 1) = "contient le son < z > alors que pas dans abr�viation"
                  End If 'chercher_dans_Lettres_Z(myinputbox.suggestions.List(h - 1))


                Next h
               



        End If 'InStr(LettresDuMilieuBrutes, "z") = 0
    'Fin de la suppression du 16.08.2011

    Dim IsInMot, IsInMyAb, NbrMotsSansZ
    


If EndIsConjug = -1 Or MyInputBox.suggestions.ListCount <= 1 Then GoTo skip9999 'l'exercice d'exclusion des sons et des terminaisons ne peut pas fonctionner
'quand on cherche une conjugaison ou si le nombre des mots n'exc�de pas 5
MyD�butExclusion = Timer
k = 0
ReDim motsz(MyInputBox.suggestions.ListCount + MyInputBox.rejet�s.ListCount)
MyNumberVerbe = MyInputBox.suggestions.ListCount






MyFinExclusion = Timer
'on enl�ve de myinputbox.suggestions les mots qui sont dans myinputbox.rejet�s
NbrMotsAvecZ = 0
'ReDim motsz(myinputbox.suggestions.ListCount - myinputbox.rejet�s.ListCount)
 l = 0
        For h = 0 To MyInputBox.suggestions.ListCount - 1
        
            For i = 0 To MyInputBox.rejet�s.ListCount - 1
            
                If MyInputBox.suggestions.List(h) = MyInputBox.rejet�s.List(i) Then
                
                GoTo Skip55
                
                End If
            
            
            Next i
        
            NbrMotsAvecZ = NbrMotsAvecZ + 1
            MyInputBox.stock.AddItem MyInputBox.suggestions.List(h)
            MyInputBox.stock.List(l, 1) = MyInputBox.suggestions.List(h, 4)
            If IsNull(MyInputBox.suggestions.List(h, 3)) = False Then MyInputBox.stock.List(1, 2) = MyInputBox.suggestions.List(h, 3)
            l = l + 1
Skip55:
        
        Next h




''''''''''''''''''''''''''''''''''''''''''''''''

'on verse le contenu d'myinputbox.stock dans myinputbox.suggestion

MyInputBox.suggestions.Clear
For i = 0 To MyInputBox.stock.ListCount - 1

    MyInputBox.suggestions.AddItem MyInputBox.stock.List(i)
    MyInputBox.suggestions.List(i, 4) = MyInputBox.stock.List(i, 1)

Next i
MyInputBox.stock.Clear


skip9999:
            MyInputBox.Caption = filename & " " & myab
            MyNumberVerbe = MyInputBox.suggestions.ListCount
                
                
'                For h = 1 To MyNumberVerbe 'calcul les �conomies
' '               on pourra ajouter ici un calcul plus fin avec les accents
'                  calcul des �conomies
'                  MyInputBox.suggestions.List(h - 1, 1) = Len(MyInputBox.suggestions.List(h - 1)) - Len(Myab)
'
'
'                Next h
                 
'les lignes suivantes visent � s�lectionner le mot le plus approchant possible
'en substance, on cherche le mot qui a le m�me nombre de consonnes que le mot
If MyNumberVerbe = 0 Then GoTo aucun_mot
If MyNumberVerbe > 1 Then

    vider_table ("mots_sugg�r�s")
    
     Set rdshippers = dbNorthwind.OpenRecordset("mots_sugg�r�s")
    
'                 MyHit = 0
                 Dim taille_�gale
                 MyAbSubstitution = myab 'on passe la valeur
'                 myLenMyAb = nettoyer_voyelle(MyAbSubstitution)
                 ReDim MyLen(MyNumberVerbe, 1)
'                 If MyNumberVerbe = 1 Then
                 MySelection = MyInputBox.suggestions.List(0)
'
'                 GoTo after
'                 End If
                 
        For h = 1 To MyNumberVerbe
                    
                    
              
'              If nettoyer_voyelle((MyInputBox.suggestions.List(h - 1, 0))) = myLenMyAb Then
'                taille_�gale = -1
'              Else
'                taille_�gale = 0
'              End If
              
              
                 With rdshippers
                        .AddNew
                        !forme = MyInputBox.suggestions.List(h - 1, 0)
                        !indice = MyInputBox.suggestions.List(h - 1, 4)
                        !taille_�gale = myLenMyAb - nettoyer_voyelle((MyInputBox.suggestions.List(h - 1, 0)))
                        !consonnes_diff�rentes = fonction_comparer_mot_et_abr�viation(MyInputBox.suggestions.List(h - 1, 0), myab)
                        !�conomies = Len(MyInputBox.suggestions.List(h - 1)) - Len(myab)
                        !infinitif = MyInputBox.suggestions.List(h - 1, 3)
                        MySameConsonnes = Round(MySameConsonnes / Len(myab), 2)
                        !m�mes_consonnes = MySameConsonnes
                                             
                        .Update
                 
                 End With

            MySameConsonnes = 0
            Next h
                 
             
                 
                 
                 
                 
                 
                 'fin de l'essai
                    MyInputBox.suggestions.Clear
    
                 Set rdshippers = dbNorthwind.OpenRecordset("mots_par_pertinence")
                 MyNumberVerbe = rdshippers.RecordCount
                 

                                        i = 0

                                If rdshippers.BOF = True Then GoTo fin
                        
                        
                                    rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                               

                                        MyInputBox.suggestions.AddItem rdshippers.Fields(0).Value 'la forme
                                        MyInputBox.suggestions.List(i, 1) = rdshippers.Fields(5).Value 'les �conomies
                                        
                                        If IsNull(rdshippers.Fields("infinitif").Value) = False Then MyInputBox.suggestions.List(i, 2) = rdshippers.Fields("infinitif").Value
                                        
                                        If EndIsConjug = -1 Then MyInputBox.suggestions.List(i, 3) = rdshippers.Fields(5).Value                                          'l'infinitif
                                        i = i + 1
                          
                            
                                       rdshippers.MoveNext
                        
                                Wend





after:
                 
 End If '  MyInputBox.suggestions.List(m, 1) = MyValeurConjugu�e
     
                 
                 
                      '  MyInputBox.zone_mot.AddItem myinputbox.suggestions.List(1 - 1)
                
'                If MySelection = "" Then
                MyInputBox.suggestions = MyInputBox.suggestions.List(0, 0)
                If MyNumberVerbe = 1 Then MyInputBox.suggestions.List(0, 1) = Len(MyInputBox.suggestions.List(0, 0)) - Len(myab)
                MyInputBox.zone_mot = MyInputBox.suggestions.List(0, 0)
'                 Else
'                    MyInputBox.suggestions = MySelection
'                  MyInputBox.zone_mot = MySelection
'                End If
                
                MyInputBox.compteur = MyInputBox.suggestions.ListCount
                
'                MyInputBox.nbre_fichiers_consult�s = MyInputBox.fichiers_consult�s.ListCount
                    
                MyInputBox.zone_abr�viation = myab
                Dim MyAbExtractiond�but, MyAbExtractionFin
                 '   Else 'cad mynumberverbe > 15
                    
                    
                'For h = 1 To MyNumberVerbe
                'MyInputBox.zone_mot.AddItem MyInputBox.suggestions.List(h - 1)
                 'MyInputBox.suggestions.AddItem MyInputBox.suggestions.List(h - 1)
                 
                
                'Next
          
      
               
     
          
          
          
            If MyInputBox.suggestions.ListCount > 15 Then
                MyInputBox.texte = MyNumberVerbe & " mots possibles pour < " & myab & " >. Mieux vaut changer d'abr�viation !"
                    Else
                MyInputBox.texte.Caption = MyInputBox.suggestions.ListCount & " mots correspondants � < " & myab & " > dans le dictionnaire"
            End If
            
            MyInputBox.suggestions.TabIndex = 0
            MyInputBox.bouton_annuler.TabIndex = 1
            
            '
            
            For i = 1 To NombreTerminaisons
            MyInputBox.terminaisons_examin�es.AddItem terminaisons(i)
            Next i
            
            For i = 1 To MyFirstLetterIsSound
            MyInputBox.sons_examin�s.AddItem SonFirstLetter(i)
            Next i
            
            GoTo skip99889
            If MyFirstLetterIsSound = 0 Then
                        MyInputBox.sons_examin�s.AddItem FirstLetter
                            Else
                        MyInputBox.sons_examin�s.AddItem FirstLetter
                        
                        For j = 1 To MyFirstLetterIsSound
                        MyInputBox.sons_examin�s.AddItem SonFirstLetter(j)
                        Next j
             End If 'MyFirstLetterIsSound = 0
skip99889:
            
             
             Dim NbrSons, NbrTerminaisons, NbrMilieux, MyTotal
             If MyInputBox.sons_examin�s.ListCount = 0 Then
             NbrSons = 1
             Else
             NbrSons = MyInputBox.sons_examin�s.ListCount
             End If
             
             NbrTerminaisons = MyInputBox.terminaisons_examin�es.ListCount - 1
             
             NbrMilieux = MyInputBox.lettresMilieu.ListCount
  
             
                              
            
            MyInputBox.nombre_combinaisons = "cette abr�viation a n�cessit� l'examen de " & NbrSons * NbrTerminaisons * NbrMilieux & " combinaisons"
'            MyInputBox.fichiers_examin�s = MyInputBox.fichiers_consult�s.ListCount & " fichiers ont �t� lus pour cette recherche"
            MyHeureFin = Timer
            MyInputBox.temps_recherche = MyHeureFin - MyHeureD�but
            
               If MyInputBox.rejet�s.ListCount = 0 Then
                MyInputBox.listes_d�roulantes.Pages(0).Caption = "0 mot rejet�"
                

                
            Else
                 MyInputBox.listes_d�roulantes.Pages(2).Caption = MyInputBox.rejet�s.ListCount & " mot rejet�s"
                 
            
            End If ' MyInputBox.rejet�s.ListCount = 0
                          
         '   log_recherche Date, Time, MyAb, NbrSons * NbrTerminaisons * NbrMilieux, MyHeureFin - MyHeureD�but, MyInputBox.suggestions.ListCount, MyInputBox.rejet�s.ListCount, (MyHeureFin - MyHeureD�but) / (NbrSons * NbrTerminaisons * NbrMilieux), NbrSons, NbrMilieux, NbrTerminaisons, MyFinExclusion - MyD�butExclusion
            
            'log_recherche Myword, myab, combinatoire, temps, nombreMotsPropos�s, NombreMotsExclus
            
            
              If MyInputBox.MotsAccords.ListCount > 0 Then
       
      
       
                 For i = 0 To MyInputBox.MotsAccords.ListCount - 1
                MyInputBox.MotsAccords.List(i, 1) = Len(MyInputBox.MotsAccords.List(i, 0)) - Len(myab)
                Next
                
            End If
            
            
            If UsageRechercheMot = "chercher_�_nouveau" Then
            
                
               ' MyAbExtractiond�but = Timer
                chercher_utilisation_abr�viation (myab)
                maj_abr�viations_utilis��es_dans_myinputbox (myab)
               ' MyAbExtractionFin = Timer
               ' MyInputBox.tempsAb = MyAbExtractionFin - MyAbExtractiond�but

            
            Else
            
                       maj_abr�viations_utilis��es_dans_myinputbox (myab)
                  peupler_ab_similaires myab, "myinputbox"
                       
                     MyInputBox.Show
            End If
            
         
   
                     
                     
   Else 'aucun mot trouv�
aucun_mot:
          
    MyInputBox.texte.Caption = "aucune abr�viation trouv�e !"
                MyInputBox.zone_abr�viation = myab
                
                    
            If UsageRechercheMot <> "chercher_�_nouveau" Then
               peupler_ab_similaires myab, "myinputbox"
                MyInputBox.Show
            End If
            
           
   End If '<= MyNumberMaxMots
Exit Sub
error:
        If err = 62 Then
        Set fso = Nothing
        MyInputBox.Caption = filename & " " & myab
        MyInputBox.compteur = MyInputBox.suggestions.ListCount
        peupler_ab_similaires myab, "myinputbox"
        MyInputBox.Show
        Else
        MsgBox error & " " & err
        
        Stop
        
        End If


End Sub



Public Sub contr�le_accord(MyNewWord, j, myab, sNombre)
 Dim m, MySuggestionA, myautocorrectsA, n, MyAutoCorrects, MyNumber
 
                           
      '
                                  For m = 1 To sNombre
                                 
                                    'on recherche les corrections
                                    'puis on regarde si elles ont fait l'objet d'une abr�viation
                                    'et on les ins�re dans la zone_suggestion_orthographe du form myinputbox
                                    'on indique s'il y a une abr�viation ou non
                                    
                                     MySuggestionA = Application.GetSpellingSuggestions(MyNewWord).Item(m)
                                     
                                     
                                     
                                     If InStr(1, MySuggestionA, " ") = 0 Then
                                     accords.suggestions.AddItem MySuggestionA
                                    ' accords.suggestions.List(m - 1, 1) = "pas d'abr�viation"
                                     End If
                                         myautocorrectsA = AutoCorrect.Entries.Count
                                         
                                          'd�sactiv� car ralentit l'ex�cution
                                          '  For n = 1 To myautocorrectsA
                                
                                           ' If AutoCorrect.Entries(n).Value = MySuggestionA Then
                                            'accords.suggestions.List(m - 1, 1) = AutoCorrect.Entries(j).Name
                                      
                                            'End If 'AutoCorrect.Entries(j).Value = MySuggestion
                                
                                
                                  Next 'm
                                        
                               ' Next 'n
                
                         
            
       
                        
                        accords.origine.Caption = "S�lection effectu�e � partir des abr�viations existantes"
                        accords.suggestions = Application.GetSpellingSuggestions(MyNewWord).Item(1)
                        accords.compteur = accords.suggestions.ListCount
                        
                        accords.BackColor = &HFF0000
                        accords.Show
                    
                       
                        
End Sub

Public Sub essai_extraction_dico()

Dim MyMot, MyNewWord, MyOrthographe, i




MyNewWord = "oment"
MyOrthographe = Application.GetSpellingSuggestions(MyNewWord).Count
For i = 1 To MyOrthographe
MsgBox Application.GetSpellingSuggestions(MyNewWord, , , , 1).Item(i)

Next i





End Sub

Public Sub essa_get_all_settings()
Dim mysettings, intsettings
mysettings = GetAllSettings(appname:="fasttype", section:="param�tres")
    For intsettings = LBound(mysettings, 1) To UBound(mysettings, 1)
        MsgBox mysettings(intsettings, 0)
    Next intsettings

End Sub
Public Sub cr�er_voyelles()
 SaveSetting appname:="fasttype", section:="voyelles", Key:="a", setting:="a"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="e", setting:="e"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="i", setting:="i"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="o", setting:="o"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="u", setting:="u"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="y", setting:="y"
 
 SaveSetting appname:="fasttype", section:="voyelles", Key:="�", setting:="�"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="�", setting:="�"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="�", setting:="�"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="�", setting:="�"
 SaveSetting appname:="fasttype", section:="voyelles", Key:="�", setting:="�"
 
End Sub
Public Sub cr�er_sons()

 SaveSetting appname:="fasttype", section:="sons", Key:="ain", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="an", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="au", setting:="o"
 SaveSetting appname:="fasttype", section:="sons", Key:="en", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="in", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="oin", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="on", setting:="h"
 SaveSetting appname:="fasttype", section:="sons", Key:="ont", setting:="h"
 SaveSetting appname:="fasttype", section:="sons", Key:="ou", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="out", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="oux", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="ch", setting:="�"
 SaveSetting appname:="fasttype", section:="sons", Key:="eau", setting:="o"
 
End Sub
Public Sub cr�er_terminaisons()
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ment", setting:="m"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="er", setting:="r"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="tion", setting:="n"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ssion", setting:="n"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="in", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ain", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="aint", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ein", setting:="�"
    SaveSetting appname:="fasttype", section:="terminaisons", Key:="de", setting:="d"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="oin", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="on", setting:="h"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ont", setting:="h"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ou", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="out", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="oux", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="che", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="eau", setting:="o"
SaveSetting appname:="fasttype", section:="terminaisons", Key:="age", setting:="g"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ble", setting:="l"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ce", setting:="�"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="gie", setting:="j"
  SaveSetting appname:="fasttype", section:="terminaisons", Key:="ie", setting:="i"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="que", setting:="q"
  SaveSetting appname:="fasttype", section:="terminaisons", Key:="ir", setting:="r"
 SaveSetting appname:="fasttype", section:="terminaisons", Key:="ive", setting:="v"
  SaveSetting appname:="fasttype", section:="terminaisons", Key:="mand", setting:="m"
  SaveSetting appname:="fasttype", section:="terminaisons", Key:="t�", setting:="t"
   SaveSetting appname:="fasttype", section:="terminaisons", Key:="ure", setting:="u"
  
  
  
  
 
End Sub

Public Sub cr�er_conjugaisons()
SaveSetting appname:="fasttype", section:="conjugaisons", Key:="rsf", setting:="eront"
SaveSetting appname:="fasttype", section:="conjugaisons", Key:="rnf", setting:="erons"
End Sub

Public Sub test()
Dim myEssai
myEssai = GetSetting("fasttype", section:="conjugaisons", Key:="rsf")
MsgBox myEssai
myEssai = GetSetting("fasttype", section:="conjugaisons", Key:="too")
If myEssai = "" Then MsgBox "pas de r�sultat"
End Sub



Public Function get_personne(sPersonne)
get_personne = GetSetting(appname:="fasttype", section:="personnes", Key:=sPersonne)
End Function

Public Function get_temps(sTemps)
get_temps = GetSetting(appname:="fasttype", section:="temps", Key:=sTemps)
End Function






Function check_existence_valeur_pour_abr�viation(myab)


On Error GoTo myerreur




MyIndexAutocorrect = AutoCorrect.Entries(myab).Index
 
check_existence_valeur_pour_abr�viation = True




myerreur:
If err = 5941 Then

check_existence_valeur_pour_abr�viation = False

End If
End Function



Function check_existence_nom_pour_abr�viation(MyMot)

On Error GoTo erreur


Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
Dim FirstLetter, trouv�, MyAutoCorrects, j, MyBegin, MyName
'ReDim MyMot(1)
Dim LastLetter
trouv� = 1
MyCompteurMots = 0
MyAutoCorrects = AutoCorrect.Entries.Count


  
                           
   For i = 1 To MyAutoCorrects
   
      MyValue = AutoCorrect.Entries(i).Value
       
                    If MyValue <> MyMot Then
                        'MyIndex = MyIndex + 1
                        GoTo LoopAhead
                    Else
                        check_existence_nom_pour_abr�viation = AutoCorrect.Entries(i).Name
                        'MsgBox check_existence_nom_pour_abr�viation
                        Exit Function
                    End If ' Len(MyValue) <> Len(MyMot)
                    
             

   
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
LoopAhead:
    Next i 'While MyEnd = 0
    
check_existence_nom_pour_abr�viation = False
        
Exit Function

erreur:
        If err = 62 Then
       ' accords.lettresMilieu.AddItem middleletters
        'accords.fichiers_consult�s.AddItem FileName & " " & MyCompteurMots
        'accords.compteur = accords.compteur + MyCompteurMots
        
        'Exit Function
        
      
        End If




End Function









Sub extractions_complexes(MyLetter, MyFirstLetter, MyTerminaison)
'cette fonction extrait les mots d�butants par une lettre de l'alphabet par les terminaisons
'myletter est le nom donn� au fichier
'myfirstletter est la premi�re lettre du mot
'ma terminaison est la terminaison de correspondance (ex : ment pour m)




Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        If fso.FileExists(get_hd & ":\mots\extractions\" & MyLetter & ".txt") Then
        
            AjoutMot = -1
           Set MyDestinationFile = fso.OpenTextFile(get_hd & ":\mots\extractions\" & MyLetter & ".txt", ForAppending, True)
           NombreMots = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
           
           
            Else
        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\extractions\" & MyLetter & ".txt")
        End If


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like MyFirstLetter & "*" & MyTerminaison & "*" Then
                MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
                   
Set fso = Nothing
'If MyPreviousNumber = "" Then MyPreviousNumber = 0
'SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
'MsgBox "end " & MyLetter & " " & MyTerminaison

End Sub
Sub lancer_extractions_complexes()
Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3


MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") 'il y a l� toutes les lettres de l'alphabet
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                    'je d�sactive cette partie pour l'essai des suivantes
                'MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="terminaisons") '
        
                 '       For n = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                          
                  '        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                   '       myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                    '     myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                     '  extractions_complexes myval2 & "XXX" & myval4, myval2, myval3
                     
                    
                      ' Next n
           
           MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="conjugaisons_deuxi�me") '
        
                        For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                          myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                          myval3 = (MySettingAccords2((n), 1)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 0)) 'lettre abr�geante (ex : m (pour ment)
                          
                       extractions_complexes_conjugaisons myval2 & "XXX" & myval4, myval2, myval3
                     
                    
                       Next n
            
            
             MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="conjugaisons_premier") '
        
                        For p = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                          
                          myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                          myval3 = (MySettingAccords2((o), 1)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((o), 0)) 'lettre abr�geante (ex : m (pour ment)
                          
                       extractions_complexes_conjugaisons myval2 & "XXX" & myval4, myval2, myval3
                     
                    
                       Next p
            
            
                
        Next m




'lancer = extractions_complexes("aXXXXm", "a", "ment")
MsgBox "c'est fini"

End Sub
Sub extractions_complexes_conjugaisons(MyLetter, MyFirstLetter, MyTerminaison)
''cette fonction extrait les mots d�butants par une lettre de l'alphabet par les conjugaisons
''myletter est le nom donn� au fichier
''myfirstletter est la premi�re lettre du mot
''ma terminaison est la terminaison de correspondance (ex : ment pour m)
'
'
'
'
'Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, trouv�
'
'filename = get_hd & ":\mots\tous_les_mots.txt"
''Filename = "f:\essai.txt"
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'Set ts = fso.OpenTextFile(filename, ForReading)
'
'
'        If fso.FileExists(get_hd & ":\mots\extractions\" & myletter & ".txt") Then
'
'            AjoutMot = -1
'           Set MyDestinationFile = fso.OpenTextFile(get_hd & ":\mots\extractions\" & myletter & ".txt", ForAppending, True)
'          ' NombreMots = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
'
'
'            Else
'
'
'
'        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\extractions\" & myletter & ".txt")
'        End If
'
'
'       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
'       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
's = ts.readline 'lit la premi�re ligne
'
'  trouv� = 0
'
'    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
'    MyIndex = MyIndex + 1
'                's = Replace(s, " ", "")
'                'j = Len(s)
'                'MyBegin = InStr(1, s, " ==== ")
'
'                MyValue = s
'
'                If MyValue Like MyFirstLetter & "*" & MyTerminaison Then
'                MyCompteurMots = MyCompteurMots + 1
'                'MyInputBox.zone_mot.AddItem s
'                'Debug.Print s
'                MyDestinationFile.WriteLine s
'
'                End If
'
'
'
'
'               s = ts.readline
'               'Debug.Print '
'               ' MyEnd = InStr(1, s, "ZYTHUMS")
'    Loop 'While MyEnd = 0
'skip10:
''Set fso = Nothing
''If MyPreviousNumber = "" Then MyPreviousNumber = 0
''SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
''MsgBox "end " & MyLetter & " " & MyTerminaison
'If MyFirstLetter = "z" Then
''MsgBox "on est arriv� � z"
'
'End If

End Sub
Sub mots_renvers�s()
'nous nous sommes arr�t�s � esuellud�m
'il faudra reprendre en changeant le mode d'acc�s au fichier.
Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, trouv�, TailleMot, result
Dim lettres
ReDim lettres(30)
filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

    
          
        
     Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\mots_renvers�s.txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  trouv� = 0
                           
    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                TailleMot = Len(s)
                
                        For i = 1 To TailleMot
                        lettres(i) = Mid(s, i, 1)
                        'Debug.Print lettres(i)
                        's = Left(s, TailleMot - i)
                        result = lettres(i) & result
                        'Debug.Print result
                        Next i
                
                  
                'If MyValue Like MyFirstLetter & "*" & myterminaison Then
                MyCompteurMots = MyCompteurMots + 1
                'MyInputBox.zone_mot.AddItem s
                'Debug.Print s
                MyDestinationFile.WriteLine result
                
                
                
               
                

               s = ts.readline
               result = ""
               'Debug.Print '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0






End Sub
Sub ajouter_les_sons_comme_d�but()

'cette fonction cr�e des fichiers des mots commen�ant par les sons r�pertori�s.
'pex le son ch s'�crit �. Il aura donc un fichier ch.txt comportant tous les mots commen�ant par ce son "che"



Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, trouv�, MyLetter
Dim MySettingAccords

MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") '
        
    For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
        
                
               MyLetter = Trim((MySettingAccords((i), 0)))
                
    '            GoTo skip0 'si on trouve la lettre, comme c'est un index unique, on sors de la boucle. Sinon, on va jusqu'au bout.
     '           End If

     
'en fait, il faudrait faire cette op�ration pour tous les fichiers qui commencent par ces lettres





filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


          Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\extractions\" & MyLetter & ".txt")
           
          
            


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  trouv� = 0
    MyIndex = 0
    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
             
                  
                If s Like MyLetter & "*" Then
                'MyCompteurMots = MyCompteurMots + 1
                'MyInputBox.zone_mot.AddItem s
                'Debug.Print s
                MyDestinationFile.WriteLine s
                
                End If
                
               
                

               s = ts.readline
               'Debug.Print '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
skip10:
'Set fso = Nothing
'If MyPreviousNumber = "" Then MyPreviousNumber = 0
'SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
'MsgBox "end " & MyLetter & " " & MyTerminaison


 Next i




End Sub

Sub lancer_d�but_lettre_fin_lettre()

'Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3
'
'
'MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") 'il y a l� toutes les lettres de l'alphabet
'
'            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
'
'                    'je d�sactive cette partie pour l'essai des suivantes
'                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
'
'                  For n = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
'
'
'                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
'                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
'                        ' myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
'
'                      cr�ation_fichier_d�but_lettre_fin_lettre myval2, myval3
'
'
'                       Next n
'
'
'
'
'        Next m
'
'
'
'
'
'End Sub
'
'Sub lancer_d�but_son_fin_lettre()
'
'Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3
'
'
'MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") '
'
'            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
'           ' For m = 20 To UBound(MySettingAccords, 1)
'                    'je d�sactive cette partie pour l'essai des suivantes
'                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
'
'                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
'
'
'                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
'                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
'                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
'
'                      cr�ation_d�but_son_fin_lettre myval2, myval3
'
'                Next n
'
'
'
'        Next m
'
'

End Sub

Sub lancer_d�but_lettre_fin_conjugaison()
Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3


MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") 'il y a l� toutes les lettres de l'alphabet
        
           For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
            'For m = 20 To UBound(MySettingAccords, 1)
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="conjugaisons_premier") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_lettre_fin_conjugaison myval2, myval3, myval4
                     
                    
                       Next n
           
                 For o = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                   MySettingAccords3 = GetAllSettings(appname:="fasttype", section:="conjugaisons_deuxi�me") '
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords3((o), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords3((o), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_lettre_fin_conjugaison myval2, myval3, myval4
                     
                    
                       Next o
            
                
        Next m



End Sub

Sub d�but_son_fin_conjugaison()

Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3


MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") '
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
         MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="conjugaisons_premier") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                         
                      cr�ation_d�but_son_fin_conjugaison myval2, myval3, myval4
                     
                    
                       Next n
           
                 For o = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                   MySettingAccords3 = GetAllSettings(appname:="fasttype", section:="conjugaisons_deuxi�me") '
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords3((o), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords3((o), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_son_fin_conjugaison myval2, myval3, myval4
                     
                    
                       Next o
                      
            
                
        Next m




End Sub
Function cr�ation_fichier_d�but_lettre_fin_lettre(LettreD�but, LettreFin)






Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettre_fin_lettre\" & LettreD�but & "XXX" & LettreFin & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" & LettreFin Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
                   
'Set fso = Nothing
'If MyPreviousNumber = "" Then MyPreviousNumber = 0
'SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
'MsgBox "end " & MyLetter & " " & MyTerminaison
End Function
Function cr�ation_d�but_lettre_fin_conjugaison(LettreD�but, AbConjug, Conjug)






        
        
     
        'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettre_fin_conjugaison\" & LettreD�but & "XXX" & Conjug & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" & Conjug Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0

End Function
Function cr�ation_d�but_son_fin_lettre(D�butSon, LettreFin)




        
        
     
        'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_son_fin_lettre\" & D�butSon & "XXX" & LettreFin & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like D�butSon & "*" & LettreFin Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0


End Function

Public Function cr�ation_d�but_son_fin_conjugaison(D�butSon, AbConjug, Conjug)

 



Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

       
        '& LettreD�but
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_son_fin_conjugaison\" & D�butSon & "XXX" & Conjug & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like D�butSon & "*" & Conjug Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0




End Function

Sub lancer_d�but_lettre_fin_terminaison()
Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3
MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
           ' For m = 20 To UBound(MySettingAccords, 1)
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="terminaisons") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_lettre_fin_terminaison myval2, myval3
                     
                Next n
                      
            
                
        Next m




End Sub


Function cr�ation_d�but_lettre_fin_terminaison(LettreD�but, SonTerminaison)



Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

       
        '& LettreD�but
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettre_fin_terminaison\" & LettreD�but & "XXX" & SonTerminaison & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" & SonTerminaison Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0




End Function









Sub signature()
Attribute signature.VB_Description = "Macro enregistr�e le 30/03/2010 par Emmanuel"
Attribute signature.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.signature"
'
' signature Macro
' Macro enregistr�e le 30/03/2010 par Emmanuel
'
    Selection.InlineShapes.AddPicture filename:=get_hd & ":\sign\eb3.JPG", _
        LinkToFile:=False, SaveWithDocument:=True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.InlineShapes(1).PictureFormat.ColorType = _
        msoPictureBlackAndWhite
End Sub



Public Sub lancer_chercher_mot_finissant_par()

Dim MyTerminaison
MyTerminaison = InputBox("indiquer la terminaison", "Recherche")
 ' chercher_mot_finissant_par MyTerminaison, "montrer"
End Sub

Public Sub chercher_son_contenu_dans_mots(MySon, Son_ou_terminaison)

'my action = "compter" ou "montrer"
Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
Dim FirstLetter, MyMot, trouv�




 Dim docNew As Document
                        'Dim dbNorthwind As DAO.Database
                        Dim rdshippers As Recordset
                        Dim SizeMot
                        Dim intRecords 'As Integer
                        Dim Filter, TailleD�butFile, TailleFinFile, AvantMiddle, Apr�sMiddle, strsql
                        Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")



Select Case Son_ou_terminaison

        Case "terminaison"

        strsql = "SELECT table_m�re.forme FROM table_m�re WHERE (((table_m�re.forme) Like ""*" & MySon & """));"
      MyMot = "la terminaison : "
      
       
        Case "milieu"
        strsql = "SELECT table_m�re.forme FROM table_m�re WHERE ((table_m�re.forme Like ""?*" & MySon & "*?""));"
         MyMot = "le son de milieu : "
               
               
          
    
          
          
        Case "d�but"
         
         
         strsql = "SELECT table_m�re.forme FROM table_m�re WHERE (((table_m�re.forme) Like """ & MySon & "*"" ));"
          MyMot = "le son de d�but : "
        
        End Select
       
                Set rdshippers = dbNorthwind.OpenRecordset(strsql)
                
       
       
       

'
'
'
'
'
                              If rdshippers.BOF = True Then GoTo fin
'
'
                                    rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                    accords.suggestions.AddItem rdshippers.Fields(0).Value

'
'                                        MyInputBox.suggestions.AddItem
'
'                                        End If
'
'skip3243:
'
'
'
'skip984545:
'
'
'
                                       rdshippers.MoveNext
'
                                Wend
'
fin:
'
'
'
'                If MyValue Like "*" & MySon & "*" & "?" Or MyValue Like MySon & "*" & "*" Then
'
'
'
'             '   MyCompteurMots = MyCompteurMots + 1
'               ' trouv� = trouv� + 1
'               ' ReDim MyMot(trouv�)
'
'                'MyMot(trouv� - 1) = s
'                MyInputBox.suggestions.AddItem s
'
'
'
'
'
'                End If
                
   
             '  s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
  '  Loop 'While MyEnd = 0
 
' Select Case MyAction
'
'    Case "compter"
'
      
'
'
'    Case "montrer"
'
'        accords.fichiers_consult�s.AddItem filename & " " & MyCompteurMots
        accords.compteur = accords.suggestions.ListCount
        accords.bouton_ne_pas_abr�ger.Visible = False
        accords.bouton_pas_trouv�.Visible = False
        accords.bouton_modifier_abr�viation.Visible = False
        accords.�tiquette.Visible = False
        accords.origine = accords.compteur & " mots comportant " & MyMot & " " & MySon
        accords.Caption = "extraction du dictionnaire de tous les mots de la langue fran�aise"
        accords.bouton_annuler.Visible = False
'        accords.nbre_fichiers_consult�s = accords.fichiers_consult�s.ListCount
        accords.Show
'
'End Select
'
'
'        Exit Sub
'erreur:
'        If err = 62 Then
'
'        Select Case MyAction
'
'    Case "compter"
'
'        modif_son_terminaisons.bouton_voir_terminaisons.Caption = MyCompteurMots & " mots concern�s"
'
'
'    Case "montrer"
'
'        MyInputBox.fichiers_consult�s.AddItem filename & " " & MyCompteurMots
'        MyInputBox.compteur = MyInputBox.suggestions.ListCount
'        MyInputBox.nbre_fichiers_consult�s = MyInputBox.fichiers_consult�s.ListCount

'
'End Select
'
'
'
'        End If

End Sub

Public Sub lancer_chercher_son_contenu_dans_mots()
Dim MySon
MySon = InputBox("son")
chercher_son_contenu_dans_mots MySon, "montrer"
End Sub



Public Function chercher_dans_Lettres_Z(MyMot)
'cette fonction cherche si un mot comporte le son "z"


'Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
'Dim FirstLetter, trouv�, lettre1, lettre2, lettre3
'ReDim MyMot(1)
'Dim LastLetter, h, MyExistingAb
'trouv� = 1
'MyCompteurMots = 0
'FirstLetter = Left(mymot, 3)
'lettre1 = Left(FirstLetter, 1)
'lettre2 = Mid(FirstLetter, 2, 1)
'lettre3 = Right(FirstLetter, 1)




'filename = get_hd & ":\mots\lettres_Z\" & lettre1 & lettre2 & lettre3 & ".txt"


'Filename = "f:\essai.txt"
'Set fso = CreateObject("Scripting.FileSystemObject")

'Set ts = fso.OpenTextFile(filename, ForReading)
'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\lettres_Z\" & Letter1 & "xxx" & Letter2 & ".txt")



       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
's = ts.readline 'lit la premi�re ligne

  
                           
 '   Do While MyIndex < 336530 - 1 ' 'And MyTrouv�Premier <> 0
  '  MyIndex = MyIndex + 1
    
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
   '             MyValue = s
    '
     '           If MyValue = mymot Then
                'iciiciici
      '         MyInputBox.fichiers_consult�s.AddItem filename
        '       Exit Function
         '       End If
                
   
       '        s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    'Loop 'While MyEnd = 0

'erreur:
 '       If err = 62 Then
 
  '      chercher_dans_Lettres_Z = 0
        'iciiciici
'         MyInputBox.fichiers_consult�s.AddItem filename
   '     Exit Function
        
        
    '    End If
'''''''''''''''''''''''''''''''''

'Dim dbNorthwind As DAO.Database

Dim MyParam�tres As Recordset
Dim LastAb



'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 

   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
      
    Set MyParam�tres = dbNorthwind.OpenRecordset("zLetters")
  
  With MyParam�tres
  .Index = "forme"
  .Seek "=", MyMot
  'get_settings_from_bdd = MyParam�tres.Fields(MyField_Param�tres)
  End With

If MyParam�tres.NoMatch = True Then
            'MyInputBox.fichiers_consult�s.AddItem filename
            
             chercher_dans_Lettres_Z = 0
    
Else
         chercher_dans_Lettres_Z = -1
    
End If

''''''''''''''''''''''''''''''''




End Function


Public Sub cr�er_cet_ordinateur()
SaveSetting appname:="fasttype", section:="param�tres", Key:="cet ordinateur", setting:="maison"
SaveSetting appname:="fasttype", section:="param�tres", Key:="ordinateur last saving", setting:="maison"
End Sub

Public Sub load_accords()
    Load accords
                                accords.suggestions.AddItem MyNewWord
                                accords.�tiquette = "Ce mot semble correspondre � l'abr�viation introduite. Appuyer sur une lettre quelconque du clavier pour valider ou double-cliquez dessus"
                                accords.origine = "Mot d�duit de vos propres abr�viations"
                                accords.suggestions = MyNewWord
                                accords.BackColor = &HFF0000
                                accords.Show
                                
                                   Select Case MyPbkMsg
                                         
                                         
                                        
                                        
                                        Case "xxxxx" 'on n'a pas trouv� l'abr�viation
                                        
                                        'la proc�dure se poursuit dans le sub "abr�ger"
                                        
                                        
                                        Case Else
                                         
                                            If InStr(MyPbkMsg, "xxx changer abr�viation") > 0 Then
                                            
                                               MyInputBox.zone_mot = MyNewWord
                                               MyInputBox.zone_abr�viation = myab
                                               
                                               Unload accords
                                               MyPbkMsg = "xxxxx"
                                               Exit Sub
                                            
                                                                                        
                                            End If 'InStr(MyPbkMsg, "xxx changer abr�viation") > 0 Then
                                            
                                            If InStr(MyPbkMsg, "xxx Ne pas abr�ger") > 0 Then
                                            
                                                Selection.TypeText Text:=Left(MyPbkMsg, InStr(MyPbkMsg, "xxx Ne pas abr�ger") - 1)
        
                                                Selection.MoveRight Unit:=wdCharacter, Count:=1
                                            
                                                Unload accords
                                                End
                                                
                                            End If 'InStr(MyPbkMsg, "xxx Ne pas abr�ger") > 0 Then
                                         
                                         
                                  '      AutoCorrect.Entries.Add myab, MyPbkMsg
                                      '  Application.ActiveDocument.Activate
            
                                    '    Selection.TypeText Text:=MyPbkMsg
        
                                     '   Selection.MoveRight Unit:=wdCharacter, Count:=1
                                        Unload accords
                                        Exit Sub
                                    
                                    End Select 'MyPbkMsg
                                
                                
End Sub


Public Function d�tecter_apostrophe(myab)
Dim MyLetter, MySettingAccords, i, MyLetter2
MyPosition = ""
'MyAb = InputBox("abr�viation")
MyLetter = Mid(myab, 2, 1)
MyLetter2 = Mid(myab, 3, 1)

MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '

            For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                      
                        If MyLetter = (MySettingAccords((i), 0)) Then   'lettre de d�but
                           
                            MyApostrophe = ""
                            MyPosition = ""
                            GoTo troisi�me
                        End If
                   
            Next i


MyPosition = 2
If Len(myab) < 3 Then


                            MyApostrophe = " "
                            MyPosition = ""
                            d�tecter_apostrophe = False
                            Exit Function
End If


troisi�me:

For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                      
                        If MyLetter2 = (MySettingAccords((i), 0)) Then   'lettre de d�but
                           
                            MyApostrophe = ""
                            If MyPosition = 2 Then
                                d�tecter_apostrophe = True
                                 MyApostrophe = ""

                                Exit Function
                            Else
                                d�tecter_apostrophe = False
                                 MyApostrophe = " "
'                                MsgBox d�tecter_apostrophe, , MyPosition
                                Exit Function
                            End If
                        End If
                   
            Next i

d�tecter_apostrophe = True
MyPosition = 3
' MsgBox d�tecter_apostrophe, , MyPosition















End Function

Public Sub cr�er_verbe_deuxi�me_groupe()
'On Error GoTo erreur
'
'Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
'Dim FirstLetter, MyMot, trouv�, filename2, ts2, s2, Myindex2, MyDestinationFile2, motprisencompte, MySettingAccords
'ReDim MyMot(1) '
'Dim LastLetter, h, MyExistingAb
'trouv� = 1
'MyCompteurMots = 0
'
'
'
''FileName = get_hd & ":\mots\verbes_premier_groupe.txt"
''filename2 = get_hd & ":\mots\tous_les_mots.txt"
'
''Filename = "f:\essai.txt"
'Set fso = CreateObject("Scripting.FileSystemObject")
''Set ts = fso.OpenTextFile(FileName, ForReading)
''Set ts2 = fso.OpenTextFile(filename2, ForReading)
'
'
'
'       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
''s = ts.readline 'lit la premi�re ligne
'
'   Set MyDestinationFile2 = fso.CreateTextFile(get_hd & ":\mots\verbes_deuxi�me_groupe_certifi�s.txt")
'
'
'
'MySettingAccords = GetAllSettings(appname:="fasttype", section:="verbe_deuxi�me") '
'
'            For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
'
'
'                        MyDestinationFile2.WriteLine (MySettingAccords((i), 1))
'
'            Next i
'
'
'MsgBox "termin�"
'erreur:
'        If err = 62 Then
'
'
'        Exit Sub
'
'
'        End If
'
'
'
'
'




End Sub

Public Sub load_m�thode()

Dim MySetting, i, MyNumber
Dim Mysons(100, 2), strsql
Dim rdshippers As Recordset

i = 0
'chargement des lettres


Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
             strsql = "SELECT m�thode_ab.Valeur, m�thode_ab.Abr�viation, m�thode_ab.d�but_mot,m�thode_ab.milieu_mot,m�thode_ab.terminaison FROM m�thode_ab ;"
             Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 


    
                            rdshippers.MoveFirst
                                    While rdshippers.EOF = False
                                  
                                    m�thode_ab.MySound.AddItem rdshippers.Fields("valeur").Value
                                    m�thode_ab.MySound.List(i, 1) = rdshippers.Fields("abr�viation").Value
                                    m�thode_ab.MySound.List(i, 2) = rdshippers.Fields("d�but_mot").Value
                                    m�thode_ab.MySound.List(i, 3) = rdshippers.Fields("milieu_mot").Value
                                    m�thode_ab.MySound.List(i, 4) = rdshippers.Fields("terminaison").Value
                                    rdshippers.MoveNext
                                    i = i + 1
                                Wend



      
'MySetting = GetAllSettings(appname:="fasttype", section:="terminaisons") '
'
'            For i = LBound(MySetting, 1) To UBound(MySetting, 1)
'
'                 m�thode_ab.terminaison.AddItem (MySetting((i), 0))
'                 m�thode_ab.terminaison.List(i, 1) = (MySetting((i), 1))
'
'            Next i
         
MySetting = GetAllSettings(appname:="fasttype", section:="accords") '

For i = LBound(MySetting, 1) To UBound(MySetting, 1)
                
                 m�thode_ab.accord.AddItem Replace((MySetting((i), 0)), "_", " ")
                 m�thode_ab.accord.List(i, 1) = (MySetting((i), 1))
                                               
            Next i
            
            
MySetting = GetAllSettings(appname:="fasttype", section:="personnes") '

For i = LBound(MySetting, 1) To UBound(MySetting, 1)
                
                 m�thode_ab.personnes.AddItem Replace((MySetting((i), 0)), "_", " ")
                 m�thode_ab.personnes.List(i, 1) = (MySetting((i), 1))
                                               
            Next i
            
MySetting = GetAllSettings(appname:="fasttype", section:="temps_lettre") '
Dim A
For i = LBound(MySetting, 1) To UBound(MySetting, 1)
        If MySetting((i), 1) <> "x" Then
        A = A + 1
                 m�thode_ab.temps_lettre.AddItem Replace((MySetting((i), 0)), "_", " ")
                 m�thode_ab.temps_lettre.List(A - 1, 1) = (MySetting((i), 1))
        End If
        
            Next i
            
'm�thode_ab.son.ListRows = m�thode_ab.son.ListCount
'm�thode_ab.terminaison.ListRows = m�thode_ab.terminaison.ListCount
'
'
'
'm�thode_ab.terminaison.TabIndex = 0
'm�thode_ab.son.TabIndex = 1







m�thode_ab.Show



End Sub





Public Function lire_fichier_verbe(MyVerb, MyNumberConjugaison)
     
        'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, MyPlace, MyApos

filename = get_hd & ":\mots\verbes\" & MyVerb & ".txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(filename) Then

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
       
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

MyPlace = InStr(1, s, MyNumberConjugaison)
MyApos = InStr(MyPlace + Len(MyNumberConjugaison) + 1, s, ";")
MyValue = Mid(s, MyPlace + Len(MyNumberConjugaison) + 1, MyApos - MyPlace - Len(MyNumberConjugaison) - 1)

lire_fichier_verbe = MyValue
Else
lire_fichier_verbe = ""
End If

  
                           

End Function


Sub lancer_lire_fichier_verbe()
Dim MyForm, MyNumber
MyNumber = GetSetting("fasttype", section:="temps_combinaison", Key:="riv")
MyForm = lire_fichier_verbe("randonner", MyNumber)
MsgBox MyForm


End Sub


Public Sub cr�er_lettres_temps()
'pr�sent indicatif
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rj", setting:="11"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rt", setting:="12"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="ri", setting:="13"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rn", setting:="14"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rv", setting:="15"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rp", setting:="16"
'pass� compos�
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpj", setting:="21"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpt", setting:="22"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpi", setting:="23"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpn", setting:="24"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpv", setting:="25"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rpp", setting:="26"
 'imparfait indicatif
  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rij", setting:="31"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rit", setting:="32"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rii", setting:="33"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rin", setting:="34"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="riv", setting:="35"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rip", setting:="36"
 'plus-que-parfait indicatif
  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqj", setting:="41"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqt", setting:="42"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqi", setting:="43"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqn", setting:="44"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqv", setting:="45"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rqp", setting:="46"
 'pass� simple
  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="raj", setting:="51"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rat", setting:="52"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rai", setting:="53"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="ran", setting:="54"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rav", setting:="55"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rap", setting:="56"
 'futur simple
  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rfj", setting:="61"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rft", setting:="62"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rfi", setting:="63"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rfn", setting:="64"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rfv", setting:="65"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rfp", setting:="66"
 'futur ant�rieur
'  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rj", setting:="11"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rt", setting:="12"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="ri", setting:="13"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rn", setting:="14"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rv", setting:="15"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rp", setting:="16"
 
 'pass� ant�rieur"
'  SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rj", setting:="11"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rt", setting:="12"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="ri", setting:="13"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rn", setting:="14"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rv", setting:="15"
' SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rp", setting:="16"
'conditionnel pass� 1
SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rlj", setting:="91"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rlt", setting:="92"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rli", setting:="93"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rln", setting:="94"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rlv", setting:="95"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rlp", setting:="96"
'subjonctif pr�sent
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rsj", setting:="101"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rst", setting:="102"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rsi", setting:="103"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rsn", setting:="104"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rsv", setting:="105"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rsp", setting:="106"
'subjonctif imparfait

'subjonctif plus que parfait

'conditionnel pr�sent
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rcj", setting:="141"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rct", setting:="142"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rci", setting:="143"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rcn", setting:="144"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rcv", setting:="145"
 SaveSetting appname:="fasttype", section:="temps_combinaison", Key:="rcp", setting:="146"
 
 



End Sub

Public Function chercher_verbe(MyPath, D�butFile, FinFile, Middletters)

'FirstLetter, terminaisons(i), LettresDuMilieu(z)
'MyPath, D�butFile, FinFile, middleletters

On Error GoTo erreur

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
Dim FirstLetter, MyMot, trouv�, TailleD�butFile, TailleFinFile, AvantMiddle, Apr�sMiddle
ReDim MyMot(1)
Dim LastLetter, h, MyExistingAb
trouv� = 1

TailleD�butFile = Len(D�butFile)
TailleFinFile = Len(FinFile)

    Select Case TailleD�butFile
    
        Case 1
        
        AvantMiddle = "?"
        
        Case 2
        
        AvantMiddle = "??"
        
        Case 3
        
        AvantMiddle = "???"
        
        
        Case 4
        
        AvantMiddle = "????"
        
        
        Case 5
        
        AvantMiddle = "?????"
        
        Case 6
        
          AvantMiddle = "??????"
        
        Case 7
        
           AvantMiddle = "???????"
    
    
    End Select 'tailled�butfile



Select Case TailleFinFile

Case 1

Apr�sMiddle = "?"

Case 2
Apr�sMiddle = "??"
Case 3
Apr�sMiddle = "???"
Case 4
Apr�sMiddle = "????"
Case 5
Apr�sMiddle = "?????"
Case 6
Apr�sMiddle = "??????"
Case 7
Apr�sMiddle = "???????"
Case 8
Apr�sMiddle = "????????"
End Select 'taillefinfile


filename = get_hd & ":\mots\infinitifs\" & MyPath & ".txt"

Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)
'MsgBox ts.Size

       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 7543 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
    
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                'If MyValue Like FirstLetter & "*" & "b" & "*" & "c" & "*" & LastLetter Then
                If MyValue Like D�butFile & "*" & Middletters & "*" & FinFile Then
                   ' If Len(MyValue) / 3 >= Len(myab) And Len(myab) >= 4 Then GoTo MotExistant  'id�e si l'abr�viation ne fait que trois lettres
                    'de ne pas retenir les mots de plus de 9 lettres
                
                If MyInputBox.suggestions.ListCount > 0 Then
                    For h = 1 To MyInputBox.suggestions.ListCount
                       If s = MyInputBox.suggestions.List(h - 1) Then
                       GoTo MotExistant
                       End If
                
                    Next h
                    End If
             
             
                     ' If myinputbox.suggestions.ListCount = 20 And MyPbkMsg < 1 And EndIsAccord <> -1 And zLettresDuMilieuBrutes <> -1 Then
                      '      sMessage "Nombre d'abr�viation sup�rieur � 20, continuer ?", "rien", "stop", "continuer", "rien", "rien", "bleu", 2
                       '         Select Case MyPbkMsg
                                    
                        '                Case 2
                      
                         '               MyInputBox.zone_abr�viation = MyAb
                          '              chercher_utilisation_abr�viation (MyAb)
                           '                 For h = 1 To myinputbox.suggestions.ListCount
                            '                    MyInputBox.Mots_trouv�s.AddItem myinputbox.suggestions.List(h - 1)
                             '                   'MyInputBox.Mots_trouv�s.List(h - 1, 3) = MyInfinitifs(h)
                                                'myinputbox.suggestions.List(h - 1, 3) = MyInfinitifs(h)
                                            
                              '              Next h
                                        
                                        
                               '         MyInputBox.texte = "Nombre d'abr�viations excessif"
                                '        MyInputBox.Show
                                 '       End
                            
                                'End Select ' MyPbkMsg
                                
                     ' End If
              
              
                MyInputBox.suggestions.AddItem MyValue
                'MyExistingAb = check_existence_nom_pour_abr�viation(MyValue)
               ' If MyExistingAb <> False Then
                '  myinputbox.suggestions.List(h - 1, 1) = MyExistingAb
                 ' Else
                  
                'End If
             
                    
               
                End If
MotExistant:
   
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
Exit Function

erreur:
        If err = 62 Then
 'iciiciici
        'MyInputBox.fichiers_consult�s.AddItem filename & " " & MyCompteurMots '& " " & ts.Size
        'myinputbox.compteur = myinputbox.compteur + MyCompteurMots
          
            
            
            
            
        If h > 0 Then MyInputBox.suggestions = MyInputBox.suggestions.List(h - 1)
        Exit Function
        
        
        End If





End Function

Public Sub cr�er_infinitifs_par_lettre(LettreD�but)

On Error GoTo erreur:

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\infinitifs.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\infinitifs\" & LettreD�but & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 7543 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" Then
                'MyCompteurMots = MyCompteurMots + 1
               
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
                   
'Set fso = Nothing
'If MyPreviousNumber = "" Then MyPreviousNumber = 0
'SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
'MsgBox "end " & MyLetter & " " & MyTerminaison
Exit Sub

erreur:

  If err = 62 Then
       ' myinputbox.lettresMilieu.AddItem middleletters
        'myinputbox.fichiers_consult�s.AddItem FileName & " " & MyCompteurMots
        'myinputbox.compteur = myinputbox.compteur + MyCompteurMots
        
        Exit Sub
        
      
        End If



End Sub





Sub cr�er_zLetters_des_verbes()


Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber
Dim FirstLetter, MyMot, trouv�
ReDim MyMot(1)
Dim LastLetter, h, MyExistingAb
trouv� = 1



                             MyCompteurMots = 0


'FileName = get_hd & ":\mots\zletters.txt"

filename = get_hd & ":\mots\infinitifs.txt"


'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\zLetters_des_verbes.txt")
        

       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 7453 - 1 '336530 - 1 ' 'And MyTroyv�Premier <> 0
    MyIndex = MyIndex + 1
    
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                
                
                'If myvalue Like FirstLetter & "*" & "b" & "*" & "c" & "*" & LastLetter Then
                If MyValue Like "*asa*" Or MyValue Like "*ase*" Or MyValue Like "*asi*" Or MyValue Like "*aso*" Or MyValue Like "*asu*" _
                Or MyValue Like "*as�*" Or MyValue Like "*as�*" Or MyValue Like "*as�*" _
                Or MyValue Like "*esa*" Or MyValue Like "*ese*" Or MyValue Like "*esi*" Or MyValue Like "*eso*" Or MyValue Like "*esu*" _
                Or MyValue Like "*isa*" Or MyValue Like "*ise*" Or MyValue Like "*isi*" Or MyValue Like "*iso*" Or MyValue Like "*isu*" _
                Or MyValue Like "*is�*" Or MyValue Like "*is�*" Or MyValue Like "*is�*" _
                Or MyValue Like "*osa*" Or MyValue Like "*ose*" Or MyValue Like "*osi*" Or MyValue Like "*oso*" Or MyValue Like "*osu*" _
                Or MyValue Like "*osy*" Or MyValue Like "*os�*" Or MyValue Like "*os�*" Or MyValue Like "*os�*" _
                Or MyValue Like "*ysa*" Or MyValue Like "*yse*" Or MyValue Like "*ysi*" Or MyValue Like "*yso*" Or MyValue Like "*ysu*" _
                Or MyValue Like "*ys�*" Or MyValue Like "*ys�*" Or MyValue Like "*ys�*" _
                Or MyValue Like "*usa*" Or MyValue Like "*use*" Or MyValue Like "*usi*" Or MyValue Like "*uso*" Or MyValue Like "*usu*" _
                Or MyValue Like "*us�*" Or MyValue Like "*us�*" Or MyValue Like "*us�*" _
                Or MyValue Like "*�sa*" Or MyValue Like "*�se*" Or MyValue Like "*�si*" Or MyValue Like "*�so*" Or MyValue Like "*�sy*" _
                Or MyValue Like "*�sy*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" _
                Or MyValue Like "*�sa*" Or MyValue Like "*�se*" Or MyValue Like "*�si*" Or MyValue Like "*�so*" Or MyValue Like "*�sy*" _
                Or MyValue Like "*�sy*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*z*?" _
                Then
            
                MyDestinationFile.WriteLine MyValue
                    
               
                End If

   
               s = ts.readline '
               
               ' MyEnd = InStr(1, s, "ZYTHyMS")
    Loop 'While MyEnd = 0
     'myinputbox.fichiers_consult�s.AddItem FileName & " " & MyCompteurMots
      '  myinputbox.compteur = myinputbox.suggestions.ListCount
       ' myinputbox.nbre_fichiers_consult�s = myinputbox.fichiers_consult�s.ListCount
        'myinputbox.Show






End Sub

Function IsVerb(MyVerbe)

'Dim dbNorthwind As DAO.Database

Dim MyParam�tres As Recordset
Dim LastAb

'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
      
    Set MyParam�tres = dbNorthwind.OpenRecordset("table_m�re_des_verbes")
  
  With MyParam�tres
  .Index = "forme"
  .Seek "=", MyVerbe

  End With

If MyParam�tres.NoMatch = True Then
            'MyInputBox.fichiers_consult�s.AddItem filename
            
             IsVerb = 0
    
Else
            IsVerb = -1
    
End If




End Function

Public Sub cr�ation_d�but_lettres_fin_f�minin(lettre1, lettre2)

        'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettres_fin_f�minin\" & lettre1 & "XXX" & lettre2 & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like lettre1 & lettre2 & "*" & "e" Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0

End Sub

Public Sub lancer_d�but_lettres_fin_f�minin()
Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3


MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") 'il y a l� toutes les lettres de l'alphabet
        
           For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
            'For m = 20 To UBound(MySettingAccords, 1)
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                            myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                            myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                            myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_lettres_fin_f�minin myval2, myval3
                     
                    
                       Next n
           
                
        Next m

End Sub


Public Sub UpDateLastAb()
'Dim dbNorthwind As DAO.Database
Dim rdshippers As Recordset
Dim MyParam�tres As Recordset
Dim intRecords 'As Integer

 
 
 
 
  'MyControls = AutoCorrect.Entries.Count
 '
 
 
 
 'Set fso = CreateObject("scripting.fileSystemObject")
 ''''''''''''''''''


'SaveSetting appname:="fasttype", section:="param�tres", Key:="date_usage", setting:=Date


   
   
   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
     
    Set MyParam�tres = dbNorthwind.OpenRecordset("param�tres")
  
  With MyParam�tres
  .Index = "PrimaryKey"
  .Seek "=", 1
  .Edit
  !date_usage = Date
  !LastAb = get_param�tres("cet ordinateur")
  !MyHeure = Time
  .Update
  End With

End Sub

Function read_lastab()

Dim fso          ' As Scripting.FileSystemObject
 Dim ts           ' As Scripting.TextStream


 
 Dim strline, bret, strDest As String
 Dim s, MyEnd, MyBegin, MyAutoCorrects, j, MyName, MyValue, filename, sDelete, CompteR�cup, sExists
 
 sDelete = 0
 CompteR�cup = 0
 'il faut d'abord effacer toutes les abr�viations du fichier des abr�viations,
 'car sinon cela doublonne
 
 
 'sExiste = fso.FileExists(sFileName2) 'on teste si le fichier existe


'If sExiste = True Then Kill sFileName2 'si le fichier n'existe pas, il sera cr�� automatiquement

 
 
 
filename = get_hd & ":\mots\lastab.txt"
 
Set fso = CreateObject("Scripting.FileSystemObject")

sExists = fso.FileExists(filename)
If sExists = False Then
read_lastab = 0
Exit Function
End If

Set ts = fso.OpenTextFile(filename, ForReading)


s = ts.readline 'lit la premi�re ligne



If s Like get_param�tres("cet ordinateur") Then


read_lastab = -1


Else
read_lastab = 0

End If


 
End Function






Sub cr�er_tableau()
Attribute cr�er_tableau.VB_Description = "Macro enregistr�e le 13/06/2010 par Emmanuel"
Attribute cr�er_tableau.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.cr�er_tableau"
'
' cr�er_tableau Macro
' Macro enregistr�e le 13/06/2010 par Emmanuel
'
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:= _
        8, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Grille du tableau" Then
            .Style = "Grille du tableau"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = True
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = True
    End With
    Selection.Sort ExcludeHeader:=False, FieldNumber:="Colonne 1", _
        SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending, _
        FieldNumber2:="", SortFieldType2:=wdSortFieldAlphanumeric, SortOrder2:= _
        wdSortOrderAscending, FieldNumber3:="", SortFieldType3:= _
        wdSortFieldAlphanumeric, SortOrder3:=wdSortOrderAscending, Separator:= _
        wdSortSeparateByCommas, SortColumn:=False, CaseSensitive:=False, _
        LanguageID:=wdFrench, SubFieldNumber:="Paragraphes", SubFieldNumber2:= _
        "Paragraphes", SubFieldNumber3:="Paragraphes"
End Sub

Public Sub stocker_toutes_les_abr�viations()

Dim fso ' As Scripting.FileSystemObject

 Dim ts, fd, fsp   ' As Scripting.TextStream

 Dim str, sFileName, sFileName2, sFile, sFichier, sExiste, sFichier2 ' As String
 
 Dim snom, sValeur, sLigne, mycontrolsn, i, MyControls, mynamecontrol
  MyControls = AutoCorrect.Entries.Count
 '
 
 GoTo saute_fichier
 
 Set fso = CreateObject("scripting.fileSystemObject")
 ''''''''''''''''''


SaveSetting appname:="fasttype", section:="param�tres", Key:="date_usage", setting:=Date

sFileName = get_hd & ":\abr�viations.txt"
sFileName2 = get_hd & ":\abr�viations\abr�viations.txt"

Set sFichier = fso.GetFile(sFileName)
Set sFichier2 = fso.GetFile(sFileName2)

fso.CopyFile sFileName, get_hd & ":\abr�viations\abr�viations.txt", True
'sExiste = fso.FileExists(sFileName2) 'on teste si le fichier existe


'If sExiste = True Then Kill sFileName2 'si le fichier n'existe pas, il sera cr�� automatiquement


'on renome data.mdb en data1.mdb

'sFichier2.Name = mycontrols & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".txt"


 
saute_fichier:
 


' sFileName = "f:\essai.txt"

'Set ts = fso.OpenTextFile(sFileName, ForWriting, True)
    
   
    
   MyControls = AutoCorrect.Entries.Count
   For i = 1 To MyControls - 1

    'ts.WriteLine AutoCorrect.Entries(i).Name & " ==== " & AutoCorrect.Entries(i).Value ''' ligne � remettre pour fichier
    
       ' ajout_entr�e_ab AutoCorrect.Entries(i).Name, AutoCorrect.Entries(i).Value, False

   
    Next i
    

    
'MsgBox "les " & i & " abr�viations a �t� sauvegard�es dans le fichier " & sFileName, vbOKOnly, "FasType"

'fso.CopyFile sFileName, get_hd & ":\abr�viations\abr�viations.txt", True

UpDateLastAb



End Sub

Public Sub stocker_verbes_automatiques()
Dim fso ' As Scripting.FileSystemObject

 Dim ts, fd, fsp   ' As Scripting.TextStream

 Dim str, sFileName, sFileName2, sFile, sFichier, sExiste, sFichier2 ' As String
 
 Dim snom, sValeur, sLigne, mycontrolsn, i, MyControls, mynamecontrol, MyNumberVerbes
  MyControls = AutoCorrect.Entries.Count
 '
 
 
 
' Set fso = CreateObject("scripting.fileSystemObject")
 ''''''''''''''''''


SaveSetting appname:="fasttype", section:="param�tres", Key:="date_usage", setting:=Date

'sFileName = get_hd & ":\abr�viations.txt"
'sFileName2 = get_hd & ":\abr�viations\abr�viations.txt"

'Set sFichier = fso.GetFile(sFileName)
'Set sFichier2 = fso.GetFile(sFileName2)

'fso.CopyFile sFileName, get_hd & ":\abr�viations\abr�viations.txt", True
 
'sFichier2.Name = MyControls & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".txt"


 
 
 'totototo
 '        Set ts = fso.OpenTextFile(sFileName, ForAppending, True)
 
     MyNumberVerbes = MyInputBox.suggestions.ListCount
            
    For i = 1 To MyNumberVerbes
   ' While FirstLetter = Left(AutoCorrect.Entries(i).Name, 1)
 
   ' ajout_entr�e_ab MyInputBox.suggestions.List(i - 1, 0), MyInputBox.suggestions.List(i - 1, 1), True
   ' FirstLetter = Left(AutoCorrect.Entries(i + 1).Name, 1)
    'GoTo next_
    'Wend
      
    ' sFileName = get_hd & ":\abr�viations\" & FirstLetter & ".txt"
    ' Set ts = fso.OpenTextFile(sFileName, ForWriting, True)
'next_:
    
    Next i
    

    
'MsgBox "l'ensembles des abr�viations a �t� sauvegard� dans le fichier " & sFileName, vbOKOnly, "sauvegarde"

'fso.CopyFile sFileName, get_hd & ":\abr�viations\abr�viations.txt", True

UpDateLastAb
End Sub

Public Sub valeur_abr�viation()
myab = InputBox("mot")
myab = nettoyer_voyelle(myab)
Dim MyValeursLettres(19, 1), MyLen, MyValue, i, j, MyLetter

MyValeursLettres(0, 0) = "b"
MyValeursLettres(0, 1) = 1

MyValeursLettres(1, 0) = "c"
MyValeursLettres(1, 1) = 4

MyValeursLettres(2, 0) = "d"
MyValeursLettres(2, 1) = 8

MyValeursLettres(3, 0) = "f"
MyValeursLettres(3, 1) = 16

MyValeursLettres(4, 0) = "g"
MyValeursLettres(4, 1) = 32

MyValeursLettres(5, 0) = "h"
MyValeursLettres(5, 1) = 64

MyValeursLettres(6, 0) = "j"
MyValeursLettres(6, 1) = 128

MyValeursLettres(7, 0) = "k"
MyValeursLettres(7, 1) = 256

MyValeursLettres(8, 0) = "l"
MyValeursLettres(8, 1) = 512

MyValeursLettres(9, 0) = "m"
MyValeursLettres(9, 1) = 1024

MyValeursLettres(10, 0) = "n"
MyValeursLettres(10, 1) = 2048

MyValeursLettres(11, 0) = "p"
MyValeursLettres(11, 1) = 4096

MyValeursLettres(12, 0) = "q"
MyValeursLettres(12, 1) = 8192

MyValeursLettres(13, 0) = "r"
MyValeursLettres(13, 1) = 16384

MyValeursLettres(14, 0) = "s"
MyValeursLettres(14, 1) = 32768

MyValeursLettres(15, 0) = "t"
MyValeursLettres(15, 1) = 65536

MyValeursLettres(16, 0) = "v"
MyValeursLettres(16, 1) = 131072

MyValeursLettres(17, 0) = "w"
MyValeursLettres(17, 1) = 262144

MyValeursLettres(18, 0) = "x"
MyValeursLettres(18, 1) = 524288

MyValeursLettres(19, 0) = "z"
MyValeursLettres(19, 1) = 1048576

'MyValeursLettres(20, 0) = "u"
'MyValeursLettres(20, 1) = 2097152


'MyValeursLettres(21, 0) = "v"
'MyValeursLettres(21, 1) = 4194304

'MyValeursLettres(22, 0) = "w"
'MyValeursLettres(22, 1) = 8388608
'
'MyValeursLettres(23, 0) = "x"
'MyValeursLettres(23, 1) = 16777216

'MyValeursLettres(24, 0) = "y"
'MyValeursLettres(24, 1) = 33554432

'MyValeursLettres(25, 0) = "z"
'MyValeursLettres(25, 1) = 67108864

'MyValeursLettres(26, 0) = "�"
'MyValeursLettres(26, 1) = 67108864

'MyValeursLettres(27, 0) = "�"
'MyValeursLettres(27, 1) = 268435456

'MyValeursLettres(28, 0) = "�"
'MyValeursLettres(28, 1) = 536870912

'MyValeursLettres(29, 0) = "�"
'MyValeursLettres(29, 1) = 1073741824

'MyValeursLettres(30, 0) = "�"
'MyValeursLettres(30, 1) = 2147483648#


MyLen = Len(myab)

'reprendre ici
MyValue = 0
For i = 1 To MyLen

  MyLetter = Mid(myab, i, 1)
        
        
         For j = LBound(MyValeursLettres, 1) To UBound(MyValeursLettres, 1)
         
            If MyLetter = MyValeursLettres(j, 0) Then
            MyValue = MyValue + MyValeursLettres(j, 1)
            End If
            
         
         Next j
    

Next i

MsgBox MyValue

    






















End Sub

Public Sub load_marqueurs()



Dim j, MyName, MyValue, s, MyAutoCorrects, filename, fso, ts, MyIndex
Dim folder, subflds, fld, fl As file, MyLen, MyInternalIndex, i, strsql
'  Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
        
Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")


    
    
    
    MyIndex = MyIndex + 1
        

 
   
    strsql = "SELECT marqueurs.�l�ment FROM marqueurs order by �l�ment;"
    
    
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

    With rdshippers
      Do While Not .EOF
      i = i + 1
          If IsNull(rdshippers.Fields(0)) = False Then
        marqueurs.�l�ments.AddItem rdshippers.Fields(0)
         marqueurs.�l�ments.List(i - 1, 1) = rdshippers.Fields(0)
        
            End If
        
         .MoveNext
      Loop
   End With
    
    

    
End Sub
Public Sub ajouter_item_dans_marqueurs(MyItem)

Dim fso ' As Scripting.FileSystemObject

 Dim ts, fd, fsp   ' As Scripting.TextStream

 Dim str, sFileName, sFileName2, sFile, sFichier, sExiste, sFichier2 ' As String
 
 Dim snom, sValeur, sLigne, mycontrolsn, i, MyControls, mynamecontrol

 '
 
 
 
 Set fso = CreateObject("scripting.fileSystemObject")
 ''''''''''''''''''




sFileName = get_hd & ":\mots\marqueurs.txt"


        Set ts = fso.OpenTextFile(sFileName, ForAppending, True)
 
     
            
            

    ts.WriteLine MyItem
End Sub


Public Function nettoyer_voyelle(MyMot)

Dim MySettingAccords, i




MySettingAccords = GetAllSettings(appname:="fasttype", section:="voyelles") '
        
         For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)

                
            MyMot = Replace(MyMot, (MySettingAccords((i), 0)), "")

         Next i




nettoyer_voyelle = Len(MyMot)




End Function
Sub envoi_mail()
Attribute envoi_mail.VB_Description = "Macro enregistr�e le 18/07/2010 par Emmanuel"
Attribute envoi_mail.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.envoi_mail"
'
' envoi_mail Macro
' Macro enregistr�e le 18/07/2010 par Emmanuel
'
End Sub

Public Function contient_une_voyelle(MyString)

Dim LenMyString, i, j, MySettingAccords

LenMyString = Len(MyString)

    
    For i = 1 To LenMyString
    
        MySettingAccords = GetAllSettings(appname:="fasttype", section:="voyelles") '
  

        For j = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)

            If Mid(MyString, i, 1) = (MySettingAccords((j), 0)) Then
                contient_une_voyelle = -1
                Exit Function
            
            End If

        Next j
       
    
    
    
    Next i

contient_une_voyelle = 0
  
End Function


Public Sub log_recherche(MyDate, heure, myab, combinatoire, temps, nombreMotsPropos�s, NombreMotsExclus, temps_par_combinaison, D�but, Milieu, fin, temps_exclusions)


Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\log_recherches.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForAppending, True)
        ts.WriteLine MyDate & ";" & heure & ";" & myab & ";" & combinatoire & ";" & temps & ";" & nombreMotsPropos�s & ";" & NombreMotsExclus & ";" & temps_par_combinaison & ";" & D�but & ";" & Milieu & ";" & fin & ";" & ";" & temps_exclusions
        

End Sub


Public Function cr�er_lettres_sans_sons()
'cette fonction permet de produire un string qui comporte toutes les lettres non abr�viatives.
'c'est cette chaine qui est retourn�e.
'le but de la fonction est de toujours produire une liste � jour de ces touches.
Dim i, j, alphabet, sons, NonSon, MyString
NonSon = 0

        alphabet = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
        sons = GetAllSettings(appname:="fasttype", section:="sons") '


        For j = LBound(alphabet, 1) To UBound(alphabet, 1)
            
            For i = LBound(sons, 1) To UBound(sons, 1)
 '           accords.suggestions.AddItem sons((i), 1)
                If alphabet((j), 0) <> sons((i), 1) Then
                NonSon = NonSon + 1
                End If
                
            
            
            Next i
           If NonSon = UBound(sons, 1) + 1 Then
'           accords.suggestions.AddItem alphabet((j), 0)
           MyString = MyString & alphabet((j), 0)
           
           End If
           
            NonSon = 0
        Next j

cr�er_lettres_sans_sons = MyString
'MsgBox mystring
'accords.Caption = accords.suggestions.ListCount
'accords.Show

End Function
Function cr�ation_fichier_d�but_lettre_fin_lettre_lettre_milieu(LettreD�but, LettreFin, lettreMilieu)






Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettre_fin_lettre\" & LettreD�but & lettreMilieu & LettreFin & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" & lettreMilieu & "*" & LettreFin Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
                   
'Set fso = Nothing
'If MyPreviousNumber = "" Then MyPreviousNumber = 0
'SaveSetting appname:="fasttype", section:="nombre_mots", Key:=MyLetter, setting:=MyCompteurMots + MyPreviousNumber
'MsgBox "end " & MyLetter & " " & MyTerminaison
End Function
Sub lancer_cr�ation_fichier_d�but_lettre_fin_lettre_lettre_milieu()


Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3, MyString, s
Dim MyLettreDuMilieu
MyString = cr�er_lettres_sans_sons


For s = 1 To Len(MyString)
MyLettreDuMilieu = Mid(MyString, s, 1)


MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") 'il y a l� toutes les lettres de l'alphabet
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
        
                  For n = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                        ' myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_fichier_d�but_lettre_fin_lettre_lettre_milieu myval2, myval3, MyLettreDuMilieu
                     
                    
                       Next n
           
            
            
                
        Next m


Next s




End Sub





Public Sub cr�ation_d�but_son_fin_terminaison_lettre_milieu(SonD�but, SonTerminaison, lettre_milieu)


Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, j

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

       
        '& LettreD�but
     
       ' Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_son_fin_terminaison\" & SonD�but & lettre_milieu & SonTerminaison & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125447 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like SonD�but & "*" & lettre_milieu & "*" & SonTerminaison Then
                'MyCompteurMots = MyCompteurMots + 1
                accords.stock.AddItem s
        
                'MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0
    If accords.stock.ListCount = 0 Then Exit Sub


     Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_son_fin_terminaison\" & SonD�but & lettre_milieu & SonTerminaison & ".txt")
       

    For j = 1 To accords.stock.ListCount
    
    MyDestinationFile.WriteLine accords.stock.List(j - 1)
    
    
    Next j

accords.stock.Clear


End Sub
Sub lancer_cr�ation_d�but_son_fin_terminaison_lettre_milieu()




Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3, MyString, s
Dim MyLettreDuMilieu
MyString = cr�er_lettres_sans_sons


For s = 1 To Len(MyString)
MyLettreDuMilieu = Mid(MyString, s, 1)


MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") 'il y a l� toutes les lettres de l'alphabet
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="terminaisons") '
        
                  For n = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                        ' myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                     cr�ation_d�but_son_fin_terminaison_lettre_milieu myval2, myval3, MyLettreDuMilieu
                     
                    
                       Next n
           
            
            
                
        Next m


Next s
MsgBox "termin�"
End Sub

Sub cr�ation_en_s�rie_lettres_du_milieu()

lancer_cr�ation_d�but_son_fin_terminaison_lettre_milieu
lancer_d�but_son_fin_lettre_lettre_du_milieu
MsgBox "termin�"

End Sub
Function cr�ation_d�but_son_fin_lettre_lettre_milieu(D�butSon, LettreFin, lettre_du_milieu)



        
        
     
        'Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\

Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)


        
        
     
        Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_son_fin_lettre\" & D�butSon & lettre_du_milieu & LettreFin & ".txt")
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like D�butSon & "*" & lettre_du_milieu & "*" & LettreFin Then
                'MyCompteurMots = MyCompteurMots + 1
                MyInputBox.zone_mot.AddItem s
        
                MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0


End Function
Sub lancer_d�but_son_fin_lettre_lettre_du_milieu()

Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3
Dim MyLettreDuMilieu, MyString, s
MyString = cr�er_lettres_sans_sons


For s = 1 To Len(MyString)
MyLettreDuMilieu = Mid(MyString, s, 1)


MySettingAccords = GetAllSettings(appname:="fasttype", section:="sons") '
        
            For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
           ' For m = 20 To UBound(MySettingAccords, 1)
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      cr�ation_d�but_son_fin_lettre_lettre_milieu myval2, myval3, MyLettreDuMilieu
                      
                     
                Next n
                      
            
                
        Next m

Next s


End Sub

Sub cr�ation_d�but_lettre_fin_terminaison_lettremilieu(LettreD�but, LettreFin, lettre_du_milieu)




Dim filename, fso, ts, MyDestinationFile, s, MyIndex, MyValue, MyCompteurMots, NombreMots, i, AjoutMot, MyPreviousNumber, j

filename = get_hd & ":\mots\tous_les_mots.txt"
'Filename = "f:\essai.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(filename, ForReading)

       
        '& LettreD�but
     
        


       ' MyPreviousNumber = GetSetting(appname:="fasttype", section:="nombre_mots", Key:=MyLetter)
       ' If mypreviounsnumber = "" Then MyPreviousNumber = 0
s = ts.readline 'lit la premi�re ligne

  
                           
    Do While MyIndex < 125457 - 1 ' 'And MyTrouv�Premier <> 0
    MyIndex = MyIndex + 1
                's = Replace(s, " ", "")
                'j = Len(s)
                'MyBegin = InStr(1, s, " ==== ")
             
                MyValue = s
                  
                If MyValue Like LettreD�but & "*" & lettre_du_milieu & "*" & LettreFin Then
                'MyCompteurMots = MyCompteurMots + 1
                'MyInputBox.zone_mot.AddItem s
                accords.stock.AddItem s
               ' MyDestinationFile.WriteLine s
                End If
                
                
               s = ts.readline '
               ' MyEnd = InStr(1, s, "ZYTHUMS")
    Loop 'While MyEnd = 0


    If accords.stock.ListCount = 0 Then Exit Sub


     Set MyDestinationFile = fso.CreateTextFile(get_hd & ":\mots\d�but_lettre_fin_terminaison\" & LettreD�but & lettre_du_milieu & LettreFin & ".txt")
       

    For j = 1 To accords.stock.ListCount
    
    MyDestinationFile.WriteLine accords.stock.List(j - 1)
    
    
    Next j

accords.stock.Clear




End Sub

Sub lancer_cr�ation_d�but_lettre_fin_terminaison_lettre_milieu()

Dim lancer, MySettingAccords, m, n, MySettingAccords2, myval, myval2, myval3, myval4, o, p, MySettingAccords3
Dim MyLettreDuMilieu, MyString, s
MyString = cr�er_lettres_sans_sons


For s = 1 To Len(MyString)
MyLettreDuMilieu = Mid(MyString, s, 1)


MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '
             For m = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)
           ' For m = 20 To UBound(MySettingAccords, 1)
                    'je d�sactive cette partie pour l'essai des suivantes
                MySettingAccords2 = GetAllSettings(appname:="fasttype", section:="terminaisons") '
        
                  For n = LBound(MySettingAccords2, 1) To UBound(MySettingAccords2, 1)
                
                          
                        myval2 = (MySettingAccords((m), 0)) 'lettre de d�but
                         myval3 = (MySettingAccords2((n), 0)) 'valeur de la lettre abr�geante (ex : ment (pour m))
                         myval4 = (MySettingAccords2((n), 1)) 'lettre abr�geante (ex : m (pour ment)
                          
                      'cr�ation_d�but_lettre_fin_terminaison_lettremilieu myval2, myval3, MyLettreDuMilieu
                      
                     
                Next n
                      
            
                
        Next m
Next s

MsgBox "termin�"
End Sub

Sub d�sactiver_correct(TrueOrFalse)
Attribute d�sactiver_correct.VB_Description = "Macro enregistr�e le 01/09/2010 par MINEFI"
Attribute d�sactiver_correct.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.d�sactiver_correct"
'
' d�sactiver_correct Macro
' Macro enregistr�e le 01/09/2010 par MINEFI
'
    With AutoCorrect
        .CorrectInitialCaps = True
        .CorrectSentenceCaps = True
        .CorrectDays = True
        .CorrectCapsLock = True
        .ReplaceText = False
        .ReplaceTextFromSpellingChecker = True
        .CorrectKeyboardSetting = TrueOrFalse
        .DisplayAutoCorrectOptions = True
        .CorrectTableCells = True
    End With
End Sub

Public Sub doc_et_t�ches()
 Dim fso, folder, subflds, fld, s, fl As file, MyIndex, MyLen, MyPath, i, MyInternalIndex, sFolderExists
 MyIndex = 0
MyFolders.dossiers.AddItem "Tous les dossiers"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder("f:\int�rieur\")
    Set subflds = folder.SubFolders

    s = ""
    For Each fld In subflds
   
        s = fld.Name
        MyFolders.dossiers.AddItem s
      
     
    Next
    MyPath = GetSetting("fasttype", section:="param�tres", Key:="LastFolder")
sFolderExists = fso.FolderExists("f:\int�rieur\" & MyPath) '

If sFolderExists = False Then
MyPath = "divers"

End If
    lire_sous_r�pertoire (MyPath)
    GoTo skip
    Set fld = fso.GetFolder("f:\int�rieur\" & MyPath & "\")
     
     s = ""
   ' For Each fld In subflds
        For Each fl In fld.Files
           ' MyIndex = MyIndex + 1
            'MyLen = Len(fl.Name)
               ' If Right(fl.Name, 3) <> "tmp" Then
            '''''''''''''
            If Right(fl.Name, 3) <> "tmp" Then
                    MyInternalIndex = MyInternalIndex + 1
            
            
                    MyFolders.fichiers.AddItem fl.Name
                      MyFolders.fichiers.List(MyInternalIndex - 1, 1) = fl.DateLastModified
                    MyFolders.fichiers.List(MyInternalIndex - 1, 2) = Replace(fl.ParentFolder, "f:\int�rieur\", "")
                
                End If
            
            
            ''''''''''''''
            maj_folder MyPath, "date_ascendante", "seulement ceux affich�s"
            
                   '  MyFolders.fichiers.AddItem fl.Name
            
                    'MyFolders.fichiers.List(MyIndex - 1, 1) = Replace(fl.ParentFolder, "f:\int�rieur\", "")
                'End If
        Next
    'Next
    'For Each fl In folder.Files
   
    'MyFolders.fichiers.AddItem fl.Name
    
    'Next

   



    
    
  
   
    
skip:
    
   
suite:
    
    maj_nombre_dossiers_fichiers
    
    
    
    peupler "type_docs"
    peupler "noms"
    peupler "version"
   
    peupler "texte"
     peupler "format"
     
    maj_nombre_dossiers_fichiers
   
MyFolders.MyDate = Date
MyFolders.Show
 MyFolders.dossiers = MyPath
   

End Sub

Sub peupler(MyChamp)


'On Error GoTo erreur
Dim j, MyName, MyValue, s, MyAutoCorrects, filename, fso, ts, MyIndex
Dim folder, subflds, fld, fl As file, MyLen, MyInternalIndex, i, strsql
  'Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
        
Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")



'filename = get_hd & ":\mots\" & MyChamp & ".txt"

Set fso = CreateObject("Scripting.FileSystemObject")

'Set ts = fso.OpenTextFile(filename, ForReading)


Select Case MyChamp
   
   Case "type_docs"

        MyFolders.type_doc.Clear

   Case "noms"
        
        MyFolders.noms.Clear
    
    Case "format"
      
        MyFolders.format.Clear
    
    Case "version"
    
        MyFolders.version.Clear
        
    Case "recherches"
   
        MyFolders.texte.Clear
   
   
End Select
    
    
    
   ' Do While MyIndex < 335000 - 1 ' 'And MyTrouv�Premier <> 0
    
    
    
    MyIndex = MyIndex + 1
        
    's = ts.readline
    
    Select Case MyChamp
    
    Case "type_docs"
    
 
   
    strsql = "SELECT noms.�l�ment FROM noms ORDER BY �l�ment;"
    
    
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

    With rdshippers
      Do While Not .EOF
      i = i + 1
          If IsNull(rdshippers.Fields(0)) = False Then
          MyFolders.type_doc.AddItem rdshippers.Fields(0)
          MyFolders.type_doc.List(i - 1, 1) = rdshippers.Fields(0)
        
            End If
        
         .MoveNext
      Loop
   End With
    

    Case "noms"
  
 strsql = "SELECT auteurs.�l�ment FROM auteurs order BY �l�ment;"
    
    
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

    With rdshippers
      Do While Not .EOF
      i = i + 1
          If IsNull(rdshippers.Fields(0)) = False Then
          MyFolders.noms.AddItem rdshippers.Fields(0)
          MyFolders.noms.List(i - 1, 1) = rdshippers.Fields(0)
        
            End If
        
         .MoveNext
      Loop
   End With
    
    Case "format"
  
       MyFolders.format.AddItem ".doc"
     MyFolders.format.AddItem ".pdf"
      MyFolders.format.AddItem ".xls"
       MyFolders.format.AddItem ".ppt"
        MyFolders.format.AddItem ".odt"
    
    Case "version"
    
    MyFolders.version.AddItem 1
     MyFolders.version.AddItem 2
      MyFolders.version.AddItem 3
       MyFolders.version.AddItem 4
        MyFolders.version.AddItem 5
  
  
    Exit Sub
   ' MyFolders.recherche.AddItem s
    
    
  '  Loop 'While MyEnd = 0

    Case "texte"
   
   
   strsql = "SELECT texte.�l�ment FROM texte order BY �l�ment;"
    
    
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 

    With rdshippers
      Do While Not .EOF
      i = i + 1
          If IsNull(rdshippers.Fields(0)) = False Then
          MyFolders.texte.AddItem rdshippers.Fields(0)
          MyFolders.texte.List(i - 1, 1) = rdshippers.Fields(0)
        
            End If
        
         .MoveNext
      Loop
        End With
   
   
   End Select
erreur:
        If err = 62 Then
       
    
        Exit Sub
        
        
        End If










End Sub


Sub maj_nombre_dossiers_fichiers()

MyFolders.nombre_dossiers.Caption = MyFolders.dossiers.ListCount - 1 & " dossiers"
MyFolders.nombre_fichier.Caption = MyFolders.fichiers.ListCount & " fichiers"



End Sub


Public Sub lire_sous_r�pertoire(MyPath)
 Dim fso, folder, subflds, fld, s, fl As file, MyIndex, MyLen, i, MyInternalIndex
 MyIndex = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
  '  Set folder = fso.GetFolder("f:\int�rieur\")
   ' Set subflds = folder.SubFolders

   ' s = ""
   ' For Each fld In subflds
   
    '    s = fld.Name
     '   MyFolders.dossiers.AddItem s
      
     
    'Next
    'MyPath = GetSetting("fasttype", section:="param�tres", Key:="LastFolder")
    Set fld = fso.GetFolder("f:\int�rieur\" & MyPath & "\")
     
     s = ""
   ' For Each fld In subflds
        For Each fl In fld.Files
           ' MyIndex = MyIndex + 1
            'MyLen = Len(fl.Name)
               ' If Right(fl.Name, 3) <> "tmp" Then
            '''''''''''''
            If Right(fl.Name, 3) <> "tmp" Then
                    MyInternalIndex = MyInternalIndex + 1
            
            
                    MyFolders.fichiers.AddItem fl.Name
                     MyFolders.fichiers.List(MyInternalIndex - 1, 1) = fl.DateLastModified
                     
                    MyFolders.fichiers.List(MyInternalIndex - 1, 2) = Replace(fl.ParentFolder, "f:\int�rieur\", "")
                
                End If
            
            
            ''''''''''''''
        
            
                   '  MyFolders.fichiers.AddItem fl.Name
            
                    'MyFolders.fichiers.List(MyIndex - 1, 1) = Replace(fl.ParentFolder, "f:\int�rieur\", "")
                'End If
        Next
End Sub

Sub lancer_recherche_mdb()


lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"
lecture_mdb "em", "que", "ph", Left("em", 1)
lecture_mdb "a", "if", "b", "a"



MyInputBox.Show


End Sub



Sub lecture_mdb(D�but, fin, Milieu, MyTable)


    Dim docNew As Document
    'Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
    Dim intRecords 'As Integer
    Dim Filter, TailleD�butFile, TailleFinFile, AvantMiddle, Apr�sMiddle
    
    

   
    
    
    Filter = D�but & "*" & Milieu & "*" & fin
    Dim strsql
    strsql = "SELECT " & MyTable & ".forme FROM " & MyTable & " WHERE (((" & MyTable & ".forme) Like """ & Filter & """));"
    

    'Set docNew = Documents.Add
    Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
    
    
    Set rdshippers = dbNorthwind.OpenRecordset(strsql)
    
    
    
    
   
    If rdshippers.BOF = True Then GoTo fin
    
    
    
    
 
   rdshippers.MoveFirst
   While rdshippers.EOF = False
   
   
   ' Set rdShippers = dbNorthwind.OpenRecordset(Name:="les_mots_sans_les_conjugaisons")
   ' For intRecords = 0 To 153
   
       ' If rdShippers.Fields(0).Value Like "*" & "" & Milieu & "" & "*" Then
        
   
    
     '   If rdShippers.Fields(0).Value Like filter Then
       MyInputBox.suggestions.AddItem rdshippers.Fields(0).Value
        
        
       ' End If
        
      '  End If
       rdshippers.MoveNext
    
   Wend
    

    
    
fin:
    
    
    
    
  ' rdShippers.Close
   'dbNorthwind.Close
   





End Sub

'Sub ajout_entr�e_ab(mynom, MyValue, cr�ation_auto As Boolean)
''a priori, fonction qui ne doit plus �tre utilis�e apr�s 5 f�vrier 2012
'
'    Dim docNew As Document
'   ' Dim dbNorthwind As DAO.Database
'    Dim rdshippers As Recordset
'    Dim intRecords 'As Integer
'    Dim i
'
'
'
'
'
'
'   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
'
'
'    Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")
'
'
'   supprimer_ab_dans_bdd (mynom)
'
'
'    With rdshippers
'   .AddNew
'   !nom = Trim(mynom)
'    !valeur = Trim(MyValue)
'    !cr�ation_auto = cr�ation_auto
'    .Update
'
'
'
'    End With
'
'
'
'
'  'rdShippers.Close
'  'dbNorthwind.Close
'
'
'
'
'
'
'
'
'
'
'
'End Sub


Sub supprimer_entr�e_ab_dans_base(MyValue)

 Dim docNew As Document
   ' Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
    Dim intRecords 'As Integer
    Dim i
    
Set dbNorthwind = OpenDatabase(get_hd & ":\fasttype\mots_reverses.mdb")

 Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")

With rdshippers
      Do While Not .EOF
        If rdshippers.Fields(2) = MyValue Then
        .Delete
        Exit Sub
        End If
         .MoveNext
      Loop
   End With





End Sub




Sub remplacement_lien()
Attribute remplacement_lien.VB_Description = "Macro enregistr�e le 30/12/2010 par Emmanuel"
Attribute remplacement_lien.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.remplacement_lien"
 On Error GoTo erreur:
Dim sR�cup, pdc As Integer, pdr As Integer
Dim Schemin, sR�fs, sTitre As Variant, sPr�R�f�rence As Variant
Dim pnX As Integer, pnY As Integer, sTaille As Integer, i, MyLink

Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject

'MyDataObject.GetFromClipboard

'sR�cup = MyDataObject.GetText


Dim MyTarget As Document




For i = 1 To Application.Selection.Hyperlinks.Count

MyLink = Application.Selection.Hyperlinks(i).Address
'signets.signet.AddItem MyTarget.Bookmarks(i).Name

Application.Selection.Hyperlinks(i).Address = Replace(MyLink, "..", "http://www.ue.espacejudiciaire.net/docs")


Next i



'Load signets
'signets.Show




        
        
erreur:
If err = 4160 Then

sMessage "Vous n'avez pas s�lectionn� de fichier cible. Collecter le nom dans le fichier cible", "ok", "ok", "ok", "ok", "Pas de fichier", "bleu", 1
Exit Sub


End If
'
End Sub
Sub Essai_insertion_ligne_tableau()
Attribute Essai_insertion_ligne_tableau.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute Essai_insertion_ligne_tableau.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Essai_insertion_ligne_tableau"
'
' Essai_insertion_ligne_tableau Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.TypeText Text:="lklkj"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="lkjlkj"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="lkjlkj"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="lkjlkj"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="lkjlkj"
End Sub
Sub test_s�lection_cellule()
Attribute test_s�lection_cellule.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute test_s�lection_cellule.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.test_s�lection_cellule"
'
' test_s�lection_cellule Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectCell
    Selection.MoveRight Unit:=wdCell
End Sub
Sub Macro3()
Attribute Macro3.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCell
    Selection.MoveLeft Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectCell
End Sub
Sub macro4()
Attribute macro4.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.macro4"
'
' macro4 Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.Tables(1).Select
    Selection.Sort ExcludeHeader:=False, FieldNumber:="Colonne 1", _
        SortFieldType:=wdSortFieldDate, SortOrder:=wdSortOrderDescending, _
        FieldNumber2:="", SortFieldType2:=wdSortFieldAlphanumeric, SortOrder2:= _
        wdSortOrderAscending, FieldNumber3:="", SortFieldType3:= _
        wdSortFieldAlphanumeric, SortOrder3:=wdSortOrderAscending, Separator:= _
        wdSortSeparateByCommas, SortColumn:=False, CaseSensitive:=False, _
        LanguageID:=wdFrench, SubFieldNumber:="Paragraphes", SubFieldNumber2:= _
        "Paragraphes", SubFieldNumber3:="Paragraphes"
    ActiveDocument.Save
End Sub
Sub Macro5()
Attribute Macro5.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.Rows.Delete
    Selection.EndKey Unit:=wdStory
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCell
End Sub
Sub macro7()
Attribute macro7.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute macro7.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.macro7"
'
' macro7 Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "??/??/????"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        
    End With
    Selection.Find.Execute
    
    While Selection.Find.Found
      
     MsgBox Selection
     
 
  
     Selection.Next
    
    
    
   
    Wend
    
End Sub
Sub Macro6()
Attribute Macro6.VB_Description = "Macro enregistr�e le 08/05/2011 par SGA-EB"
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro6"
'
' Macro6 Macro
' Macro enregistr�e le 08/05/2011 par SGA-EB
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    With Selection.Find
        .Text = "??/??/????"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub fond_de_page()
Attribute fond_de_page.VB_Description = "Macro enregistr�e le 21/05/2011 par SGA-EB"
Attribute fond_de_page.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.fond_de_page"
'
' fond_de_page Macro
' Macro enregistr�e le 21/05/2011 par SGA-EB
Dim PowerPlusWaterMarkObject1, MyNumber, i

fond_de_pages.MyListe.AddItem "confidentiel"
fond_de_pages.MyListe.AddItem "brouillon"
fond_de_pages.MyListe.AddItem Date
fond_de_pages.MyListe.AddItem "version 1"
fond_de_pages.MyListe.AddItem "version 2"
fond_de_pages.MyListe.AddItem "version 3"
fond_de_pages.MyListe.AddItem "version 4"
fond_de_pages.MyListe.AddItem "version 5"
fond_de_pages.MyListe.AddItem "int�rieur"
fond_de_pages.MyListe.AddItem "version finale"
fond_de_pages.MyListe.AddItem "version provisoire"

ActiveDocument.Sections(1).Range.Select
ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
MyNumber = Selection.HeaderFooter.Shapes.Count

For i = 1 To MyNumber
If Selection.HeaderFooter.Shapes(i).Name = "PowerPlusWaterMarkObject1" Then

   fond_de_pages.bouton_supprimer.Enabled = True
fond_de_pages.MyListe = Selection.HeaderFooter.Shapes(i).TextEffect.Text

End If





Next i


fond_de_pages.Show
Select Case MyPbkMsg



Case 1
    
    
     ActiveDocument.Sections(1).Range.Select
ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
MyNumber = Selection.HeaderFooter.Shapes.Count
For i = 1 To MyNumber
If Selection.HeaderFooter.Shapes(i).Name = "PowerPlusWaterMarkObject1" Then

    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.Shapes("PowerPlusWaterMarkObject1").Select
    Selection.Delete
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument


End If





Next i
    
    
    
    





    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.Shapes.AddTextEffect(PowerPlusWaterMarkObject1, _
        MySaisie, "Times New Roman", 1, False, False, 0, 0).Select
    Selection.ShapeRange.Name = "PowerPlusWaterMarkObject1"
    Selection.ShapeRange.TextEffect.NormalizedHeight = False
    Selection.ShapeRange.Line.Visible = False
    Selection.ShapeRange.Fill.Visible = True
    Selection.ShapeRange.Fill.Solid
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 255)
    Selection.ShapeRange.Fill.Transparency = 0.4
    Selection.ShapeRange.Rotation = 315
    Selection.ShapeRange.LockAspectRatio = True
    Selection.ShapeRange.Height = CentimetersToPoints(4.1)
    Selection.ShapeRange.Width = CentimetersToPoints(18.46)
    Selection.ShapeRange.WrapFormat.AllowOverlap = True
    Selection.ShapeRange.WrapFormat.Side = wdWrapNone
    Selection.ShapeRange.WrapFormat.Type = 3
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Left = wdShapeCenter
    Selection.ShapeRange.Top = wdShapeCenter
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
Case 2

    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.Shapes("PowerPlusWaterMarkObject1").Select
    Selection.Delete
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Select

        Selection.MoveRight Unit:=wdCharacter
    
End Sub

Sub vider_table(MyTable)



    'Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset

'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
Set dbNorthwind = OpenDatabase(get_hd & "\fasttype\mots_reverses.mdb")

 Set rdshippers = dbNorthwind.OpenRecordset(MyTable)



'on efface les donn�es de la table
With rdshippers
      Do While Not .EOF
        'If rdShippers.Fields(2) = MyValue Then
        .Delete
       
        
         .MoveNext
      Loop
   End With

 'on remplit la table avec les valeurs du r�pertoire





End Sub




Public Sub maj_folder(MyFolderToUpdate, MyOrder, MySource)
 Dim fso, folder, subflds, fld, s, fl As file, MyIndex, MyLen, MyInternalIndex, i
 ' Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
        
Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")

 Set rdshippers = dbNorthwind.OpenRecordset("documents")
    
    
    
    
    
    
    vider_table "documents"
    
    Select Case MySource
    
    Case "tous"
    
    Set fso = CreateObject("Scripting.FileSystemObject")

'totototo
    Set folder = fso.GetFolder("f:\int�rieur\" & MyFolderToUpdate & "\")
    'fl = fso.GetFileName(folder)
    For Each fl In folder.Files
    MyLen = Len(fl.Name)
                If Right(fl.Name, 3) <> "tmp" Then
    
    
                
    
                    With rdshippers
                          .AddNew
                          !Date = fl.DateLastModified
                          !nom = fl.Name
                          !dossier = Replace(fl.ParentFolder, "f:\int�rieur\", "")
                          .Update
                         
                       End With



    
    
    
    
  
                End If
    Next
    rdshippers.Close
    
    Case "seulement ceux affich�s"
   
    Dim MyNumber
    MyNumber = MyFolders.fichiers.ListCount
  
   
    Set rdshippers = dbNorthwind.OpenRecordset("documents")
    For i = 1 To MyNumber
   
   
   'MyInputBox.suggestions.List(h - 1, 1)
   
     With rdshippers
                          .AddNew
                          !Date = MyFolders.fichiers.List(i - 1, 1)
                          !nom = MyFolders.fichiers.List(i - 1, 0)
                          !dossier = MyFolders.fichiers.List(i - 1, 2)
                          .Update
                         
                       End With

     Next
    End Select
   ' SELECT documents.compteur, documents.date, documents.nom
'FROM Documents
'ORDER BY documents.date DESC;

     Dim strsql
    
    ''''''''''''''''''''''''''''
    Select Case MyOrder
    
    Case "date_descendante"
    
    
    
   
    strsql = "SELECT documents.compteur,documents.date,documents.nom,documents.dossier FROM documents ORDER BY documents.date DESC;"
    
    Case "date_ascendante"
      strsql = "SELECT documents.compteur,documents.date,documents.nom,documents.dossier FROM documents ORDER BY documents.date ;"
    
    Case "nom_ascendant"
      strsql = "SELECT documents.compteur,documents.date,documents.nom ,documents.dossier FROM documents ORDER BY documents.NOM ;"
    Case "nom_descendant"
      strsql = "SELECT documents.compteur,documents.date,documents.nom ,documents.dossier FROM documents ORDER BY documents.nom DESC;"
    
    End Select
    
     Set rdshippers = dbNorthwind.OpenRecordset(strsql)
 
MyFolders.fichiers.Clear
    With rdshippers
      Do While Not .EOF
       MyIndex = MyIndex + 1
          If IsNull(rdshippers.Fields(2)) = False And IsNull(rdshippers.Fields(1)) = False Then
          MyFolders.fichiers.AddItem rdshippers.Fields(2)
        MyFolders.fichiers.List(MyIndex - 1, 1) = rdshippers.Fields(1)
         MyFolders.fichiers.List(MyIndex - 1, 2) = rdshippers.Fields(3)
            End If
        
         .MoveNext
      Loop
   End With
    
    
    
     maj_nombre_dossiers_fichiers
End Sub




Public Sub ajoute_toutes_entr�es_bdd()
On Error GoTo wrong

    Dim docNew As Document
    'Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
    Dim intRecords 'As Integer
    Dim i, MyControls
    
    

   
    
    
   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
    
    
    Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")
    
     MyControls = AutoCorrect.Entries.Count
   
   For i = 1 To MyControls - 1

    'ts.WriteLine AutoCorrect.Entries(i).Name & " ==== " & AutoCorrect.Entries(i).Value ''' ligne � remettre pour fichier
    
        'ajout_entr�e_ab AutoCorrect.Entries(i).Name, AutoCorrect.Entries(i).Value

   'le fait que l'index "nom" soit unique fait qu'aucun doublon n'est possible et surtout, �vite la perte de la date d'enregistrement de l'abr�viation
   

     
    With rdshippers
   .AddNew
   !nom = AutoCorrect.Entries(i).Name
    !valeur = AutoCorrect.Entries(i).Value
    .Update

    
    
    End With
    
     Next i

MsgBox i & " abr�viations sauvegard�es dans la base de donn�es", vbInformation, "Sauvegarde"
    
  'rdShippers.Close
  'dbNorthwind.Close
   
wrong:
If err = 3022 Then Resume Next









End Sub
Public Sub supprimer_ab_dans_bdd(MyName, MyValue)


    Dim rdshippers As Recordset
   
    Dim strsql, jamais_dans_registre, registre
  
    
   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
    
strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.r�f�rence, abr�viations.registre, abr�viations.jamais_dans_registre FROM abr�viations WHERE (((abr�viations.nom)=""" & MyName & """) AND ((abr�viations.valeur)=""" & MyValue & """));"

   
    Set rdshippers = dbNorthwind.OpenRecordset(strsql)
    
  

 'premi�re phase, on supprime de la table abr�viation toutes les abr�viations qui ont la m�me r�f�rence
 'on extrait donc de la table abr�viations la r�f�rence (id)
 'il faut aussi extraire la valeur jamais_dans_registre : si elle est fausse, alors il faut d'abord supprimer la valeur du
 'registre


     
    With rdshippers
        .MoveFirst
            MyId = rdshippers.Fields("r�f�rence")
            jamais_dans_registre = rdshippers.Fields("jamais_dans_registre")
            registre = rdshippers.Fields("registre")
    End With
            
 strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.r�f�rence FROM abr�viations WHERE (((abr�viations.r�f�rence)=" & MyId & "));"
      Set rdshippers = dbNorthwind.OpenRecordset(strsql)
            
      With rdshippers
      .MoveFirst
             While rdshippers.EOF = False
                    If jamais_dans_registre = False And registre = True Then
           
                 If abr�viation_existe(rdshippers.Fields("nom"), rdshippers.Fields("valeur")) = 0 Then
           
                  AutoCorrect.Entries(rdshippers.Fields("nom")).Delete
                 End If
               
               
                End If
                  .MoveNext
             Wend
      End With
    
 'il faut maintenant effacer les enregistrements dans la table abr�viation
 
    With rdshippers
    .MoveFirst
    While rdshippers.EOF = False
    .Delete
    .MoveNext
    Wend
    End With
 

 
 'deuxi�me phase : il faut regarder si existe encore la m�me abr�viation, et si elle n'entre pas dans les doublons.
 'dans ce cas, il faut ins�rer dans le registre toute ses formes
 
strsql = "SELECT abr�viations.nom, abr�viations.r�f�rence, abr�viations.registre FROM abr�viations WHERE (((abr�viations.nom)=""" & MyName & """) AND ((abr�viations.registre)=0));"
Set rdshippers = dbNorthwind.OpenRecordset(strsql)
If rdshippers.RecordCount = 1 Then
    With rdshippers
        .MoveFirst
         MyId = rdshippers.Fields("r�f�rence")
    End With
    
    strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.registre, abr�viations.r�f�rence FROM abr�viations WHERE (((abr�viations.r�f�rence)=" & MyId & ") AND ((abr�viations.jamais_dans_registre)=0));"
      Set rdshippers = dbNorthwind.OpenRecordset(strsql)
      With rdshippers
        .MoveFirst
            While rdshippers.EOF = False
            AutoCorrect.Entries.Add rdshippers.Fields("nom"), rdshippers.Fields("valeur")
            .Edit
            !registre = -1
            .Update
            .MoveNext
            
            Wend
    End With
End If


   
End Sub


Public Function get_settings_from_bdd(MyField_Param�tres)

'0 = compteur
'1 = date_usage
'2 = MyHeure
'3 = Lastab

 
'Dim dbNorthwind As DAO.Database

Dim MyParam�tres As Recordset
Dim LastAb





   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
      
    Set MyParam�tres = dbNorthwind.OpenRecordset("param�tres")
  
  With MyParam�tres
  .Index = "PrimaryKey"
  .Seek "=", 1
  get_settings_from_bdd = MyParam�tres.Fields(MyField_Param�tres)
  End With
End Function







Public Sub conjuguer_un_verbe_depuis_table(sVerbe, myab)
Dim mysettings, intsettings, conjugu�, orthographe, sNombre, i, Terminaison, taille_terminaison, personne, temps, sTailleVerbe, sRacineVerbe
Dim sGroupeVerbe, MySettings2, intsettings2, MySettings3, RacineMyAb, TailleMyAb, ParticipePr�sent, accord, final, intsettings3, adjectif, MySettingAccords
Dim finalAb, FinaleVerbe, m, MyName, MyTerminaison, MyValeurConjugu�e, MyValue, strsql3
TailleMyAb = Len(myab)


'Dim dbNorthwind As DAO.Database

Dim MyParam�tres As Recordset
Dim MyTemps As Recordset
Dim LastAb


'cr�ation de l'infinitif (puisqu'on part d'une forme conjugu�e
 ' AutoCorrect.Entries.Add MyAb, sVerbe
  
' ajout_entr�e_ab MyAb, sVerbe, False

stocker_abr�viations myab, sVerbe, False, False, MyId
'enseigner_abr�viations MyAb, sVerbe, False
'cr�ation des conjugaisons

   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
      
    Set MyParam�tres = dbNorthwind.OpenRecordset("table_m�re_des_verbes")
  
  With MyParam�tres
  .Index = "forme"
  .Seek "=", sVerbe
  'get_settings_from_bdd = MyParam�tres.Fields(MyField_Param�tres)
  End With

If MyParam�tres.NoMatch = False Then
            'MyInputBox.fichiers_consult�s.AddItem filename
            
  'GoTo skip_essai
  Set MyTemps = dbNorthwind.OpenRecordset("temps_combinaison")
'''''''''''''''''''''''''''''''
    With MyTemps
      Do While Not .EOF
      
        
        Select Case Len(myab)
 
                        Case 1
                        RacineMyAb = myab
                        MyName = RacineMyAb & Right(MyTemps.Fields("nom"), Len(MyTemps.Fields("nom")) - 1)
                       
                        Case Else
                        
                        RacineMyAb = Left(myab, Len(myab) - 1)
                        MyName = RacineMyAb & MyTemps.Fields("nom")
                        End Select
                        
                        
                        MyValeurConjugu�e = MyParam�tres.Fields("" & MyTemps.Fields("donn�es") & "")
                       
                        If MyValeurConjugu�e <> "" Then
                               ' MyValue = lire_fichier_verbe(sVerbe, (MySettingAccords((m), 1)))
                               
                              ' If Len(MyValeurConjugu�e) < 4 Then GoTo Stockage_sp�cial
                                If Application.CheckSpelling(MyName) = True Then GoTo Stockage_sp�cial
                                If Len(MyValeurConjugu�e) - Len(MyName) < 1 Then GoTo SkipEnregistrement
                               
                               
                               
                               
                               
                               If MyTemps.Fields("automatique") = True Then
                               
                                stocker_abr�viations MyName, MyValeurConjugu�e, True, False, MyId
                                
                                Else 'le champ "jamais_dans_registre" de la table abr�viation prend la valeur -1
                                
Stockage_sp�cial:
                                 stocker_abr�viations MyName, MyValeurConjugu�e, True, True, MyId
                                
                                
                                
                                End If
                                
                                
                                
                                
                               ' ajout_entr�e_ab : il faudrait, � cet endroit, imm�diatement sauver l'abr�viation
                               'cr�er dans la bdd
                               ' MyInputBox.suggestions.AddItem MyName
                               ' MyInputBox.suggestions.List(m, 1) = MyValeurConjugu�e
                        End If
       
SkipEnregistrement:
       

         .MoveNext
      Loop
   End With
            
  
  'participe pr�sent
  
        
                Select Case Len(myab)
 
                        Case 1
                       
                        MyName = myab & "�" ' � �crire en dynamique plus tard
                        
                        Case Else
                        
                        
                        MyName = Left(myab, Len(myab) - 1) & "�" '� �crire en dynamique plus tard
                
                End Select
                        MyValeurConjugu�e = MyParam�tres.Fields("100")
                        If MyValeurConjugu�e <> "" Then
                              
                              'AutoCorrect.Entries.Add MyName, MyValeurConjugu�e
                            stocker_abr�viations MyName, MyValeurConjugu�e, True, False, MyId
                             
                                m = m + 1
                        End If
                        
  
  
  'participe pass�
    'masculin et masculin pluriel
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'il faut revoir cela, car le r�sultat ne semble pas bon
    
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
                MyTerminaison = Right(MyParam�tres.Fields("99"), 1)
                 
                 Select Case Len(myab)
                       
                        Case 1
                       
                     
                        MyName = myab & MyTerminaison
                        
                        Case Else
                        
                        
                        MyName = Left(myab, Len(myab) - 1) & MyTerminaison
                
                End Select
                        MyValeurConjugu�e = MyParam�tres.Fields("99")
                        If MyValeurConjugu�e <> "" Then
                                ' MyValue = lire_fichier_verbe(sVerbe, 99)
                                ' AutoCorrect.Entries.Add MyName, MyValeurConjugu�e
                                    stocker_abr�viations MyName, MyValeurConjugu�e, True, False, MyId
                               '  MyInputBox.suggestions.AddItem MyName
                               '  MyInputBox.suggestions.List(m, 1) = MyValue
                                 m = m + 1
                        End If
                        
                        
                        
                        
                        MyValeurConjugu�e = MyParam�tres.Fields("99") & "s"
                        If MyValeurConjugu�e <> "" Then
                       ' MyValue = lire_fichier_verbe(sVerbe, 99) & "s"
                        'AutoCorrect.Entries.Add MyName & get_accord("singulier_pluriel"), MyValeurConjugu�e
                        stocker_abr�viations MyName & get_accord("singulier_pluriel"), MyValeurConjugu�e, True, False, MyId
                       ' MyInputBox.suggestions.AddItem MyName
                       ' MyInputBox.suggestions.List(m, 1) = MyValue
                          m = m + 1
                        End If
                        
                    
    ' f�minin et f�minin pluriel
    
                 MyTerminaison = Right(MyParam�tres.Fields("98"), 2)
                 MyTerminaison = Left(MyTerminaison, 1)
                 Select Case Len(myab)
                       
                        Case 1
                       
                     
                        MyName = myab & MyTerminaison
                        
                        Case Else
                        
                        
                        MyName = Left(myab, Len(myab) - 1) & MyTerminaison
                
                End Select
                        
                        MyValeurConjugu�e = MyParam�tres.Fields("98")
                        If MyValeurConjugu�e <> "" Then
                                ' MyValue = lire_fichier_verbe(sVerbe, 98)
                                ' AutoCorrect.Entries.Add MyName & get_accord("f�minin"), MyValeurConjugu�e
                                 stocker_abr�viations MyName & get_accord("f�minin"), MyValeurConjugu�e, True, False, MyId
                                
                                ' MyInputBox.suggestions.AddItem MyName
                                ' MyInputBox.suggestions.List(m, 1) = MyValue
                                  m = m + 1
                        End If
                        
                        
                        
                        MyValeurConjugu�e = MyParam�tres.Fields("98") & "s"
                        If MyValeurConjugu�e <> "" Then
                            ' MyValue = lire_fichier_verbe(sVerbe, 98) & "s"
                             'AutoCorrect.Entries.Add MyName & get_accord("f�minin_pluriel"), MyValeurConjugu�e
                                stocker_abr�viations MyName & get_accord("f�minin_pluriel"), MyValeurConjugu�e, True, False, MyId
                                
                             'MyInputBox.suggestions.AddItem MyName
                             'MyInputBox.suggestions.List(m, 1) = MyValue
                              m = m + 1
                        End If
    
   
    'ayant plus participe
    
  
  
'stocker_verbes_automatiques






'End If

End If 'MyParam�tres.NoMatch = False

End Sub

Public Function acc�der_verbe_dans_table(MyVerb, MyNumberConjugaison)

Dim mysettings, intsettings, conjugu�, orthographe, sNombre, i, Terminaison, taille_terminaison, personne, temps, sTailleVerbe, sRacineVerbe
Dim sGroupeVerbe, MySettings2, intsettings2, MySettings3, RacineMyAb, TailleMyAb, ParticipePr�sent, accord, final, intsettings3, adjectif, MySettingAccords
Dim finalAb, FinaleVerbe, m, MyName, MyTerminaison, MyValeurConjugu�e, MyValue
TailleMyAb = Len(myab)


'Dim dbNorthwind As DAO.Database

Dim MyParam�tres As Recordset
Dim LastAb





   Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
      
    Set MyParam�tres = dbNorthwind.OpenRecordset("table_m�re_des_verbes")
  
  With MyParam�tres
  .Index = "forme"
  .Seek "=", MyVerb
  'get_settings_from_bdd = MyParam�tres.Fields(MyField_Param�tres)
  End With

If MyParam�tres.NoMatch = False Then
            'MyInputBox.fichiers_consult�s.AddItem filename
       
      
                        
                        
                        





acc�der_verbe_dans_table = MyParam�tres.Fields(MyNumberConjugaison)
Else
acc�der_verbe_dans_table = ""
End If

  
       
End Function

Public Sub chercher_mot_ou_verbe_pour_ab_courte(myab, MyPbkMsg)

 Dim MyLookupField, MyIndex
'   Dim dbNorthwind As DAO.Database
   Dim rdshippers As Recordset
   Dim MyAbForme, MyAbF�minin, MyAbF�mininPluriel, MyAbPluriel
   Dim MyForme, MyF�minin, MyF�mininPluriel, MyPluriel, sTerminaison_�_1
   
    Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
   Set rdshippers = dbNorthwind.OpenRecordset("table_m�re")
   
'Select Case MyPbkMsg
'
'                                        Case "xxxxx"
'
'
'
'
'
'                                        Case Else
'
'                                            'Dim MyString
'                                            'MyString =
'                                            If InStr(1, MyPbkMsg, "xxx") > 1 Then
'                                              If IsVerb(MyInfinitif) Then conjuguer_un_verbe_depuis_table MyInfinitif, Left(MyAb, Len(MyAb) - Len(MyConjug) + 1)
'                                              Selection.TypeText Text:=Left(MyPbkMsg, Len(MyPbkMsg) - 18)
'
'                                            Selection.MoveRight Unit:=wdCharacter, Count:=1
'                                            Exit Sub
'                                            End
'
'
'                                            Else
                                           
                                        ' If EndIsConjug Then GoTo skip_accord_genre_et_nombre
                                           
                                        'enregistrement automatique de toutes les abr�viations pour l'accord en genre et en nombre
                                        
                                        
 If Len(myab) > 1 Then
                                        
                                        Select Case Right(Trim(myab), 1)

                                                Case get_accord("f�minin")

'                                                        MyLookupField = 1
'                                                        MyIndex = "f�minin"

                                                Case get_accord("f�minin_pluriel")
'
'                                                        MyLookupField = 2
'                                                        MyIndex = "f�minin_pluriel"

                                                Case get_accord("singulier_pluriel")

'                                                        MyLookupField = 3
'                                                        MyIndex = "pluriel"

                                                Case Else


                                                        MyLookupField = 0
                                                        MyIndex = "forme"


                                        End Select

   End If 'len(myab) > 1 then
  
                                  With rdshippers
                                  .Index = "forme"
                                  .Seek "=", MyPbkMsg
                                    
                                  
                                 
                            
                                  
                                  End With
    If rdshippers.NoMatch = True Then GoTo NoMatchTrue
    
    
    If IsVerb(MyPbkMsg) = False Then
                   ' Select Case MyLookupField
                                    
                                    '    Case 0 'forme au singulier
                                        
                                        MyAbF�minin = myab & get_accord("f�minin")
                                        MyF�minin = rdshippers.Fields(2)
                                        If IsNull(MyF�minin) = False Then
                                        AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
                                     
                                        stocker_abr�viations MyAbF�minin, MyF�minin, True, False, MyId
                                        End If
                                        
                                        
                                         MyAbF�mininPluriel = myab & get_accord("f�minin_pluriel")
                                        MyF�mininPluriel = rdshippers.Fields(3)
                                        If IsNull(MyF�mininPluriel) = False Then
                                        AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel, True, False, MyId
                                        End If
                                        
                                        
                                        MyAbPluriel = myab & get_accord("singulier_pluriel")
                                        MyPluriel = rdshippers.Fields(4)
                                        If IsNull(MyPluriel) = False Then
                                        AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
                                        stocker_abr�viations MyAbPluriel, MyPluriel, True, False, MyId
                                        End If
                                        
                                        
                                        
                                        
'                                        Case 1 'f�minin
'
'                                        MyAbForme = Left(MyAb, Len(MyAb) - 1)
'                                        MyForme = rdShippers.Fields(0)
'
'                                        AutoCorrect.Entries.Add MyAbForme, MyForme
'                                        stocker_abr�viations MyAbForme, MyForme
'
'
'
'                                        MyAbF�mininPluriel = Left(MyAb, Len(MyAb) - 1) & get_accord("f�minin_pluriel")
'                                        MyF�mininPluriel = rdShippers.Fields(2)
'                                        If IsNull(MyF�mininPluriel) = False Then
'                                        AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
'                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel
'                                        End If
'
'
'                                        MyAbPluriel = Left(MyAb, Len(MyAb) - 1) & get_accord("singulier_pluriel")
'                                        MyPluriel = rdShippers.Fields(3)
'                                        If IsNull(MyPluriel) = False Then
'                                        AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
'                                        stocker_abr�viations MyAbPluriel, MyPluriel
'                                        End If
'
'
'
'
'                                        Case 2 'f�minin pluriel
'
'                                        MyAbForme = Left(MyAb, Len(MyAb) - 1)
'                                        MyForme = rdShippers.Fields(0)
'                                        AutoCorrect.Entries.Add MyAbForme, MyForme
'                                        stocker_abr�viations MyAbForme, MyForme
'
'                                       MyAbF�minin = Left(MyAb, Len(MyAb) - 1) & get_accord("f�minin")
'                                        MyF�minin = rdShippers.Fields(1)
'                                        If IsNull(MyF�minin) = False Then
'                                        AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
'                                        stocker_abr�viations MyAbF�minin, MyF�minin
'                                        End If
'
'
'                                        MyAbPluriel = Left(MyAb, Len(MyAb) - 1) & get_accord("singulier_pluriel")
'                                        MyPluriel = rdShippers.Fields(3)
'                                        If IsNull(MyPluriel) = False Then
'                                        AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
'                                        stocker_abr�viations MyAbPluriel, MyPluriel
'                                        End If
'
'                                        Case 3 'pluriel
'
'                                        MyAbForme = Left(MyAb, Len(MyAb) - 1)
'                                        MyForme = rdShippers.Fields(0)
'                                        AutoCorrect.Entries.Add MyAbForme, MyForme
'                                        stocker_abr�viations MyAbForme, MyForme
'
'
'
'                                         MyAbF�minin = Left(MyAb, Len(MyAb) - 1) & get_accord("f�minin")
'                                        MyF�minin = rdShippers.Fields(1)
'                                        If IsNull(MyF�minin) = False Then
'                                        AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
'                                        stocker_abr�viations MyAbF�minin, MyF�minin
'                                        End If
'
'
'                                        MyAbF�mininPluriel = Left(MyAb, Len(MyAb) - 1) & get_accord("f�minin_pluriel")
'                                        MyF�mininPluriel = rdShippers.Fields(2)
'                                        If IsNull(MyF�mininPluriel) = False Then
'                                        AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
'                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel
'                                        End If
'
'
'
'
'
'                                    End Select
skip_accord_genre_et_nombre:
                                        AutoCorrect.Entries.Add myab, MyPbkMsg
                                         stocker_abr�viations myab, MyPbkMsg, True, False, MyId
                                         
                                         
                                      Else ' IsVerb(MyPbkMsg) = False
                                        
                                       
                                       
                                       
                                        
                                        
                                               ' If sTerminaison_�_1 = "r" Or sTerminaison_�_1 = "e" Then
                                      conjuguer_un_verbe_depuis_table MyPbkMsg, myab
                                                    
                                                
                                                     
                                                
                                                
                                  End If 'IsVerb(MyPbkMsg) = False

                                       
                                       ' End If
                                        
                                            
                                   ' End Select 'MyPbkMsg




 

NoMatchTrue:





End Sub



Public Function fonction_comparer_mot_et_abr�viation(MyMot, myab)


Dim MySettingAccords, i, MyStr, j, MyLen, MyAbSubstitution, MyMotDeFonction, MyPositionStr
MyAbSubstitution = myab
MyMotDeFonction = MyMot

If EndIsAccord Then MyAbSubstitution = Left(MyAbSubstitution, Len(MyAbSubstitution) - 1)

MySettingAccords = GetAllSettings(appname:="fasttype", section:="voyelles") '

         For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)


            MyMotDeFonction = Replace(MyMotDeFonction, (MySettingAccords((i), 0)), "")

         Next i

''''''''''''''''''''''IL FAUT VIRER LES DOUBLES CONSONNES ������������������������������������

MyLen = Len(MyMotDeFonction)

'il faut qu'un doublement de consonne ne soit pas pris en compte
Dim doublement()
ReDim doublement(MyLen)


For i = 1 To MyLen

    MyStr = Left(MyMotDeFonction, 1)
    doublement(i) = MyStr
       
    MyMotDeFonction = Right(MyMotDeFonction, Len(MyMotDeFonction) - 1)
    
 

If InStr(i, MyAbSubstitution, MyStr) > 0 Then


         MySameConsonnes = MySameConsonnes + 1


Else

   If i > 1 Then
       If MyStr <> doublement(i - 1) Then
        
            fonction_comparer_mot_et_abr�viation = fonction_comparer_mot_et_abr�viation + 1
            
       End If
        

   End If
    
End If


Next


End Function
Public Function fonction_d�tecter_ponctuation(myab)

Dim MyLetter, MySettingAccords, i, MyLetter2
'MyPosition = ""
'MyAb = InputBox("abr�viation")
MyLetter = Right(myab, 1)
'MyLetter2 = Mid(MyAb, 3, 1)
'
MySettingAccords = GetAllSettings(appname:="fasttype", section:="nombre_mots_firstLetters") '

            For i = LBound(MySettingAccords, 1) To UBound(MySettingAccords, 1)


                        If MyLetter = (MySettingAccords((i), 0)) Then   'lettre de d�but

                          fonction_d�tecter_ponctuation = False
'                            MsgBox fonction_d�tecter_ponctuation
                          Exit Function
                        End If

            Next i

fonction_d�tecter_ponctuation = True
MyPonctuation = MyLetter


End Function

Public Sub mettre_�_jour_table()
Dim strsql, MyTable
Dim MyLookupField, MyIndex
 '  Dim dbNorthwind As DAO.Database
   Dim rdshippers As Recordset
   Dim MyAbForme, MyAbF�minin, MyAbF�mininPluriel, MyAbPluriel
   Dim MyForme, MyF�minin, MyF�mininPluriel, MyPluriel

    Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")


dbNorthwind.Execute "supprimer_mots_avec_f�minin_et_f�minin_pluriel_sans_z"
dbNorthwind.Execute "supprimer_mots_avec_f�minin_et_f�minin_pluriel_avec_z"
dbNorthwind.Execute "supprimer_mots_avec_pluriel_avec_z"
dbNorthwind.Execute "supprimer_mots_avec_pluriel_sans_z"
dbNorthwind.Execute "supprimer_seulement_les_mots_Z"
dbNorthwind.Execute "supprimer_tout_sans_les_z"

dbNorthwind.Execute "remplir_mots_avec_f�minin_et_f�minin_pluriel_sans_z"
dbNorthwind.Execute "remplir_mots_avec_f�minin_et_f�minin_pluriel_avec_z"
dbNorthwind.Execute "remplir_mots_avec_pluriel_avec_z"
dbNorthwind.Execute "remplir_mots_avec_pluriel_sans_z"
dbNorthwind.Execute "remplir_seulement_les_mots_Z"
dbNorthwind.Execute "remplir_tout_sans_les_z"
'

MsgBox "termin�"
End Sub

Public Function IsZ(MyValue)
  If MyValue Like "*asa*" Or MyValue Like "*ase*" Or MyValue Like "*asi*" Or MyValue Like "*aso*" Or MyValue Like "*asu*" _
                Or MyValue Like "*as�*" Or MyValue Like "*as�*" Or MyValue Like "*as�*" _
                Or MyValue Like "*esa*" Or MyValue Like "*ese*" Or MyValue Like "*esi*" Or MyValue Like "*eso*" Or MyValue Like "*esu*" _
                Or MyValue Like "*isa*" Or MyValue Like "*ise*" Or MyValue Like "*isi*" Or MyValue Like "*iso*" Or MyValue Like "*isu*" _
                Or MyValue Like "*is�*" Or MyValue Like "*is�*" Or MyValue Like "*is�*" _
                Or MyValue Like "*osa*" Or MyValue Like "*ose*" Or MyValue Like "*osi*" Or MyValue Like "*oso*" Or MyValue Like "*osu*" _
                Or MyValue Like "*osy*" Or MyValue Like "*os�*" Or MyValue Like "*os�*" Or MyValue Like "*os�*" _
                Or MyValue Like "*ysa*" Or MyValue Like "*yse*" Or MyValue Like "*ysi*" Or MyValue Like "*yso*" Or MyValue Like "*ysu*" _
                Or MyValue Like "*ys�*" Or MyValue Like "*ys�*" Or MyValue Like "*ys�*" _
                Or MyValue Like "*usa*" Or MyValue Like "*use*" Or MyValue Like "*usi*" Or MyValue Like "*uso*" Or MyValue Like "*usu*" _
                Or MyValue Like "*us�*" Or MyValue Like "*us�*" Or MyValue Like "*us�*" _
                Or MyValue Like "*�sa*" Or MyValue Like "*�se*" Or MyValue Like "*�si*" Or MyValue Like "*�so*" Or MyValue Like "*�sy*" _
                Or MyValue Like "*�sy*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" _
                Or MyValue Like "*�sa*" Or MyValue Like "*�se*" Or MyValue Like "*�si*" Or MyValue Like "*�so*" Or MyValue Like "*�sy*" _
                Or MyValue Like "*�sy*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*�s�*" Or MyValue Like "*z*?" _
                Then
                
          IsZ = True
        Else
        
            IsZ = False
        End If
                
End Function
Sub Macro8()
Attribute Macro8.VB_Description = "Macro enregistr�e le 01/09/2011 par SGA-EB"
Attribute Macro8.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro8"
'
' Macro8 Macro
' Macro enregistr�e le 01/09/2011 par SGA-EB
'
    ActiveWindow.Panes(1).Activate
    Selection.MoveLeft Unit:=wdCharacter, Count:=13, Extend:=wdExtend
    Selection.Copy
    ActiveWindow.Panes(2).Activate
    Selection.PasteAndFormat (wdPasteDefault)
End Sub

Public Sub enseigner_abr�viations(myab, MyValeur)
   Dim MyLookupField, MyIndex
   'Dim dbNorthwind As DAO.Database
   Dim rdshippers As Recordset
   Dim MyAbForme, MyAbF�minin, MyAbF�mininPluriel, MyAbPluriel
   Dim MyForme, MyF�minin, MyF�mininPluriel, MyPluriel
   
   
   'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
    Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
    Set rdshippers = dbNorthwind.OpenRecordset("table_m�re")
   
'

Select Case MyPbkMsg
                                        
                                        Case "xxxxx"
                                        
                                        
                                       
                                        
                                        
                                        Case Else
                                         
                                            'Dim MyString
                                            'MyString =
                                            If InStr(1, MyPbkMsg, "xxx") > 1 Then
                                              If IsVerb(MyInfinitif) Then conjuguer_un_verbe_depuis_table MyInfinitif, Left(myab, Len(myab) - Len(MyConjug) + 1)
                                              Selection.TypeText Text:=Left(MyPbkMsg, Len(MyPbkMsg) - 18)
        
                                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                                            Exit Sub
                                            End
                                            
                                            
                                            Else 'InStr(1, MyPbkMsg, "xxx") > 1
                                           
                                         If EndIsConjug Then GoTo pas_de_recherche
                                         If Len(myab) = 1 Then 'cas o� l'on arrive du formulaire d'apprentissage des lettres seules
                                         
                                                        MyLookupField = 0
                                                        MyIndex = "forme"
                                                        
                                         GoTo skip_because_SingleLetter
                                         
                                         End If
                                         
                                           
                                        'enregistrement automatique de toutes les abr�viations pour l'accord en genre et en nombre
essaye_encore:
                                        Select Case Right(Trim(myab), 1)
                                        
                                                Case get_accord("f�minin")
                                                    
                                                        MyLookupField = 1
                                                        MyIndex = "f�minin"
                                                    
                                                Case get_accord("f�minin_pluriel")
                                                        
                                                        MyLookupField = 2
                                                        MyIndex = "f�minin_pluriel"
                                            
                                                Case get_accord("singulier_pluriel")
                                            
                                                        MyLookupField = 3
                                                        MyIndex = "pluriel"
                                                        
                                                Case Else
                                                    
                                                    
                                                        MyLookupField = 0
                                                        MyIndex = "forme"
                                                    
                                                    
                                        End Select 'Right(Trim(MyAb), 1)
skip_because_SingleLetter:
                
  
                                  With rdshippers
                                  .Index = MyIndex
                                  .Seek "=", MyPbkMsg
                                  End With
                 
If rdshippers.NoMatch = True Then
     
'         If GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = "vrai" Then
'
'                 AddMot.forme = MyPbkMsg
'                 AddMot.f�minin = MyPbkMsg
'                 AddMot.f�minin_pluriel = MyPbkMsg
'                 AddMot.singulier_pluriel = MyPbkMsg
'                 AddMot.Caption = "toutes les formes de : " & MyAb
'                     If IsZ(MyPbkMsg) = True Then
'                             AddMot.z = True
'                             AddMot.z.Caption = "� contr�ler"
'                     Else
'                             AddMot.z = False
'                     End If
'                         AddMot.Show
'                         If ResultAddMot = 9999 Then Exit Sub
'
'
'                 GoTo essaye_encore 'grace aux indications fournies dans le form addmot, on peut ensuite cr�er les abr�viations accord�es en genre et en nombre
'                 'pour cela, il faut repartir en arri�re
'
'             End If 'GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = -1
        
        
        
     Else ' rdShippers.NoMatch = True
                 
                    Select Case MyLookupField
                                    
                                        Case 0 'forme au singulier
                                        
                                        MyAbF�minin = myab & get_accord("f�minin")
                                        MyF�minin = rdshippers.Fields("f�minin")
                                        If IsNull(MyF�minin) = False And MyF�minin <> "" Then
                                       ' AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
                                        stocker_abr�viations MyAbF�minin, MyF�minin, True, False, MyId
                                        End If
                                        
                                        
                                         MyAbF�mininPluriel = myab & get_accord("f�minin_pluriel")
                                        MyF�mininPluriel = rdshippers.Fields("f�minin_pluriel")
                                        If IsNull(MyF�mininPluriel) = False And MyF�mininPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel, True, False, MyId
                                        End If
                                        
                                        
                                        MyAbPluriel = myab & get_accord("singulier_pluriel")
                                        MyPluriel = rdshippers.Fields("pluriel")
                                        If IsNull(MyPluriel) = False And MyPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
                                        stocker_abr�viations MyAbPluriel, MyPluriel, True, False, MyId
                                        End If
                                        
                                        
                                        
                                        
                                        Case 1 'f�minin
                                        
                                        MyAbForme = Left(myab, Len(myab) - 1)
                                        MyForme = rdshippers.Fields("forme")
                                       
                                        'AutoCorrect.Entries.Add MyAbForme, MyForme
                                        stocker_abr�viations MyAbForme, MyForme, True, False, MyId
                                        
                                        
                                        
                                        MyAbF�mininPluriel = Left(myab, Len(myab) - 1) & get_accord("f�minin_pluriel")
                                        MyF�mininPluriel = rdshippers.Fields("f�minin_pluriel")
                                        If IsNull(MyF�mininPluriel) = False And MyF�mininPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel, True, False, MyId
                                        End If
                                        
                                        
                                        MyAbPluriel = Left(myab, Len(myab) - 1) & get_accord("singulier_pluriel")
                                        MyPluriel = rdshippers.Fields("pluriel")
                                        If IsNull(MyPluriel) = False And MyPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
                                        stocker_abr�viations MyAbPluriel, MyPluriel, True, False, MyId
                                        
                                        End If
                                        
                                        
                                        
                                        
                                        Case 2 'f�minin pluriel
                                        
                                        MyAbForme = Left(myab, Len(myab) - 1)
                                        MyForme = rdshippers.Fields("forme")
                                        'AutoCorrect.Entries.Add MyAbForme, MyForme
                                        stocker_abr�viations MyAbForme, MyForme, True, False, MyId
                                        
                                       MyAbF�minin = Left(myab, Len(myab) - 1) & get_accord("f�minin")
                                        MyF�minin = rdshippers.Fields("f�minin")
                                        If IsNull(MyF�minin) = False And MyF�minin <> "" Then
                                        'AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
                                        stocker_abr�viations MyAbF�minin, MyF�minin, True, False, MyId
                                        End If
                                        
                                        
                                        MyAbPluriel = Left(myab, Len(myab) - 1) & get_accord("singulier_pluriel")
                                        MyPluriel = rdshippers.Fields("pluriel")
                                        If IsNull(MyPluriel) = False And MyPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbPluriel, MyPluriel
                                        stocker_abr�viations MyAbPluriel, MyPluriel, True, False, MyId
                                        End If
                                        
                                        Case 3 'pluriel
                                   
                                        MyAbForme = Left(myab, Len(myab) - 1)
                                        MyForme = rdshippers.Fields("forme")
                                        'AutoCorrect.Entries.Add MyAbForme, MyForme
                                        stocker_abr�viations MyAbForme, MyForme, True, False, MyId
                                         
                                         
                                         
                                         MyAbF�minin = Left(myab, Len(myab) - 1) & get_accord("f�minin")
                                        MyF�minin = rdshippers.Fields("f�minin")
                                        If IsNull(MyF�minin) = False And MyF�minin <> "" Then
                                        'AutoCorrect.Entries.Add MyAbF�minin, MyF�minin
                                        stocker_abr�viations MyAbF�minin, MyF�minin, True, False, MyId
                                        End If
                                        
                                        
                                        MyAbF�mininPluriel = Left(myab, Len(myab) - 1) & get_accord("f�minin_pluriel")
                                        MyF�mininPluriel = rdshippers.Fields("f�minin_pluriel")
                                        If IsNull(MyF�mininPluriel) = False And MyF�mininPluriel <> "" Then
                                        'AutoCorrect.Entries.Add MyAbF�mininPluriel, MyF�mininPluriel
                                        stocker_abr�viations MyAbF�mininPluriel, MyF�mininPluriel, True, False, MyId
                                        End If
                                        
                                      
                                    
                                    End Select 'MyLookupField
                                    
                                    

     
     
     End If ' rdshippers.NoMatch = False
     
     
     
    'cas d'un participe pr�sent ou d'un participe pass� pour enseigner le verbe


'a) participe pr�sent : le mot se termine en "ant" : il faut encore extraire cette donn�e convenablement !!!!!!!!!!!!!!!!!!!!!!!!!!! attention !!!!!!!!!!!!!!!

Dim MyValeurAb, BeginMyValeurAb, MyLastLetterAb

If Right(MyPbkMsg, 3) = "ant" Then
    Set rdshippers = dbNorthwind.OpenRecordset("table_m�re_des_verbes")
        With rdshippers
          .Index = "participe_pr�sent"
          .Seek "=", MyPbkMsg
        End With
       
       If rdshippers.NoMatch = False Then
       
           If check_existence_nom_pour_abr�viation(rdshippers.Fields("forme").Value) = True Then GoTo pas_de_recherche
     
     
     
            If Right(rdshippers.Fields("forme").Value, 2) = "re" Or Right(rdshippers.Fields("forme").Value, 2) = "ir" Then
                myab = InputBox("indiquer l'abr�viation finissant par r pour le verbe : " & rdshippers.Fields("forme").Value, "verbe du deuxi�me ou du troisi�me groupe")
                If IsNull(myab) Then GoTo saute_participes
                BeginMyValeurAb = Left(myab, Len(myab) - 1)
                MyValeurAb = "r"
            
            Else
                
                MyLastLetterAb = Mid(myab, Len(myab))
                BeginMyValeurAb = Left(myab, Len(myab) - 1)
                MyValeurAb = Replace(myab, MyLastLetterAb, "r", Len(myab), 1)
            
            End If 'Right(rdShippers.Fields("forme").Value, 2) = "re" Or Right(rdShippers.Fields("forme").Value, 2) = "ir"
    

 
             conjuguer_un_verbe_depuis_table rdshippers.Fields("forme").Value, BeginMyValeurAb & MyValeurAb
            GoTo saute_participes
    End If 'rdShippers.NoMatch = False
    
    

End If 'Right(MyPbkMsg, 3) = "ant"

'participe pass� singulier : !!!!!!!!!!!!!!!!!!! voir le pb des verbes autre que du premier groupe

Set rdshippers = dbNorthwind.OpenRecordset("table_m�re_des_verbes")

  With rdshippers
    .Index = "participe_pass�_singulier"
    .Seek "=", MyPbkMsg
    End With

    


    If rdshippers.NoMatch = False Then
    If check_existence_nom_pour_abr�viation(rdshippers.Fields("forme").Value) = True Then GoTo pas_de_recherche
         If Right(rdshippers.Fields("forme").Value, 2) = "re" Or Right(rdshippers.Fields("forme").Value, 2) = "ir" Then
         
         '� compl�ter !!!!!!!!!!!!!!!!!!!!!!!!!!!! attention !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         
            Else
            
                BeginMyValeurAb = Left(myab, Len(myab) - 1)
                MyLastLetterAb = Mid(myab, Len(myab))
                BeginMyValeurAb = Left(myab, Len(myab) - 1)
                MyValeurAb = Replace(myab, MyLastLetterAb, "r", Len(myab), 1)
                conjuguer_un_verbe_depuis_table rdshippers.Fields("forme").Value, BeginMyValeurAb & MyValeurAb
                GoTo saute_participes

        End If 'Right(rdShippers.Fields("forme").Value, 2) = "re" Or Right(rdShippers.Fields("forme").Value, 2) = "ir"
    
    End If 'rdShippers.NoMatch = False
    
 'participe pass� f�minin !!!!!!!!!!!!!!!!!!! voir le pb des verbes autre que du premier groupe
 
 Set rdshippers = dbNorthwind.OpenRecordset("table_m�re_des_verbes")

        With rdshippers
          .Index = "participe_pass�_f�minin"
          .Seek "=", MyPbkMsg
        End With


    If rdshippers.NoMatch = False Then
             
            If check_existence_nom_pour_abr�viation(rdshippers.Fields("forme").Value) = True Then GoTo pas_de_recherche
             If Right(rdshippers.Fields("forme").Value, 2) = "re" Or Right(rdshippers.Fields("forme").Value, 2) = "ir" Then
             
                '� compl�ter
             
                Else
                    BeginMyValeurAb = Left(myab, Len(myab) - 2)
                    MyLastLetterAb = Left(Mid(myab, Len(myab) - 1), 1)
                    MyValeurAb = Left(Replace(myab, MyLastLetterAb, "r", Len(myab) - 1, 1), 1)
                    conjuguer_un_verbe_depuis_table rdshippers.Fields("forme").Value, BeginMyValeurAb & MyValeurAb
                    GoTo saute_participes
             End If 'Right(rdShippers.Fields("forme").Value, 2) = "re" Or Right(rdShippers.Fields("forme").Value, 2) = "ir"
    End If 'rdShippers.NoMatch = False
 
 
 'participe pass� pluriel : il n'existe pas dans le tableau des conjugaison. Il faut donc le v�rifier d'abord
 'dans la table_m�re, puis ensuite dans la table m�re des verbes, il faut rechercher l'infinitif
 '!!!!!!!!!!!!!!!!!!! voir le pb des verbes autre que du premier groupe
 
  Set rdshippers = dbNorthwind.OpenRecordset("table_m�re")

            With rdshippers
              .Index = "f�minin_pluriel"
              .Seek "=", MyPbkMsg
            End With
    
    If rdshippers.NoMatch = False Then 'Num�ro 1
            
            Dim MyParticpePass�Singulier
            MyParticpePass�Singulier = rdshippers.Fields("forme").Value
            
            Set rdshippers = dbNorthwind.OpenRecordset("table_m�re_des_verbes")
            
            With rdshippers
                .Index = "participe_pass�_singulier"
                .Seek "=", MyParticpePass�Singulier
            End With


                   If rdshippers.NoMatch = False Then 'num�ro 2
                       If check_existence_nom_pour_abr�viation(rdshippers.Fields("forme").Value) = True Then GoTo pas_de_recherche
                       If Right(rdshippers.Fields("forme").Value, 2) = "re" Or Right(rdshippers.Fields("forme").Value, 2) = "ir" Then
                       
                          '� compl�ter !!!!!!!!!!!!!!!!!!!!!! attention !!!!!!!!!!!!!!!!!!!!!!!!!!
                       
                           Else
                               BeginMyValeurAb = Left(myab, Len(myab) - 2)
                               MyLastLetterAb = Left(Mid(myab, Len(myab) - 1), 1)
                               MyValeurAb = Left(Replace(myab, MyLastLetterAb, "r", Len(myab) - 1, 1), 1)
                            
                       End If 'Right(rdShippers.Fields("forme").Value, 2) = "re" Or Right(rdShippers.Fields("forme").Value, 2) = "ir"
                       
                      conjuguer_un_verbe_depuis_table rdshippers.Fields("forme").Value, BeginMyValeurAb & MyValeurAb
                      GoTo saute_participes
                   
                   End If 'rdShippers.NoMatch = False 'num�ro 2
            End If 'rdShippers.NoMatch = False 'Num�ro 1
 
 
 
pas_de_recherche:
saute_participes:

'si on a un verbe pass� soit par l'infinitif trouv� pendant la recherche, soit par l'infinitif pass� dans myinputbox.suggestion 3�me colonnne
'alors, on conjugue le verbe directement

        MyNewWord = MyPbkMsg
       ' sTerminaison_�_1 = Right(Trim(MyAb), 1)
        
               
        If IsVerb(MyNewWord) Then conjuguer_un_verbe_depuis_table MyNewWord, myab
        If IsVerb(MyInfinitif) Then conjuguer_un_verbe_depuis_table MyInfinitif, Left(myab, Len(myab) - Len(MyConjug) + 1)
        
        
End If
                
        End Select 'MyPbkMsg







End Sub

Public Sub extraire_sons_et_terminaisons()

'attention, il faut changer quelques termes si l'on veut charger les terminaisons ou les sons
Dim mysettings
Dim intsettings
Dim rdshippers As Recordset
 Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Set rdshippers = dbNorthwind.OpenRecordset("m�thode_ab")
'


 mysettings = GetAllSettings(appname:="fasttype", section:="terminaisons")
    For intsettings = LBound(mysettings, 1) To UBound(mysettings, 1)
    
    
     With rdshippers
    
   .AddNew
   !valeur = mysettings(intsettings, 0)
    !Abr�viation = mysettings(intsettings, 1)
    !Terminaison = -1
    
    
    .Update

    End With
    
    
    
    
   
        
 
    Next intsettings

   



End Sub

Public Function contr�le_coh�rence_abr�viative(myab, MyWord)
Dim strsql, j, NbrMotsAvecZ
Dim rdshippers As Recordset
If EndIsConjug = -1 Then
contr�le_coh�rence_abr�viative = 0
Exit Function

End If

'  Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 strsql = "SELECT m�thode_ab.Valeur, m�thode_ab.Abr�viation FROM m�thode_ab WHERE (((m�thode_ab.terminaison)=Yes));"
 Set rdshippers = dbNorthwind.OpenRecordset(strsql)


 If rdshippers.BOF = True Then GoTo SkipTerminaisonContr�le

'For h = 0 To MyNumberVerbe - 1

 rdshippers.MoveFirst

        While rdshippers.EOF = False
   '    If rdshippers.Fields("valeur").Value = "de" Then Stop
        ' If rdshippers.Fields("valeur").Value = "d�" Then Stop
       ' Debug.Print MyWord
       'tototo
'            If rdShippers.Fields("valeur").Value = "aire" Then Stop
            ''If InStr(Len(myab) - 1, myab, Trim(rdshippers.Fields("abr�viation").Value)) = 0 Then 'la lettre repr�sentant une finale ne se trouve pas comme finale de l'abr�viation
               If Len(MyWord) - Len(Trim(rdshippers.Fields("valeur").Value)) + 1 > 0 Then '�viter une valeur n�gative qui arrive parfois
               If InStr(Len(MyWord) - Len((rdshippers.Fields("valeur").Value)) + 1, MyWord, rdshippers.Fields("valeur").Value) > 0 Then

                    For j = 0 To NombreTerminaisons  'il faut voir si la terminaison en question n'est pas contenue pas dans l'une des terminaisons possibles
                    'par exemple : "re" ne doit pas entra�ner l'exclusion de "ure" ou de "oire"
                    'on r�cup�re � cette fin les terminaisons possibles dans l'array "terminaisons"




                        If InStr(1, terminaisons(j + 1), Trim(rdshippers.Fields("valeur").Value)) > 0 Then GoTo skip333 'il faut ajouter 1 car la premi�re terminaison,

                        'dans l'array "terminaison" correspond � la lettre elle-m�me (cela sert pour prendre en compte aussi la lettre elle-m�me..



                    Next j



                NbrMotsAvecZ = MyInputBox.rejet�s.ListCount + 1
                MyInputBox.rejet�s.AddItem MyWord
                MyInputBox.rejet�s.List(NbrMotsAvecZ - 1, 1) = "terminaison < " & Trim(rdshippers.Fields("valeur").Value) & " > dans le mot mais pas dans l'abr�viation : < " & Right(myab, 1) & " > au lieu de <" & Trim(rdshippers.Fields("abr�viation").Value) & ">."
                MyInputBox.rejet�s.List(NbrMotsAvecZ - 1, 2) = Left(myab, Len(myab) - 1) & Trim(rdshippers.Fields("abr�viation").Value)
                contr�le_coh�rence_abr�viative = -1

                End If 'InStr(1, myinputbox.suggestions.List(h - 1), MySettingAccords((i), 1)) > 0
                End If 'Len(MyInputBox.suggestions.List(h)) - Len(MySettingAccords((i), 0)) + 1 > 0

            ''End If
skip333:
            rdshippers.MoveNext
        Wend
SkipTerminaisonContr�le:
     

'Next h
End Function

Sub Hyper()
Attribute Hyper.VB_Description = "Macro enregistr�e le 16/12/2011 par ebarbe-adc"
Attribute Hyper.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Hyper"
Dim sR�cup As Variant
Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject


MyDataObject.GetFromClipboard
sR�cup = MyDataObject.GetText

'
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        sR�cup, _
        SubAddress:="", ScreenTip:="", TextToDisplay:=sR�cup
End Sub







Public Function lookup_ab(Terminaison)





End Function

Public Sub ajout_entr�e_conjugaisons_rares(myab, MyValeur)




    Dim docNew As Document
   ' Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
    Dim intRecords 'As Integer
    Dim i
    
    

   
    
    
   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
    
    
    Set rdshippers = dbNorthwind.OpenRecordset("conjugaisons_rares")
    
      
   supprimer_conjugaison_rare (myab)

     
    With rdshippers
   .AddNew
   !nom = Trim(myab)
    !valeur = Trim(MyValeur)
    
    .Update

    
    
    End With
    
    
 
    
  'rdShippers.Close
  'dbNorthwind.Close
   












End Sub



Public Sub supprimer_conjugaison_rare(MyNom)
   Dim docNew As Document
   ' Dim dbNorthwind As DAO.Database
    Dim rdshippers As Recordset
    Dim intRecords 'As Integer
    Dim i, MyControls
    
    

   
    
    
   Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
    
    
    Set rdshippers = dbNorthwind.OpenRecordset("conjugaisons_rares")
    
    
   
   

    'ts.WriteLine AutoCorrect.Entries(i).Name & " ==== " & AutoCorrect.Entries(i).Value ''' ligne � remettre pour fichier
    
        'ajout_entr�e_ab AutoCorrect.Entries(i).Name, AutoCorrect.Entries(i).Value

   

     
    With rdshippers
    .Index = "nom"
    
   .Seek "=", MyNom
   
   If rdshippers.NoMatch = False Then
   
   
   .Edit
   .Delete

        End If
    
    End With
    
    


    
  'rdShippers.Close
  'dbNorthwind.Close
   
End Sub

Public Function recherche_conjugaison_rare(myab)
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")12/04/2015
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
Dim rdshippers As Recordset
Dim i, MyDoublon
MyDoublon = 0




'recherche d'un doublon
Dim strsql
strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.nom, abr�viations.Nombre_usage FROM abr�viations WHERE (((abr�viations.nom) In (SELECT [nom] FROM [abr�viations] As Tmp GROUP BY [nom] HAVING Count(*)>1 )) AND ((abr�viations.nom)=""" & myab & """)) ORDER BY abr�viations.Nombre_usage DESC"
            
             Set rdshippers = dbNorthwind.OpenRecordset(strsql)
                
                
                
    If rdshippers.RecordCount >= 1 Then
             
                rdshippers.MoveFirst
                        While rdshippers.EOF = False
                        
                      
                          choix_abr�viation.choix_ab.AddItem rdshippers.Fields("valeur")
                            recherche_conjugaison_rare = 1
                     
                        rdshippers.MoveNext
                        Wend
                        ChoixAbr�viation = choix_abr�viation.choix_ab.List(0)
       
    
            End If
'rechercher des conjugaisons rares
            
strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.nombre_usage, abr�viations.jamais_dans_registre FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """) AND ((abr�viations.jamais_dans_registre)=-1)) ORDER BY abr�viations.nombre_usage DESC;"
        
             Set rdshippers = dbNorthwind.OpenRecordset(strsql)
                 If rdshippers.RecordCount >= 1 Then
                    rdshippers.MoveFirst
                        While rdshippers.EOF = False
                            If choix_abr�viation.choix_ab.ListCount = 0 Then GoTo SkipExamenList
                    
                            For i = 0 To choix_abr�viation.choix_ab.ListCount
                                If rdshippers.Fields("valeur") = choix_abr�viation.choix_ab.List(i) Then
                                    
                                  GoTo SkipRecord
                                
                                Else
                                
                                MyDoublon = MyDoublon + 0
                                
                                
                                End If
                                
             
             

            
                            Next


                If MyDoublon = 0 Then
SkipExamenList:
                choix_abr�viation.choix_ab.AddItem rdshippers.Fields("valeur")
                recherche_conjugaison_rare = 1
                    ChoixAbr�viation = choix_abr�viation.choix_ab.List(0)
                End If
                

SkipRecord:

                rdshippers.MoveNext
                Wend
                
                
            End If
            
    
If choix_abr�viation.choix_ab.ListCount = 0 Then recherche_conjugaison_rare = 0
If choix_abr�viation.choix_ab.ListCount = 1 And rdshippers.RecordCount = 1 Then
choix_abr�viation.bouton_supprimer_jamais_dans_registre.Visible = True
choix_abr�viation.bouton_supprimer_jamais_dans_registre.Caption = "Restaurer le d�veloppement automatique de <" & choix_abr�viation.choix_ab.List(0) & ">"
Else
choix_abr�viation.bouton_supprimer_jamais_dans_registre.Visible = False
End If


'End If

End Function


Public Sub chercher_utilisation_abr�viation(myab)

Dim strsql, MyNumberRecords, i
Dim rdshippers As Recordset




''Dim MyAb1, MyAb2, MyAutoCorrects, MyExistingAb, j, k, MyValue
''
''
''
''k = 0
''MyAb1 = MyAb
''MyAb1 = MyAb1 & "*"
''MyAb2 = "*" & MyAb & "*"""
MyInputBox.zone_abr�viations_approchantes.Clear
MyInputBox.ZoneMotsCorrespondants.Clear
''
''
'' 'MyAutoCorrects = AutoCorrect.Entries.Count
''  MyExistingAb = 0
''  '      For j = 1 To MyAutoCorrects
''
''        If check_existence_valeur_pour_abr�viation(MyAb) Then  'la valeur existe d�j�
''            MyInputBox.ZoneMotsCorrespondants.Clear
''            MyValue = AutoCorrect.Entries(MyIndexAutocorrect).Value
''            MyInputBox.ZoneMotsCorrespondants.AddItem MyValue
''            MyExistingAb = MyExistingAb + 1
''
''         End If 'AutoCorrect.Entries(j).Name = Me.zone_abr�viation
''
''         If check_existence_valeur_pour_abr�viation(MyAb1) Or check_existence_valeur_pour_abr�viation(MyAb2) Then
''          k = k + 1
''         MyInputBox.zone_abr�viations_approchantes.AddItem AutoCorrect.Entries(MyIndexAutocorrect).Name
''         MyInputBox.zone_abr�viations_approchantes.List(k - 1, 1) = AutoCorrect.Entries(MyIndexAutocorrect).Value
''        ' Me.zone_abr�viations_approchantes.SetFocus
''
''         End If 'AutoCorrect.Entries(j).Name Like MyAb1 Or AutoCorrect.Entries(j).Name Like MyAb2
''
''     '   Next 'j
''
''   If MyExistingAb = 0 Then MyInputBox.ZoneMotsCorrespondants.Clear
 
strsql = "SELECT Count(abr�viations.nom) AS CompteDenom FROM abr�viations HAVING (((abr�viations.nom)=""" & myab & """));"
 'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
Set rdshippers = dbNorthwind.OpenRecordset(strsql)
MyNumberRecords = rdshippers.Fields("CompteDenom")

If MyNumberRecords > 0 Then
strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """));"
Set rdshippers = dbNorthwind.OpenRecordset(strsql)
rdshippers.MoveFirst
    For i = 1 To MyNumberRecords
        MyInputBox.ZoneMotsCorrespondants.AddItem rdshippers.Fields(1)
    
    
    rdshippers.MoveNext
    Next



End If



 
 
 
MyInputBox.listes_d�roulantes.Value = 0

End Sub



Public Function extraire_id_current_version(MyValue As Variant, MyTable)
'Dim i, ThereIsVerb, MyType
'Dim mysearch As Variant
''MyValue = "mod�r�ment"
''MyTable = "table_m�re"
' Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
'  Dim rdshippers As Recordset
'If MyTable = "infinitifs_avec_z" Or MyTable = "infinitifs_sans_z" Or IsEmpty(MyTable) Then
'
'    If MyInfinitif <> "" Then
'    mysearch = MyInfinitif
'    Else
'    mysearch = MyValue
'    End If
'
'    Set rdshippers = dbNorthwind.OpenRecordset("table_m�re")
'
'    With rdshippers
'        .Index = "forme"
'         .Seek "=", mysearch
'
'
'        End With
'
'      If rdshippers.NoMatch = False Then
'
'                  extraire_id = rdshippers.Fields("id")
'
'                    Exit Function
'        End If
'
'Else
'
'    mysearch = MyValue
'
'
'
'
'
'
'   ' Dim dbNorthwind As DAO.Database
'
'    Set rdshippers = dbNorthwind.OpenRecordset(MyTable)
'    Dim Matabledef As TableDef
'    Set Matabledef = dbNorthwind.TableDefs(MyTable)
'    For i = 0 To Matabledef.Indexes.Count - 1
'
'
'    With rdshippers
'        .Index = Matabledef.Indexes(i).Name
'        If rdshippers.Fields(Matabledef.Indexes(i).Name).Type = 10 Then
'
'
'        .Seek "=", mysearch
'        Else
'        GoTo NextIndex
'        End If
'
'        End With
'
'                   If rdshippers.NoMatch = False Then
'
'                  extraire_id = rdshippers.Fields("id")
'
'                    Exit Function
'
'
'                   End If
'NextIndex:
'
'    Next
'End If
'

End Function

Public Function IsLettreSeuleOk(myab)

Dim strsql
Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Dim rdshippers As Recordset
strsql = "SELECT lettres_seules.lettre FROM lettres_seules WHERE (((lettres_seules.lettre)=""" & myab & """));"


Set rdshippers = dbNorthwind.OpenRecordset(strsql)

If rdshippers.RecordCount = 0 Then
IsLettreSeuleOk = False
Else
IsLettreSeuleOk = True
End If





End Function

Public Function extraire_id(MyValue, MyTable)
Dim i, ThereIsVerb, MyType, MyIndex
Dim Mysearch As Variant
MyIndex = Array("forme", "f�minin", "f�minin_pluriel", "pluriel")

MyTable = "table_m�re"
'MyValue = "mod�r�ment"
'MyTable = "table_m�re"
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
 
 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
  Dim rdshippers As Recordset
'If MyTable = "infinitifs_avec_z" Or MyTable = "infinitifs_sans_z" Or IsEmpty(MyTable) Then

    If MyInfinitif <> "" Then
    Mysearch = MyInfinitif
    Else
    Mysearch = MyValue
    End If
    
    Set rdshippers = dbNorthwind.OpenRecordset("table_m�re")
    
    For i = 0 To 3
    With rdshippers
        .Index = MyIndex(i)
       
         .Seek "=", Mysearch
        
        End With
    
      If rdshippers.NoMatch = False Then
                   
                  extraire_id = rdshippers.Fields("id")
                   
                    Exit Function
                    
                    
        End If
     Next
'Else

'    Mysearch = MyValue
'
'
'
'
'
'
'   ' Dim dbNorthwind As DAO.Database
'
'    Set rdshippers = dbNorthwind.OpenRecordset(MyTable)
'    Dim Matabledef As TableDef
'    Set Matabledef = dbNorthwind.TableDefs(MyTable)
'    For i = 0 To Matabledef.Indexes.Count - 1
'
'
'    With rdshippers
'        .Index = Matabledef.Indexes(i).Name
'        If rdshippers.Fields(Matabledef.Indexes(i).Name).Type = 10 Then
'
'
'        .Seek "=", Mysearch
'        Else
'        GoTo NextIndex
'        End If
'
'        End With
'
'                   If rdshippers.NoMatch = False Then
'
'                  extraire_id = rdshippers.Fields("id")
'
'                    Exit Function
'
'
'                   End If
'NextIndex:
'
'    Next
'End If
  

End Function

Sub maj_abr�viations_utilis��es_dans_myinputbox(myab)


Dim MyNumberSuggestions, i

MyNumberSuggestions = MyInputBox.suggestions.ListCount

For i = 0 To MyNumberSuggestions - 1

MyInputBox.zone_mot.AddItem MyInputBox.suggestions.List(i)


Next

SendKeys "+{home}"


End Sub





Public Function rechercher_previous_search(myab)
Dim rdshippers As Recordset, i, strsql, D�j�Abr�g�, j
D�j�Abr�g� = choix_abr�viation.choix_ab.ListCount
choix_abr�viation.previous_search_non_selected.Clear
j = 0
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
    Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
   Set rdshippers = dbNorthwind.OpenRecordset("previous_search")

With rdshippers
        .Index = "nom"
        .Seek "=", myab

End With

If rdshippers.NoMatch = False Then
rechercher_previous_search = -1
strsql = "SELECT previous_search.nom, previous_search.valeur FROM previous_search WHERE (((previous_search.nom)=""" & myab & """))ORDER BY previous_search.compteur;"
    Set rdshippers = dbNorthwind.OpenRecordset(strsql)
        
        rdshippers.MoveFirst
        
        While Not rdshippers.EOF
        
            If D�j�Abr�g� > 0 Then
                For i = 0 To D�j�Abr�g� - 1
                    If rdshippers.Fields("valeur") <> choix_abr�viation.choix_ab.List(i) Then j = j + 1
                      
                Next i
        
                   
                    
                
            If j = D�j�Abr�g� Then choix_abr�viation.previous_search_non_selected.AddItem rdshippers.Fields("valeur")
            j = 0
            End If
            
            If D�j�Abr�g� = 0 Then
                
                  choix_abr�viation.previous_search_non_selected.AddItem rdshippers.Fields("valeur")
            
        
            End If 'D�j�Abr�g� > 0
        
        rdshippers.MoveNext
        Wend

Else
rechercher_previous_search = 0

End If




End Function
Sub essai()
Attribute essai.VB_Description = "Macro enregistr�e le 31/03/2012 par SGA-EB"
Attribute essai.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.essai"
Dim rdshippers As Recordset, strsql


Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """));"


'
    Selection.TypeText Text:=" "
End Sub

Public Sub Search_direct_access_ab(myab)

Dim rdshippers As Recordset, strsql
strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.registre FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """) AND ((abr�viations.registre)=Yes));"
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")
Set rdshippers = dbNorthwind.OpenRecordset(strsql)

 If rdshippers.BOF = False Then
    rdshippers.MoveFirst
    choix_abr�viation.choix_ab.AddItem rdshippers.Fields("valeur")
    
 End If
 
 






End Sub

Public Function valeur_lettres_abr�viations(myab)

Dim rdshippers As Recordset
 
Dim MyLen, MyLetter, i, MyValue
'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")


MyLen = Len(myab)

For i = 1 To MyLen

MyLetter = Mid(myab, i, 1)

Set rdshippers = dbNorthwind.OpenRecordset("lettres")

With rdshippers
        .Index = "lettre"
        .Seek "=", MyLetter

End With

valeur_lettres_abr�viations = valeur_lettres_abr�viations + rdshippers.Fields(1)



Next




valeur_lettres_abr�viations = valeur_lettres_abr�viations + MyLen
















End Function

Public Sub peupler_valeur_lettres_ab()

Dim rdshippers As Recordset
Dim MyNomAb, i

 
Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")


Set rdshippers = dbNorthwind.OpenRecordset("abr�viations")


With rdshippers
i = 1
.MoveFirst
        MyNomAb = rdshippers.Fields("nom")
        'Debug.Print MyNomAb
        While rdshippers.EOF = False
          MyNomAb = rdshippers.Fields("nom")
        'Debug.Print MyNomAb
        
        .Edit
        !valeur_lettres_ab = valeur_lettres_abr�viations(MyNomAb)
        'Debug.Print valeur_lettres_abr�viations(MyNomAb)
        .Update
        
        
        .MoveNext
       i = i + 1
        Wend


End With

MsgBox i


End Sub


Public Sub peupler_ab_similaires(myab, MyForm)

Dim rdshippers As Recordset

 'Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")
Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")





Dim strsql, MyValAb, i
MyValAb = valeur_lettres_abr�viations(myab)

strsql = "SELECT abr�viations.nom, abr�viations.valeur, abr�viations.valeur_lettres_ab FROM abr�viations WHERE (((abr�viations.valeur_lettres_ab)=" & MyValAb & ") AND ((abr�viations.taille)=" & Len(myab) & "));"
'WHERE (((abr�viations.valeur_lettres_ab)=16) AND ((abr�viations.taille)=6));

Set rdshippers = dbNorthwind.OpenRecordset(strsql)

i = 1

With rdshippers

If .NoMatch = True Or .RecordCount < 1 Then Exit Sub
.MoveFirst
While .EOF = False


        Select Case MyForm

        Case "MyInputbox"

        If myab <> .Fields("valeur") Then
            
            If Left(.Fields("nom"), 1) = Left(myab, 1) Then
        
                
                
                    
        
        
        MyInputBox.zone_ab_approchantes.AddItem .Fields("valeur")
      
        MyInputBox.zone_ab_approchantes.List(i - 1, 1) = .Fields("nom")
        MyInputBox.�tiquette_inversion.Visible = True
        
          i = i + 1
           
           End If
            
        End If
        
                    Case "choix_abr�viation"
           
         If myab <> .Fields("nom") Then
            
            If Left(.Fields("nom"), 1) = Left(myab, 1) Then
           
           
       choix_abr�viation.zone_ab_approchantes.AddItem .Fields("valeur")
         choix_abr�viation.zone_ab_approchantes.List(i - 1, 1) = .Fields("nom")
        choix_abr�viation.�tiquette_inversion.Visible = True
           
          i = i + 1
           
           End If
            
        End If
           
           
                End Select
                
           
           
           
         
        .MoveNext
        
        
        Wend



End With




End Sub

Sub espace()
Attribute espace.VB_Description = "Macro enregistr�e le 06/05/2012 par SGA-EB"
Attribute espace.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.espace"
'
' espace Macro
' Macro enregistr�e le 06/05/2012 par SGA-EB
'
    Selection.TypeText Text:=" "
End Sub

Public Sub d�velopper_espace()
'd�sactiver_correct (False)
'm�thode_ab.Hide

Load choix_abr�viation
 FonctionEnCours = "d�velopper"
UsageRechercheMot = ""
'hypoth�se de d�part : on a tap� une abr�viation mais rien n'appara�t. On va donc
'pouvoir cr�er la valeur correspondant � l'abr�viation. Appel� par "Control f"

Dim MyDate, MyOrdinateur, MyOrdinateurLastSave
Dim MyVerb, strsql

Dim MyDataObject As MSForms.DataObject
Set MyDataObject = New MSForms.DataObject
Dim Myentry, MyMsg, MySuggestion, MyWord As String
Dim mySpelling, MySpellingAb As Boolean, MyNewWordModifi�, MySettingAccords
Dim MyAutoCorrects, i, MyReplaceEntry, j, sNombre, sPasDeSuggestion As Integer
Dim MySpell As Dictionary, MyWordSansS, TailleMyExistingAb, MyExistingAbLastLetter, TailleMyExistingWord
Dim MyValue As Variant, myText, myAbSansS, MyOrthographe, myAbLastLetter
Dim MyActiveDocument As Document, MyAb1, MyAb2, k, TailleMyAb, MyExistingAb, TailleMyNewAb, TailleMyNewWord, MyPreviousSearch
Dim MyApos 'ajout�e le 29/04/2009 �ventuellement un doublon
Dim rdshippers As Recordset
Dim MyRegister

Dim MyConjugaisonRare

MyConjug = ""
MyReplaceEntry = 0
MyApostrophe = ""
myab = ""
MyPbkMsg = ""
MyValue = ""

'Dim myFootNote As Boolean, MyNomDoc, MyWindow As Window
Set MyActiveDocument = Application.ActiveDocument
'ICI ICI ICI MODIFICATION
'�a s'�crivait comme cela :
' Set dbNorthwind = OpenDatabase(Name:=get_hd & ":\fasttype\mots_reverses.mdb")

 Set dbNorthwind = OpenDatabase(Name:=get_hd & "\fasttype\mots_reverses.mdb")

MyHeureD�but = Timer
Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
myab = LCase(Selection.Text)
MyTable = Empty

Dim sTerminaison_�_1, MyPonctuation
MyDate = get_param�tres("date_usage")
MyOrdinateur = get_param�tres("cet ordinateur")
MyOrdinateurLastSave = get_param�tres("ordinateur last saving")

MyInputBox.zone_mot.Clear
MyInputBox.suggestions.Clear
MyInputBox.rejet�s.Clear
MyInputBox.stock.Clear
MyInputBox.zone_abr�viation_existantes.Clear
MyInputBox.zone_abr�viations_approchantes.Clear

Dim LastAb, ThisComputer
'''''''''''''''''''''''
'
'MyAutoCorrects = AutoCorrect.Entries.Count
'strsql = "SELECT Count(abr�viations.nom) AS CompteDenom, abr�viations.registre FROM abr�viations GROUP BY abr�viations.registre HAVING (((abr�viations.registre)=-1));"
'Set rdshippers = dbNorthwind.OpenRecordset(strsql)
'MyRegister = rdshippers.Fields("CompteDeNom")
'If MyAutoCorrects < MyRegister Then extraire_abr�viations




'''MyDate = GetSetting("fasttype", section:="param�tres", Key:="date_usage")
'''ThisComputer = GetSetting("fasttype", section:="param�tres", Key:="cet ordinateur")
'''LastAb = get_settings_from_bdd(3)
'''If ThisComputer <> LastAb Then extraire_abr�viations


''''''''''''''''''''''''''''''''

'UpDateLastAb


If Selection.Information(wdInFootnote) Then

            Set MyWindow = MyActiveDocument.ActiveWindow
            myFootNote = True

End If 'Selection.Information(wdInFootnote)



myab = LCase(myab)
myab = Replace(myab, " ", "")
myab = Trim(myab)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


strsql = "SELECT Count(abr�viations.nom) AS CompteDenom FROM abr�viations GROUP BY abr�viations.nom HAVING (((abr�viations.nom)=""" & myab & """));"




''strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """));"
Set rdshippers = dbNorthwind.OpenRecordset(strsql)
If rdshippers.RecordCount = 0 Then GoTo v�rif_ortho


Select Case rdshippers.Fields("CompteDenom")

Case 1


 strsql = "SELECT abr�viations.nom, abr�viations.valeur FROM abr�viations WHERE (((abr�viations.nom)=""" & myab & """));"
 Set rdshippers = dbNorthwind.OpenRecordset(strsql)
Selection.TypeText Text:=rdshippers.Fields("valeur")
Selection.TypeText Text:=" "
  End

Case 0
v�rif_ortho:
MySpellingAb = Application.CheckSpelling(myab)

Select Case MySpellingAb

Case True
MyValue = myab
GoTo conjugaison_rare_d�tect�e

End Select







End Select

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




If Len(myab) < 2 Then 'on oblige, si il s'agit d'une seule lettre, � passer par le formulaire

        
    If fonction_d�tecter_ponctuation(myab) = True Then
    
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    
         myab = RTrim(Selection.Text)
    End If
    
    
            If fonction_d�tecter_ponctuation(myab) = True Then 'il faut r�p�ter la fonction pour les espaces ins�cables ins�r�s
                 'automatiquement par Word avant un ! ou : ou ;
                
                 myab = Left(myab, Len(myab) - 1)
                 MyPonctuation = 1
            
            End If
        
End If 'Len(MyAb) < 2


'ins�rer ici la recherche dans la table conjugaisons rares







If Len(myab) < 2 Then


   Select Case IsLettreSeuleOk(myab)
        Case -1
 GoTo chercher_utilisation_abr�viation:
   
        Case 0
            sMessage "la lettre :" & Chr(10) & Chr(10) & myab & Chr(10) & Chr(10) & " ne peut servir d'abr�viation car est elle est signifiante", "annuler", "rien", "rien", "rien", "Lettre signifiante", 255, 0
        End
        Exit Sub


    End Select 'MayBeAlone(MyAb)


End If 'Len(MyAb) < 2

If Len(myab) > 2 Then

    If fonction_d�tecter_ponctuation(myab) = True Then
    
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    
         myab = LCase(RTrim(Selection.Text))
    
        If fonction_d�tecter_ponctuation(myab) = True Then 'il faut r�p�ter la fonction pour les espaces ins�cables ins�r�s
            'automatiquement par Word avant un ! ou : ou ;
            'Selection.MoveLeft Unit:=wdCharacter, Count:=1, extend=wdExtend
            'Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
            myab = Right(myab, Len(myab) - 1)
        
        
        End If
        
    End If
    
        If d�tecter_apostrophe(myab) = True Then

                MyApostrophe = Left(myab, MyPosition)
                myab = LCase(Trim(Right(myab, Len(myab) - MyPosition)))

        End If
    
    
    
End If ' Len(MyAb) > 2



MyConjugaisonRare = recherche_conjugaison_rare(myab)
MyPreviousSearch = rechercher_previous_search(myab)

        choix_abr�viation.choix_ab = ChoixAbr�viation
''
''



If MyConjugaisonRare <> 0 Then

        Select Case MyPreviousSearch


            Case 0
'         choix_abr�viation.choix_ab.Clear
'         recherche_conjugaison_rare (myab)
     If choix_abr�viation.previous_search_non_selected.ListCount = 0 Then
         choix_abr�viation.Width = 330.75
         Else
         choix_abr�viation.Width = 633.75
         End If
        peupler_ab_similaires myab, "choix_abr�viation"
        choix_abr�viation.Show
        
            Case -1
'              choix_abr�viation.choix_ab.Clear
'            recherche_conjugaison_rare (myab)
'                choix_abr�viation.choix_ab = ChoixAbr�viation
            choix_abr�viation.previous_search_non_selected.Clear
            rechercher_previous_search (myab)
            Search_direct_access_ab (myab)
        

         
         Load choix_abr�viation
          peupler_ab_similaires myab, "choix_abr�viation"
         choix_abr�viation.Show
        
        End Select
    
    If MySelectionInPrevious = -1 Then GoTo AfterAcronyme

        If ChoixAbr�viation <> 0 Then
        MyValue = ChoixAbr�viation

        GoTo conjugaison_rare_d�tect�e

       End If




End If

If MyConjugaisonRare = 0 And MyPreviousSearch = -1 Then
    choix_abr�viation.previous_search_non_selected.Clear
    Search_direct_access_ab (myab)
    rechercher_previous_search (myab)
     peupler_ab_similaires myab, "choix_abr�viation"
    choix_abr�viation.Show
End If

If MySelectionInPrevious = -1 Then GoTo AfterAcronyme


'il faut rechercher si l'abr�viation n'est pas un acronyme

' A REVOIR


MySpellingAb = Application.CheckSpelling(myab)
        
        Select Case MySpellingAb
        
              Case True 'MySpellingAb. Le mot est correctement orthographi�, donc a priori, il ne peut
              'servir d'abr�viation. On renvoie � l'abr�viation des mots.
              
             ' sMessage myab & " est dans le dictionnaire. Que voulez-vous faire ", "Annuler", "Cr�er quand m�me une abr�viation", "rien", "rien", "Groupe de lettres pr�sent dans le dictionnaire", 255, 0
              MyInputBox.texte2 = "attention : " & myab & " est dans le dictionnaire"
              MyInputBox.texte2.ForeColor = 255
              MyInputBox.zone_abr�viation.BackColor = &H80FFFF
              
             
        
        End Select 'MySpellingAb pour myab


Set rdshippers = dbNorthwind.OpenRecordset("temps_combinaison")

With rdshippers
        .Index = "nom"
        .Seek "=", Right(myab, 2)

End With

If rdshippers.NoMatch = False Then


                EndIsConjug = -1
               
                MyConjug = Right(myab, 2)
                
                GoTo skip2

End If

With rdshippers
        .Index = "nom"
        .Seek "=", Right(myab, 3)

End With

If rdshippers.NoMatch = False Then


                EndIsConjug = -1
               
                MyConjug = Right(myab, 3)
                
                GoTo skip2

End If

skip2:
If MyConjug <> "" Then 'si on veut d�clencher aussi cela pour la lettre r finale, il faut voir ensuite au niveau des r�sultats
'on passe l'information qu'il y a une conjugaison
'MyTerminaison1 = GetSetting(appname:="fasttype", section:="conjugaisons_deuxi�me", Key:=MyConjug)
'MyTerminaison2 = GetSetting(appname:="fasttype", section:="conjugaisons_premier", Key:=MyConjug)
sMessage "Pensiez-vous � un verbe conjugu� ?", "rien", "Verbe conjugu�", "autre (y compris infinitif)", "rien", "Orienter la recherche", "bleu", 2
MyHeureD�but = Timer

    Select Case MyPbkMsg
    Case 2
        EndIsConjug = -1 '
    Case Else
         EndIsConjug = ""
    End Select

End If 'MyConjug <> ""
 
 

'
MyAutoCorrects = AutoCorrect.Entries.Count
TailleMyAb = Len(myab)

myAbLastLetter = Right(myab, 1) 'on analyse quelle est la derni�re lettre de l'abr�viation


'Select Case myAbLastLetter
''skip_recherche_accord:
'
'    Case "m" 'l'abr�viation correspond � l'abr�viation de ment (adverbe)
'    'on recherche si existe une abr�viation sans le "m" (abr�viation de "ment")
'    '� faire : il faudra extraire le m convenablement, si jamais on change de m�thode pour abr�ger les adverbes
'
'        If check_existence_valeur_pour_abr�viation(Left(myab, TailleMyAb - 1)) Then
'
'        'si la valeur est trouv�e, on examine sa derni�re lettre
'
'        TailleMyExistingAb = Len(AutoCorrect.Entries(MyIndexAutocorrect).Name)
'        MyExistingAbLastLetter = Right(AutoCorrect.Entries(MyIndexAutocorrect).Name, 1)
'        TailleMyExistingWord = Len(AutoCorrect.Entries(MyIndexAutocorrect).Value)
'
'
'
'                    MyNewWord = Left(AutoCorrect.Entries(MyIndexAutocorrect).Value, TailleMyExistingWord - 1) & "ment"
'
'
'                    MyOrthographe = Application.CheckSpelling(MyNewWord)
'
'                        Select Case MyOrthographe 'le mot qui pourrait correpondre est-il bien orthographi�
'
'                            Case True 'le mot est bien orthographi�
'
'                              load_accords
'
'                                Select Case MyPbkMsg
'
'                                        Case "xxxxx"
'
'                                        Case Else
'
'                                       ' AutoCorrect.Entries.Add MyAb, MyPbkMsg
'                                        stocker_abr�viations myab, MyPbkMsg, False, False, MyId
'
'                                        Application.ActiveDocument.Activate
'                                        If MyApostrophe <> "" Then MyPbkMsg = MyApostrophe & MyPbkMsg
'                                        Selection.TypeText Text:=MyPbkMsg
'
'                                        Selection.MoveRight Unit:=wdCharacter, Count:=1
'                                        Exit Sub
'
'                                End Select 'MyPbkMsg
'
'                            Case False
'
'
'                End Select
'
'
'
'
'        End If 'check_existence_valeur_pour_abr�viation(Left(MyAb, TailleMyAb - 1))
'
'   ' Next j
'
'
'
'
'        If check_existence_valeur_pour_abr�viation(myab & "m") Then  'on recherche si l'adverbe qui formerait le mot dont on a
'        'form� l'abr�viation existe
'        'il faudra rechercher en d�terminant la lettre abr�viative de "ment"
'
'        'si la valeur est trouv�e, on examine sa derni�re lettre
'
'        TailleMyExistingAb = Len(AutoCorrect.Entries(MyIndexAutocorrect).Name)
'        MyExistingAbLastLetter = Right(AutoCorrect.Entries(MyIndexAutocorrect).Name, 1)
'        TailleMyExistingWord = Len(AutoCorrect.Entries(MyIndexAutocorrect).Value)
'
'
'
'
'
'
'                    MyNewWord = Left(AutoCorrect.Entries(MyIndexAutocorrect).Value, TailleMyExistingWord - 4)
'
'
'                    MyOrthographe = Application.CheckSpelling(MyNewWord)
'
'                        Select Case MyOrthographe 'le mot qui pourrait correpondre est-il bien orthographi�
'
'                            Case True 'le mot est bien orthographi�
'
'                            load_accords
'
'                                Select Case MyPbkMsg
'
'                                        Case "xxxxx"
'
'                                        Case Else
'
'                                        'AutoCorrect.Entries.Add MyAb, MyPbkMsg
'                                         stocker_abr�viations myab, MyPbkMsg, False, False, MyId
'                                         enseigner_abr�viations myab, MyPbkMsg
'                                        Application.ActiveDocument.Activate
'                                        If MyApostrophe <> "" Then MyPbkMsg = MyApostrophe & MyPbkMsg
'                                        Selection.TypeText Text:=MyPbkMsg & " "
'                                                                              '
'
'
'                                        Selection.MoveRight Unit:=wdCharacter, Count:=1
'                                        Exit Sub
'
'                                    End Select 'MyPbkMsg
'
'                            Case False
'
'
'                    End Select 'MyOrthographe
'
'
'
'
'
'            End If 'AutoCorrect.Entries(j).Name = myab & "m"
'
'
'End Select 'myablastletter

If Len(myab) = 2 Then
'une abr�viation de trois lettres sort pratiquement toujours des r�sultats trop nombreux






GoTo chercher_utilisation_abr�viation:

End If

recherche_mot_depuis_abr�viation myab


'on ressort ici quand on a utilis� une abr�viation de 3 lettres et plus
'on enregistre imm�diatement la bonne abr�viation

'rechercher si l'abr�viation choisie existe d�j� dans le registre
'zorro



'si ce n'est pas le cas, on garde les choses telles quelles

'si c'est le cas, l'abr�viation doit �tre marqu�e comme n'allant plus dans le registre

'si l'abr�viation existe avec une valeur diff�rente, il faut la supprimer du registre
If MyRepeat = 99 Then GoTo AfterAcronyme 'cette ligne est �crite pour l'hypoth�se ou on utilise le champ zone_ab_approchantes pour ins�rer dans le texte
'sans enregistrer l'abr�vation (cas d'une inversion de lettre).


MyId = extraire_id(MyPbkMsg, MyTable)


  If GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = "vrai" And IsEmpty(MyId) Then
        
            AddMot.forme = MyPbkMsg
            AddMot.f�minin = MyPbkMsg
            AddMot.f�minin_pluriel = MyPbkMsg
            AddMot.singulier_pluriel = MyPbkMsg
            AddMot.Caption = "toutes les formes de : " & myab
                If IsZ(MyPbkMsg) = True Then
                        AddMot.z = True
                        AddMot.z.Caption = "� contr�ler"
                Else
                        AddMot.z = False
                End If
                    AddMot.Show
                    If ResultAddMot = 9999 Then
                    End
                    Exit Sub
                    End If
                    
   
         
        End If 'GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = -1




 stocker_abr�viations myab, MyPbkMsg, False, False, MyId
enseigner_abr�viations myab, MyPbkMsg
   
  
   
   'insertion du mot trouv� dans le texte
        
        
        Application.ActiveDocument.Activate
            If MyApostrophe <> "" Then MyNewWord = MyApostrophe & MyNewWord
            If MyPonctuation = 1 Then MyNewWord = MyNewWord & " "
        
        If myFootNote = True Then
            
          
            MyWindow.Panes(2).Activate
        End If

        
        
        
            Selection.TypeText Text:=MyNewWord & " "
'            Selection.MoveRight Unit:=wdCharacter, Count:=1
            dbNorthwind.Close
            End
            Exit Sub
            
                
 


chercher_utilisation_abr�viation myab
chercher_utilisation_abr�viation:

CarryOn:
recommencer:
MyInputBox.zone_abr�viation_existantes.Enabled = True
MyInputBox.zone_abr�viation_existantes.Visible = True
MyInputBox.bouton_supprimer_abr�viation.Visible = True
myText = "Entrez le mot correspondant � l'abr�viation"
'chercher_utilisation_abr�viation myab

'ouverture de myinputbox

OpenMyInputBox myText, myab


'si on myab = taille de 2 caract�res, on sort ici.
GoTo AfterAcronyme
acronyme:
OpenMyInputBox myab & " est d�j� l'acronyme de " & MyValue & ". Si vous souhaitez le remplacer, entrez une nouvelle valeur sinon annulez ?", myab


AfterAcronyme:


MyValue = MyPbkMsg
MySaisie = MyPbkMsg 'on a besoin d'une variable en plus, car mypbkmsg sert aussi dans une autre fonction
'et sa valeur peut changer

'MyAb = MyAbr�viation

If MyValue = 0 Then Exit Sub

     If MyRepeat = 99 Then 'on veut juste ins�rer le mot sans l'abr�ger. La valeur de myrepeat est pass�e par le bouton
     'ins�rer_mot du formulaire myinputbox
     If MyApostrophe <> "" Then MyValue = MyApostrophe & MyValue
     If MyPonctuation = 1 Then MyValue = MyValue & " "
     
     If myFootNote = True Then
            
          MyWindow.Panes(2).Activate
            
     End If
     
     
     Selection.TypeText Text:=MyValue & " "
        dbNorthwind.Close
        End
'    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Exit Sub
End If

 
createab:
 
MyId = extraire_id(MyValue, MyTable)

 If GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = "vrai" And IsEmpty(MyId) Then
        
            AddMot.forme = MyPbkMsg
            AddMot.f�minin = MyPbkMsg
            AddMot.f�minin_pluriel = MyPbkMsg
            AddMot.singulier_pluriel = MyPbkMsg
            AddMot.Caption = "toutes les formes de : " & myab
                If IsZ(MyPbkMsg) = True Then
                        AddMot.z = True
                        AddMot.z.Caption = "� contr�ler"
                Else
                        AddMot.z = False
                End If
                    AddMot.Show
                    If ResultAddMot = 9999 Then Exit Sub
   
   
         
        End If 'GetSetting(appname:="fasttype", section:="param�tres", Key:="AddMot") = -1

 
MyId = extraire_id(MyValue, MyTable)

stocker_abr�viations myab, MyValue, False, False, MyId


enseigner_abr�viations myab, MyValue '� compl�ter avec la recherche des participes

'If IsVerb(MyValue) Then conjuguer_un_verbe_depuis_table MyValue, MyAb



conjugaison_rare_d�tect�e:


If MyApostrophe <> "" Then MyValue = MyApostrophe & MyValue
If MyPonctuation = 1 Then MyValue = MyValue & " "

        If myFootNote = True Then
            
          MyWindow.Panes(2).Activate
            
        End If


Selection.TypeText Text:=MyValue & " "
dbNorthwind.Close
'Selection.MoveRight Unit:=wdCharacter, Count:=1


    Select Case MyRepeat
 
        Case 10

        GoTo recommencer:

 
 
    End Select
 End
'd�sactiver_correct (True)
Exit Sub


End Sub
