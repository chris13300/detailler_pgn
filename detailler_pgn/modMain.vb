Imports System.IO
Imports VB = Microsoft.VisualBasic

Module modMain

    Sub Main()
        Dim PGNFile As String, chaine As String, suffixeDetailled As String
        Dim detailledFile As String, cleanedFile As String, completedFile As String

        PGNFile = Replace(Command(), """", "")

        If InStr(nomFichier(PGNFile), " ") > 0 Then
            My.Computer.FileSystem.RenameFile(PGNFile, Replace(nomFichier(PGNFile), " ", "_"))
            PGNFile = Replace(PGNFile, nomFichier(PGNFile), Replace(nomFichier(PGNFile), " ", "_"))
        End If

        suffixeDetailled = "_detailled"

        If PGNFile = "" Then
            MsgBox("file not found")
            Exit Sub
        ElseIf PGNFile <> "" And (Not My.Computer.FileSystem.FileExists(PGNFile) Or extensionFichier(PGNFile) <> ".pgn") Then
            MsgBox("file not found")
            Exit Sub
        End If

        '1. scid import => Deleted comments => export
        Console.WriteLine("1. USE SCID to delete comments")
        Console.WriteLine("Press a key when it's done...")
        Console.ReadLine()

        '2. clean stats and header
        chaine = nettoyer(PGNFile)
        cleanedFile = Replace(PGNFile, ".pgn", "_cleaned.pgn")
        My.Computer.FileSystem.WriteAllText(cleanedFile, chaine, False)
        Console.Clear()
        Console.WriteLine("2. CLEANING : terminated !")

        '3. export detailled moves (a4 => a2a4)
        pgnDETAIL(cleanedFile, suffixeDetailled)
        Console.Clear()
        Console.WriteLine("3. PGN-EXTRACT : terminated !")

        '4. rebuild pgn with detailled moves
        detailledFile = Replace(cleanedFile, ".pgn", suffixeDetailled & ".pgn")
        chaine = coupsDetailles(cleanedFile, detailledFile)
        completedFile = Replace(PGNFile, ".pgn", "_complete.pgn")
        My.Computer.FileSystem.WriteAllText(completedFile, chaine, False)
        Console.Clear()
        Console.WriteLine("4. DETAILLED MOVES : terminated !")

        'fichier_cleaned.pgn
        My.Computer.FileSystem.DeleteFile(cleanedFile, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

        'fichier_cleaned.log
        My.Computer.FileSystem.DeleteFile(Replace(cleanedFile, ".pgn", ".log"), FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

        'fichier_cleaned_detailled.pgn
        My.Computer.FileSystem.DeleteFile(detailledFile, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

        Console.WriteLine("Press a key to exit")
        Console.ReadLine()

    End Sub

    Private Function coupsDetailles(cleanedFile As String, detailledFile As String) As String
        Dim base As String, detail As String, tabBase() As String, tabDetail() As String, depart As Integer
        Dim i As Integer, tmpBase() As String, tmpDetail() As String, j As Integer, separateur As String

        base = My.Computer.FileSystem.ReadAllText(cleanedFile)
        base = Replace(base, ". ... ", ". ")
        base = Replace(base, " ... ", " ")
        While base.LastIndexOf(vbCrLf) = Len(base) - 2
            base = gauche(base, Len(base) - 2)
            System.Threading.Thread.Sleep(10)
        End While
        base = Replace(base, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        tabBase = Split(base, vbCrLf & vbCrLf)

        detail = My.Computer.FileSystem.ReadAllText(detailledFile)
        detail = Replace(detail, ". ... ", ". ")
        detail = Replace(detail, " ... ", " ")
        While detail.LastIndexOf(vbCrLf) = Len(detail) - 2
            detail = gauche(detail, Len(detail) - 2)
            System.Threading.Thread.Sleep(10)
        End While
        detail = Replace(detail, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        tabDetail = Split(detail, vbCrLf & vbCrLf)

        If tabBase.Length <> tabDetail.Length Then
            MsgBox("problem #1")
            Return ""
        End If

        depart = Environment.TickCount
        For i = 0 To UBound(tabBase)
            If InStr(tabBase(i), "1. ") = 1 And InStr(tabDetail(i), "1. ") = 1 _
            And InStr(tabBase(i), "[") = 0 And InStr(tabDetail(i), "[") = 0 _
            And InStr(tabBase(i), "]") = 0 And InStr(tabDetail(i), "]") = 0 Then
                base = Replace(tabBase(i), vbCrLf, " ")
                Do
                    base = Replace(base, "  ", " ")
                Loop While InStr(base, "  ") > 0
                tmpBase = Split(base, " ")

                detail = Replace(tabDetail(i), vbCrLf, " ")
                Do
                    detail = Replace(detail, "  ", " ")
                Loop While InStr(detail, "  ") > 0
                tmpDetail = Split(detail, " ")

                If tmpBase.Length <> tmpDetail.Length Then
                    If tmpBase.Length = tmpDetail.Length + 1 _
                    And tmpDetail(UBound(tmpDetail)) = tmpBase(UBound(tmpBase)) _
                    And InStr(tmpBase(UBound(tmpBase) - 1), ".") > 0 Then
                        tmpBase(UBound(tmpBase) - 1) = tmpBase(UBound(tmpBase))
                        ReDim Preserve tmpBase(UBound(tmpBase) - 1)
                    Else
                        MsgBox("problem #2")
                        Return ""
                    End If
                End If

                For j = 0 To UBound(tmpBase)
                    If InStr(tmpBase(j), ".") = 0 And InStr(tmpDetail(j), ".") = 0 _
                    And tmpBase(j) <> "*" And tmpDetail(j) <> "*" _
                    And tmpBase(j) <> "1/2-1/2" And tmpDetail(j) <> "1/2-1/2" _
                    And tmpBase(j) <> "1-0" And tmpDetail(j) <> "1-0" _
                    And tmpBase(j) <> "0-1" And tmpDetail(j) <> "0-1" Then

                        separateur = "-"
                        If InStr(tmpBase(j), "x") > 0 Then
                            separateur = "x"
                        End If

                        If (tmpBase(j) = "O-O" Or tmpBase(j) = "0-0") And (tmpDetail(j) = "e8g8" Or tmpDetail(j) = "e1g1") Then

                            tmpBase(j) = "0-0"

                        ElseIf (tmpBase(j) = "O-O+" Or tmpBase(j) = "0-0+") And (tmpDetail(j) = "e8g8+" Or tmpDetail(j) = "e1g1+") Then

                            tmpBase(j) = "0-0+"

                        ElseIf (tmpBase(j) = "O-O#" Or tmpBase(j) = "0-0#") And (tmpDetail(j) = "e8g8#" Or tmpDetail(j) = "e1g1#") Then

                            tmpBase(j) = "0-0#"

                        ElseIf (tmpBase(j) = "O-O-O" Or tmpBase(j) = "0-0-0") And (tmpDetail(j) = "e8c8" Or tmpDetail(j) = "e1c1") Then

                            tmpBase(j) = "0-0-0"

                        ElseIf (tmpBase(j) = "O-O-O+" Or tmpBase(j) = "0-0-0+") And (tmpDetail(j) = "e8c8+" Or tmpDetail(j) = "e1c1+") Then

                            tmpBase(j) = "0-0-0+"

                        ElseIf (tmpBase(j) = "O-O-O#" Or tmpBase(j) = "0-0-0#") And (tmpDetail(j) = "e8c8#" Or tmpDetail(j) = "e1c1#") Then

                            tmpBase(j) = "0-0-0#"

                        ElseIf Len(tmpBase(j)) = 2 And Len(tmpDetail(j)) = 4 Then
                            'a3 et a2a3 => a2-a3
                            tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 2)

                        ElseIf Len(tmpBase(j)) = 2 And Len(tmpDetail(j)) = 5 Then
                            If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                'g4 et g3g4+ => g3-g4+
                                'g4 et g3g4# => g3-g4#
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 3)
                            Else
                                MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                            End If

                        ElseIf Len(tmpBase(j)) = 3 And Len(tmpDetail(j)) = 5 Then

                            If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                'a3+ et a2a3+ => a2-a3+
                                'a3# et a2a3# => a2-a3#
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 3)
                            Else
                                'Ca3 et Cb1a3 => Cb1-a3
                                tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 2)
                            End If

                        ElseIf Len(tmpBase(j)) = 3 And Len(tmpDetail(j)) = 6 Then
                            If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                'Rb8 et Rb7b8+ => Rb7-b8+
                                'Rb8 et Rb7b8# => Rb7-b8#
                                tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 3)
                            Else
                                MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                            End If

                        ElseIf Len(tmpBase(j)) = 4 And Len(tmpDetail(j)) = 4 Then
                            'axb3 et a2b3 => a2xb3
                            tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 2)

                        ElseIf Len(tmpBase(j)) = 4 And Len(tmpDetail(j)) = 5 Then

                            If InStr(tmpBase(j), "=") > 0 Then
                                'a8=D et a7a8D => a7-a8=D
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & tmpBase(j)
                            Else
                                If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                    If InStr(tmpBase(j), "+") > 0 Or InStr(tmpBase(j), "#") > 0 Then
                                        MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                    ElseIf separateur = "x" Then
                                        'fxg5 et f4g5+ => f4xg5+
                                        tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 3)
                                    Else
                                        MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                    End If
                                Else
                                    'Cxa3 et Cb1a3 => Cb1xa3
                                    'Cba3 et Cb1a3 => Cb1-a3
                                    tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 2)
                                End If
                            End If

                        ElseIf Len(tmpBase(j)) = 4 And Len(tmpDetail(j)) = 6 Then

                            If InStr(tmpDetail(j), "ep") > 0 Then
                                'axb3 et a2b3ep => a2xb3ep
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 4)
                            Else
                                If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                    If InStr(tmpBase(j), "+") > 0 Or InStr(tmpBase(j), "#") > 0 Then
                                        'Ca3+ et Cb1a3+ => Cb1-a3+
                                        'Ca3# et Cb1a3# => Cb1-a3#
                                        tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 3)
                                    ElseIf separateur = "x" Then
                                        'Cxa3 et Cb1a3+ => Cb1xa3+
                                        'Cxa3 et Cb1a3# => Cb1xa3#
                                        tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 3)
                                    ElseIf InStr(tmpBase(j), "=") > 0 Then
                                        'e8=Q et e7e8Q+ => e7-e8=Q+
                                        tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & tmpBase(j) & droite(tmpDetail(j), 1)
                                    Else
                                        MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                    End If
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            End If

                        ElseIf Len(tmpBase(j)) = 5 And Len(tmpDetail(j)) = 5 Then

                            If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                'axb3+ et a2b3+ => a2xb3+
                                'axb3# et a2b3# => a2xb3#
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 3)
                            Else
                                'Cbxa3 et Cb1a3 => Cb1xa3
                                tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 2)
                            End If

                        ElseIf Len(tmpBase(j)) = 5 And Len(tmpDetail(j)) = 6 Then

                            If InStr(tmpBase(j), "=") > 0 Then
                                If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                    'a8=D+ et a7a8D+ => a7-a8=D+
                                    'a8=D# et a7a8D# => a7-a8=D#
                                    tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & tmpBase(j)
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            Else
                                If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                    'Cxb3+ et Cb1a3+ => Cb1xa3+
                                    'Cxb3# et Cb1a3# => Cb1xa3#
                                    tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 3)
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            End If

                        ElseIf Len(tmpBase(j)) = 5 And Len(tmpDetail(j)) = 7 Then

                            If InStr(tmpDetail(j), "ep") > 0 Then
                                If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                    'axb3+ et a2b3ep+ => a3xb3ep+
                                    'axb3# et a2b3ep# => a3xb3ep#
                                    tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpDetail(j), 5)
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            Else
                                MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                            End If

                        ElseIf Len(tmpBase(j)) = 6 And Len(tmpDetail(j)) = 5 Then

                            If InStr(tmpBase(j), "=") > 0 Then
                                'axb8=D et a7b8D => a7xb8=D
                                tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpBase(j), 4)
                            Else
                                MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                            End If

                        ElseIf Len(tmpBase(j)) = 6 And Len(tmpDetail(j)) = 6 Then

                            If InStr(tmpDetail(j), "+") > 0 Or InStr(tmpDetail(j), "#") > 0 Then
                                'Cbxa3+ et Cb1a3+ => Cb1xa3+
                                'Cbxa3# et Cb1a3# => Cb1xa3#
                                tmpBase(j) = gauche(tmpDetail(j), 3) & separateur & droite(tmpDetail(j), 3)
                            Else
                                MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                            End If

                        ElseIf Len(tmpBase(j)) = 7 And Len(tmpDetail(j)) = 6 Then

                            If InStr(tmpBase(j), "=") > 0 Then
                                If InStr(tmpBase(j), "+") > 0 Or InStr(tmpBase(j), "#") > 0 Then
                                    'axb8=D+ et a7b8D+ => a7xb8=D+
                                    'axb8=D# et a7b8D# => a7xb8=D#
                                    tmpBase(j) = gauche(tmpDetail(j), 2) & separateur & droite(tmpBase(j), 5)
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            Else
                                If InStr(tmpBase(j), "+") > 0 Or InStr(tmpBase(j), "#") > 0 Then
                                    'plusieurs reines sur l'échiquier d'où la présence des coordonnées de départ
                                    'Qf2xf7+ et Qf2f7+ => Qf2xf7+
                                    'Qf2xf7# et Qf2f7# => Qf2xf7#
                                    'donc pas de modification
                                Else
                                    MsgBox("problem" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                                End If
                            End If
                        Else
                            MsgBox("contact Chris" & vbCrLf & "base = " & tmpBase(j) & vbCrLf & "detail = " & tmpDetail(j))
                        End If
                    End If
                Next
                tabBase(i) = String.Join(" ", tmpBase)
            End If
            If i Mod 100 = 0 Then
                Console.Clear()
                Console.WriteLine("4. detailled moves @ " & Format((i + 1) / tabBase.Length, "0%") & " (" & Format(DateAdd(DateInterval.Second, (tabBase.Length - (i + 1)) * ((Environment.TickCount - depart) / 1000) / (i + 1), Now), "HH'h'mm'm'ss") & ")")
                System.Threading.Thread.Sleep(10)
            End If
        Next

        Return String.Join(vbCrLf & vbCrLf, tabBase)
    End Function

    Public Function droite(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Right(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Public Function extensionFichier(chemin As String) As String
        Dim fichier As New FileInfo(chemin)
        Return fichier.Extension
    End Function

    Public Function gauche(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Left(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Private Function nettoyer(fichier As String) As String
        Dim chaine As String, tabChaine() As String, tmp() As String, i As Integer, j As Integer, depart As Integer

        'on charge le fichier PGN
        chaine = My.Computer.FileSystem.ReadAllText(fichier)
        chaine = Replace(chaine, "�", "é")
        chaine = Replace(chaine, "\""]", """]")

        If InStr(chaine, vbLf & vbLf) > 0 Then
            chaine = Replace(chaine, vbLf & vbLf, vbCr & vbCr)
            chaine = Replace(chaine, "]" & vbLf, "]" & vbCr)
            chaine = Replace(chaine, vbCrLf & vbCrLf, vbCr & vbCr)

            chaine = Replace(chaine, "]" & vbCrLf, "]" & vbCr)

            chaine = Replace(chaine, vbCrLf, vbLf)
            chaine = Replace(chaine, vbLf, " ")
            chaine = Replace(chaine, vbCr, vbCrLf)
        ElseIf InStr(chaine, "]" & vbCrLf) > 0 Then
            chaine = Replace(chaine, "]" & vbCrLf & vbCrLf, "]" & vbCr & vbCr)
            chaine = Replace(chaine, "]" & vbCrLf, "]" & vbCr)
            chaine = Replace(chaine, vbCrLf & vbCrLf, vbCr & vbCr)

            chaine = Replace(chaine, vbCrLf, vbLf)
            chaine = Replace(chaine, vbLf, " ")
            chaine = Replace(chaine, vbCr, vbCrLf)
        End If

        GC.Collect()
        tabChaine = Split(chaine, vbCrLf)

        'nettoyage fichier PGN
        chaine = ""
        GC.Collect()
        depart = Environment.TickCount
        For i = 0 To UBound(tabChaine)
            tabChaine(i) = Trim(tabChaine(i))
            'on nettoie les infos (entre accolades)
            While InStr(tabChaine(i), "{") > 0 And InStr(tabChaine(i), "}") > 0
                tabChaine(i) = Replace(tabChaine(i), tabChaine(i).Substring(tabChaine(i).IndexOf("{"), tabChaine(i).IndexOf("}") - tabChaine(i).IndexOf("{") + 1), "")
                tabChaine(i) = Replace(tabChaine(i), "  ", " ")
                System.Threading.Thread.Sleep(10)
            End While

            If InStr(tabChaine(i), "[Event ") = 1 _
            Or InStr(tabChaine(i), "[Site ") = 1 _
            Or InStr(tabChaine(i), "[Date ") = 1 _
            Or InStr(tabChaine(i), "[Round ") = 1 _
            Or InStr(tabChaine(i), "[White ") = 1 _
            Or InStr(tabChaine(i), "[Black ") = 1 _
            Or InStr(tabChaine(i), "[Result ") = 1 _
            Or InStr(tabChaine(i), "[FEN ") = 1 _
            Or InStr(tabChaine(i), "[PlyCount ") = 1 _
            Or InStr(tabChaine(i), "[SetUp ") = 1 _
            Or InStr(tabChaine(i), "[ECO ") = 1 _
            Or InStr(tabChaine(i), "[Opening ") = 1 _
            Or InStr(tabChaine(i), "[TimeControl ") = 1 _
            Or tabChaine(i) = "" _
            Or InStr(tabChaine(i), "1. ") = 1 _
            Or InStr(tabChaine(i), "1.") = 1 Then
                chaine = chaine & tabChaine(i) & vbCrLf
            End If
            If i Mod 100 = 0 Then
                Console.Clear()
                Console.WriteLine("2a. removing statistics @ " & Format((i + 1) / tabChaine.Length, "0%") & " (" & Format(DateAdd(DateInterval.Second, (tabChaine.Length - (i + 1)) * ((Environment.TickCount - depart) / 1000) / (i + 1), Now), "HH'h'mm'm'ss") & ")")
                System.Threading.Thread.Sleep(10)
            End If
        Next

        'on s'occupe des ouvertures
        If MsgBox("Keep players names ?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            'on efface le retour à la ligne superflu
            While chaine.LastIndexOf(vbCrLf) = chaine.Length - 2
                chaine = gauche(chaine, Len(chaine) - 2)
                System.Threading.Thread.Sleep(10)
            End While

            GC.Collect()
            tabChaine = Split(chaine, vbCrLf)

            chaine = ""
            GC.Collect()

            depart = Environment.TickCount
            For i = 0 To UBound(tabChaine) - 3
                If InStr(tabChaine(i), "[White ") = 1 Then
                    j = i + 1
                    While InStr(tabChaine(j), "1.") = 0
                        If InStr(tabChaine(j), "[Opening ") = 1 And InStr(tabChaine(j), "[Opening ""?""") = 0 Then
                            tmp = tabChaine(j).Split("""")
                            tabChaine(i) = "[White """ & tmp(1) & """]"
                            Exit While
                        Else
                            j = j + 1
                        End If
                        If j >= UBound(tabChaine) Then
                            Exit While
                        End If
                        System.Threading.Thread.Sleep(10)
                    End While
                    'on n'a pas trouvé "[Opening "
                    If InStr(tabChaine(j), "1.") > 0 Then
                        j = i + 1
                        While InStr(tabChaine(j), "1.") = 0
                            If InStr(tabChaine(j), "[ECO ") = 1 Then
                                tmp = tabChaine(j).Split("""")
                                tabChaine(i) = "[White """ & tmp(1) & """]"
                                Exit While
                            Else
                                j = j + 1
                            End If
                            If j >= UBound(tabChaine) Then
                                Exit While
                            End If
                            System.Threading.Thread.Sleep(10)
                        End While
                    End If
                ElseIf InStr(tabChaine(i), "[Black ") = 1 Then
                    j = i + 1
                    While InStr(tabChaine(j), "1.") = 0
                        If InStr(tabChaine(j), "[Opening ") = 1 And InStr(tabChaine(j), "[Opening ""?""") = 0 Then
                            tmp = tabChaine(j).Split("""")
                            tabChaine(i) = "[Black """ & tmp(1) & """]"
                            Exit While
                        Else
                            j = j + 1
                        End If
                        If j >= UBound(tabChaine) Then
                            Exit While
                        End If
                        System.Threading.Thread.Sleep(10)
                    End While
                    If InStr(tabChaine(j), "1.") > 0 Then
                        j = i + 1
                        While InStr(tabChaine(j), "1.") = 0
                            If InStr(tabChaine(j), "[ECO ") = 1 Then
                                tmp = tabChaine(j).Split("""")
                                tabChaine(i) = "[Black """ & tmp(1) & """]"
                                Exit While
                            Else
                                j = j + 1
                            End If
                            If j >= UBound(tabChaine) Then
                                Exit While
                            End If
                            System.Threading.Thread.Sleep(10)
                        End While
                    End If
                End If
                If i Mod 100 = 0 Then
                    Console.Clear()
                    Console.WriteLine("2b. cleaning headers @ " & Format((i + 1) / tabChaine.Length, "0%") & " (" & Format(DateAdd(DateInterval.Second, (tabChaine.Length - (i + 1)) * ((Environment.TickCount - depart) / 1000) / (i + 1), Now), "HH'h'mm'm'ss") & ")")
                    System.Threading.Thread.Sleep(10)
                End If
            Next

            GC.Collect()
            chaine = String.Join(vbCrLf, tabChaine)
        End If

        tabChaine = Nothing
        GC.Collect()

        'on efface le retour à la ligne superflu
        While chaine.LastIndexOf(vbCrLf) = chaine.Length - 2
            chaine = gauche(chaine, Len(chaine) - 2)
            System.Threading.Thread.Sleep(10)
        End While

        Return chaine
    End Function

    Public Function nomFichier(chemin As String) As String
        Return My.Computer.FileSystem.GetName(chemin)
    End Function

    Public Sub pgnDETAIL(fichier As String, suffixe As String)
        Dim nom As String, commande As New Process()
        Dim dossierFichier As String, dossierTravail As String

        nom = Replace(nomFichier(fichier), ".pgn", "")

        dossierFichier = fichier.Substring(0, fichier.LastIndexOf("\"))
        dossierTravail = Environment.CurrentDirectory

        'si pgn-extract.exe ne se trouve dans le même dossier que le notre application
        If Not My.Computer.FileSystem.FileExists("pgn-extract.exe") Then
            'on cherche s'il se trouve dans le même dossier que le fichierPGN
            dossierTravail = dossierFichier
            If Not My.Computer.FileSystem.FileExists(dossierFichier & "\pgn-extract.exe") Then
                dossierTravail = "D:\JEUX\ARENA CHESS 3.5.1\Databases\PGN Extract GUI"
                If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then
                    dossierTravail = "E:\JEUX\ARENA CHESS 3.5.1\Databases\PGN Extract GUI"
                    If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then
                        'pgn-extract.exe est introuvable
                        MsgBox("Veuillez copier pgn-extract.exe dans :" & vbCrLf & dossierTravail, MsgBoxStyle.Critical)
                        dossierTravail = Environment.CurrentDirectory
                        If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        'si le fichierPGN ne se trouve pas dans le dossier de travail
        If dossierFichier <> dossierTravail Then
            'on recopie temporairement le fichierPGN dans le dossierTravail
            My.Computer.FileSystem.CopyFile(fichier, dossierTravail & "\" & nom & ".pgn")
        End If

        'on écrit le chemin du fichierPGN dans le fichier .lst
        My.Computer.FileSystem.WriteAllText(dossierTravail & "\" & nom & ".lst", dossierTravail & "\" & nom & ".pgn", False, System.Text.Encoding.ASCII)

        commande.StartInfo.FileName = dossierTravail & "\pgn-extract.exe"
        commande.StartInfo.WorkingDirectory = dossierTravail
        commande.StartInfo.Arguments = " -WelalgPNBRQK -o" & nom & suffixe & ".pgn" & " -f" & nom & ".lst" & " -l" & nom & ".log"
        commande.StartInfo.UseShellExecute = False
        commande.Start()
        commande.WaitForExit()

        My.Computer.FileSystem.DeleteFile(dossierTravail & "\" & nom & ".lst")

        'si le dossierTravail ne correspond pas au dossier du fichierPGN
        If dossierFichier <> dossierTravail Then
            'on déplace le fichier _detailled.pgn
            My.Computer.FileSystem.DeleteFile(dossierTravail & "\" & nom & ".pgn")
            My.Computer.FileSystem.MoveFile(dossierTravail & "\" & nom & suffixe & ".pgn", dossierFichier & "\" & nom & suffixe & ".pgn")
            My.Computer.FileSystem.MoveFile(dossierTravail & "\" & nom & ".log", dossierFichier & "\" & nom & ".log")
        End If

    End Sub


End Module
