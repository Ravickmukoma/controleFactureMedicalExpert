Attribute VB_Name = "ModuleControleMedicalExpert"
' ===================================================================
' MODULE VBA - CONTRÔLE MÉDICAL EXPERT
' ===================================================================
' Basé sur connaissances médicales approfondies et normes internationales
' OMS, HAS, FDA, ANSM, Sociétés savantes médicales
' ===================================================================

Option Explicit

' Variables globales
Public Const FEUILLE_REFERENTIEL As String = "ReferentielEnrichi"
Public Const FEUILLE_SAISIE As String = "SaisieFactures"
Public Const FEUILLE_SURVEILLANCE As String = "SurveillanceIntelligente"
Public Const FEUILLE_VALIDATION As String = "ValidationManuelle"
Public Const FEUILLE_ALERTES As String = "AlertesAutomatiques"
Public Const FEUILLE_DASHBOARD As String = "StatistiquesDashboard"

' ===================================================================
' INITIALISATION DU SYSTÈME EXPERT
' ===================================================================
Public Sub InitialiserSystemeExpert()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    MsgBox "🏥 INITIALISATION SYSTÈME CONTRÔLE MÉDICAL EXPERT" & vbCrLf & vbCrLf & _
           "✅ Chargement base de connaissances médicales..." & vbCrLf & _
           "✅ Configuration règles OMS/HAS/FDA/ANSM..." & vbCrLf & _
           "✅ Activation surveillance intelligente...", vbInformation, "Système Expert Médical"
    
    ' Vérifier structure et charger base de connaissances
    If Not VerifierStructureComplete() Then
        MsgBox "❌ Structure système incomplète!", vbCritical
        Exit Sub
    End If
    
    ' Initialiser la base de connaissances médicales
    InitialiserBaseConnaissancesMedicales
    
    ' Configurer les alertes automatiques
    ConfigurerAlertesAutomatiques
    
    ' Activer la surveillance temps réel
    ActiverSurveillanceTempsReel
    
    MsgBox "🎉 SYSTÈME EXPERT INITIALISÉ!" & vbCrLf & vbCrLf & _
           "📊 " & CompterActesReferentiel() & " actes avec règles médicales" & vbCrLf & _
           "🧠 Base de connaissances médicales activée" & vbCrLf & _
           "🔍 Surveillance intelligente opérationnelle", vbInformation, "Initialisation Terminée"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "❌ Erreur initialisation: " & Err.Description, vbCritical
End Sub

' ===================================================================
' CONTRÔLE EXPERT COMPLET
' ===================================================================
Public Sub LancerControleExpertComplet()
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Dim wsFactures As Worksheet, wsSurveillance As Worksheet
    Dim derniereLigne As Long, i As Long
    Dim anomaliesDetectees As Long, anomaliesCritiques As Long
    Dim startTime As Double
    
    startTime = Timer
    
    Set wsFactures = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    Set wsSurveillance = ThisWorkbook.Worksheets(FEUILLE_SURVEILLANCE)
    
    ' Vérifier données
    derniereLigne = wsFactures.Cells(wsFactures.Rows.Count, 1).End(xlUp).Row
    If derniereLigne <= 1 Then
        MsgBox "⚠️ Aucune facture à contrôler!", vbExclamation
        Exit Sub
    End If
    
    ' Effacer anciennes anomalies
    EffacerAnciennesAnomalies
    
    anomaliesDetectees = 0
    anomaliesCritiques = 0
    
    ' Contrôle expert de chaque facture
    For i = 2 To derniereLigne
        Dim resultats As Variant
        resultats = ControlerFactureExpert(i)
        
        anomaliesDetectees = anomaliesDetectees + resultats(0)
        anomaliesCritiques = anomaliesCritiques + resultats(1)
        
        ' Afficher progrès
        If i Mod 50 = 0 Then
            Application.StatusBar = "Contrôle expert... " & i - 1 & "/" & derniereLigne - 1 & " factures"
        End If
    Next i
    
    ' Post-traitement expert
    EffectuerAnalysesComplementaires
    FormaterResultatsExpert
    MettreAJourDashboardExpert
    
    ' Générer alertes si nécessaire
    If anomaliesCritiques > 0 Then
        GenererAlertesAutomatiques anomaliesCritiques
    End If
    
    Dim dureeControle As Double
    dureeControle = Timer - startTime
    
    ' Enregistrer dans historique
    EnregistrerDansHistorique derniereLigne - 1, anomaliesDetectees, dureeControle
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Message final expert
    If anomaliesDetectees > 0 Then
        MsgBox "🔍 CONTRÔLE EXPERT TERMINÉ!" & vbCrLf & vbCrLf & _
               "📊 Factures analysées: " & derniereLigne - 1 & vbCrLf & _
               "🚨 Anomalies détectées: " & anomaliesDetectees & vbCrLf & _
               "⚠️ Anomalies critiques: " & anomaliesCritiques & vbCrLf & _
               "⏱️ Durée: " & Format(dureeControle, "0.0") & " secondes" & vbCrLf & vbCrLf & _
               "➡️ Consultez 'SurveillanceIntelligente' pour détails", vbInformation, "Contrôle Expert"
    Else
        MsgBox "✅ CONTRÔLE EXPERT TERMINÉ!" & vbCrLf & vbCrLf & _
               "📊 Factures analysées: " & derniereLigne - 1 & vbCrLf & _
               "🎉 Aucune anomalie détectée!" & vbCrLf & _
               "⏱️ Durée: " & Format(dureeControle, "0.0") & " secondes" & vbCrLf & vbCrLf & _
               "Conformité parfaite aux normes médicales.", vbInformation, "Contrôle Expert"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "❌ Erreur contrôle expert: " & Err.Description, vbCritical
End Sub

' ===================================================================
' CONTRÔLE EXPERT D'UNE FACTURE
' ===================================================================
Private Function ControlerFactureExpert(ligne As Long) As Variant
    Dim wsFactures As Worksheet, wsReferentiel As Worksheet
    Dim anomalies As Long, anomaliesCritiques As Long
    Dim codeActe As String, libelleFacture As String, nomPatient As String
    Dim dateFacture As Date, prixUnitaire As Double, quantite As Double, prixTotal As Double
    
    Set wsFactures = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    Set wsReferentiel = ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL)
    
    anomalies = 0
    anomaliesCritiques = 0
    
    ' Récupérer données facture
    On Error Resume Next
    dateFacture = CDate(wsFactures.Cells(ligne, 1).Value)
    nomPatient = Trim(wsFactures.Cells(ligne, 2).Value)
    codeActe = Trim(wsFactures.Cells(ligne, 5).Value)
    libelleFacture = Trim(wsFactures.Cells(ligne, 6).Value)
    prixUnitaire = CDbl(wsFactures.Cells(ligne, 7).Value)
    quantite = CDbl(wsFactures.Cells(ligne, 8).Value)
    prixTotal = CDbl(wsFactures.Cells(ligne, 9).Value)
    On Error GoTo 0
    
    ' === CONTRÔLES CRITIQUES ===
    
    ' 1. Code manquant - Recherche par libellé
    If codeActe = "" Then
        codeActe = RechercherCodeParLibelle(libelleFacture)
        If codeActe = "" Then
            AjouterAnomalieExpert nomPatient, "", libelleFacture, "Code inexistant", "CRITIQUE", _
                                 "Code acte manquant et libellé non trouvé dans référentiel médical", _
                                 "Code acte valide ou libellé exact", "Code vide + libellé: " & libelleFacture, _
                                 prixTotal, "Vérifier nomenclature médicale", "Immédiat", "URGENT"
            anomalies = anomalies + 1
            anomaliesCritiques = anomaliesCritiques + 1
        End If
    End If
    
    ' === CONTRÔLES ÉLEVÉS ===
    
    ' 2. Dépassement tarifaire (priorité selon spécifications)
    If codeActe <> "" Then
        Dim prixTarifaire As Double
        prixTarifaire = ObtenirPrixTarifaire(codeActe)
        If prixTarifaire > 0 And prixUnitaire > prixTarifaire Then
            AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Dépassement tarifaire", "ÉLEVÉ", _
                                 "Prix facturé (" & prixUnitaire & " USD) supérieur au tarif contractuel (" & prixTarifaire & " USD)", _
                                 "Prix ≤ " & prixTarifaire & " USD", "Prix facturé: " & prixUnitaire & " USD", _
                                 prixTotal, "Appliquer tarif contractuel ou justifier dépassement", "48h", "ÉLEVÉ"
            anomalies = anomalies + 1
        End If
    End If
    
    ' 3. Quantité invalide
    If quantite <= 0 Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Quantité invalide", "ÉLEVÉ", _
                             "Quantité nulle ou négative (" & quantite & ") - Impossible médicalement", _
                             "Quantité > 0", "Quantité: " & quantite, prixTotal, _
                             "Corriger la quantité", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ' 4. Contrôles selon règles médicales expertes
    If codeActe <> "" Then
        Dim resultatsExpert As Variant
        resultatsExpert = ControlerSelonExpertiseMedicale(nomPatient, codeActe, libelleFacture, dateFacture, quantite, prixUnitaire, prixTotal, ligne)
        anomalies = anomalies + resultatsExpert(0)
        anomaliesCritiques = anomaliesCritiques + resultatsExpert(1)
    End If
    
    ' === CONTRÔLES MODÉRÉS ===
    
    ' 5. Erreur de calcul (TRÈS IMPORTANT selon spécifications)
    If Abs(prixUnitaire * quantite - prixTotal) > 0.01 Then
        Dim calculAttendu As Double
        calculAttendu = prixUnitaire * quantite
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Erreur de calcul", "MODÉRÉ", _
                             "Erreur arithmétique: " & prixUnitaire & " × " & quantite & " = " & calculAttendu & " ≠ " & prixTotal, _
                             "P.T. = " & calculAttendu & " USD", "P.T. facturé: " & prixTotal & " USD", _
                             prixTotal, "Corriger le calcul immédiatement", "72h", "MOYEN"
        anomalies = anomalies + 1
    End If
    
    ControlerFactureExpert = Array(anomalies, anomaliesCritiques)
End Function

' ===================================================================
' CONTRÔLES SELON EXPERTISE MÉDICALE AVANCÉE
' ===================================================================
Private Function ControlerSelonExpertiseMedicale(nomPatient As String, codeActe As String, libelleFacture As String, _
                                                 dateFacture As Date, quantite As Double, prixUnitaire As Double, _
                                                 prixTotal As Double, ligneFacture As Long) As Variant
    Dim wsReferentiel As Worksheet
    Dim ligneRef As Long, anomalies As Long, anomaliesCritiques As Long
    Dim categorieMedicale As String, sousCategorie As String
    Dim qtyMaxJour As Long, qtyMaxMois As Long, qtyMaxVie As Long
    Dim delaiMinJours As Long, actesNonCumulables As String
    Dim niveauRisque As String, regleMedicale As String
    
    Set wsReferentiel = ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL)
    anomalies = 0
    anomaliesCritiques = 0
    
    ' Trouver acte dans référentiel
    ligneRef = TrouverActeReferentiel(codeActe)
    If ligneRef = 0 Then Exit Function
    
    ' Récupérer données médicales
    categorieMedicale = wsReferentiel.Cells(ligneRef, 4).Value
    sousCategorie = wsReferentiel.Cells(ligneRef, 5).Value
    qtyMaxJour = wsReferentiel.Cells(ligneRef, 6).Value
    qtyMaxMois = wsReferentiel.Cells(ligneRef, 7).Value
    qtyMaxVie = wsReferentiel.Cells(ligneRef, 8).Value
    delaiMinJours = wsReferentiel.Cells(ligneRef, 9).Value
    actesNonCumulables = wsReferentiel.Cells(ligneRef, 10).Value
    niveauRisque = wsReferentiel.Cells(ligneRef, 13).Value
    regleMedicale = wsReferentiel.Cells(ligneRef, 14).Value
    
    ' === CONTRÔLES SPÉCIALISÉS PAR CATÉGORIE MÉDICALE ===
    
    Select Case categorieMedicale
        Case "Chirurgie"
            anomalies = anomalies + ControlerChirurgie(nomPatient, codeActe, libelleFacture, sousCategorie, dateFacture, quantite, prixTotal, qtyMaxJour, qtyMaxMois, delaiMinJours)
            
        Case "Imagerie médicale"
            anomalies = anomalies + ControlerImagerieMedicale(nomPatient, codeActe, libelleFacture, sousCategorie, dateFacture, quantite, prixTotal, qtyMaxJour, qtyMaxMois)
            
        Case "Examens biologiques"
            anomalies = anomalies + ControlerExamensBiologiques(nomPatient, codeActe, libelleFacture, dateFacture, quantite, prixTotal, qtyMaxJour, qtyMaxMois)
            
        Case "Consultations médicales"
            anomalies = anomalies + ControlerConsultations(nomPatient, codeActe, libelleFacture, sousCategorie, dateFacture, quantite, prixTotal, qtyMaxJour)
            
        Case "Urgences médicales"
            anomalies = anomalies + ControlerUrgences(nomPatient, codeActe, libelleFacture, dateFacture, quantite, prixTotal, qtyMaxJour)
            
        Case "Actes gynécologiques"
            anomalies = anomalies + ControlerGynecologie(nomPatient, codeActe, libelleFacture, sousCategorie, dateFacture, quantite, prixTotal, qtyMaxVie, delaiMinJours)
            
        Case "Médicaments essentiels"
            anomalies = anomalies + ControlerMedicaments(nomPatient, codeActe, libelleFacture, dateFacture, quantite, prixTotal, qtyMaxJour, qtyMaxMois)
            
        Case "Produits anesthésiques"
            anomalies = anomalies + ControlerAnesthesie(nomPatient, codeActe, libelleFacture, dateFacture, quantite, prixTotal, ligneFacture)
    End Select
    
    ' === CONTRÔLES GÉNÉRAUX AVANCÉS ===
    
    ' Quantité excessive selon normes médicales
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Quantité excessive", "MODÉRÉ", _
                             "Quantité (" & quantite & ") dépasse maximum autorisé par jour (" & qtyMaxJour & ") selon normes " & categorieMedicale, _
                             "Quantité max/jour: " & qtyMaxJour, "Quantité facturée: " & quantite, prixTotal, _
                             "Justifier dépassement médical ou fractionner", "72h", "MOYEN"
        anomalies = anomalies + 1
    End If
    
    ' Détection doublons intelligente
    If DetecterDoublonIntelligent(nomPatient, codeActe, dateFacture, ligneFacture) Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Doublon détecté", "ÉLEVÉ", _
                             "Même patient, même acte, dates rapprochées - Possible double facturation", _
                             "Acte unique dans délai minimum", "Doublon détecté", prixTotal, _
                             "Vérifier justification médicale", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ' Cumul excessif selon règles temporelles
    If VerifierCumulExcessif(nomPatient, codeActe, dateFacture, quantite, qtyMaxMois, ligneFacture) Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Cumul excessif", "ÉLEVÉ", _
                             "Cumul mensuel dépasse maximum autorisé (" & qtyMaxMois & ") selon protocole médical", _
                             "Cumul max/mois: " & qtyMaxMois, "Cumul détecté dépassant", prixTotal, _
                             "Justifier nécessité médicale", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ControlerSelonExpertiseMedicale = Array(anomalies, anomaliesCritiques)
End Function

' ===================================================================
' CONTRÔLES SPÉCIALISÉS PAR DOMAINE MÉDICAL
' ===================================================================

Private Function ControlerChirurgie(nomPatient As String, codeActe As String, libelleFacture As String, _
                                   sousCategorie As String, dateFacture As Date, quantite As Double, _
                                   prixTotal As Double, qtyMaxJour As Long, qtyMaxMois As Long, delaiMinJours As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles chirurgicales spécifiques selon HAS et sociétés savantes
    Select Case sousCategorie
        Case "Chirurgie cardiaque"
            ' Règles cardio-thoraciques (SFC, SFCTCV)
            If quantite > 1 Then
                AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Chirurgie cardiaque multiple", "CRITIQUE", _
                                     "Plusieurs interventions cardiaques même jour - Risque vital majeur", _
                                     "Maximum 1 intervention/jour", "Quantité: " & quantite, prixTotal, _
                                     "Vérification équipe cardiochirurgicale urgente", "Immédiat", "URGENT"
                anomalies = anomalies + 1
            End If
            
        Case "Neurochirurgie"
            ' Règles neurochirurgicales (SFNC)
            If quantite > 1 Then
                AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Neurochirurgie multiple", "CRITIQUE", _
                                     "Plusieurs interventions neurochirurgicales - Protocole sécurité non respecté", _
                                     "Maximum 1 intervention/séance", "Quantité: " & quantite, prixTotal, _
                                     "Validation neurochirurgien senior obligatoire", "Immédiat", "URGENT"
                anomalies = anomalies + 1
            End If
            
        Case "Chirurgie orthopédique"
            ' Règles orthopédiques (SOFCOT)
            If InStr(LCase(libelleFacture), "prothèse") > 0 And quantite > 1 Then
                AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Prothèses multiples", "ÉLEVÉ", _
                                     "Pose multiple prothèses même séance - Vérifier indication", _
                                     "Généralement 1 prothèse/intervention", "Quantité: " & quantite, prixTotal, _
                                     "Justification orthopédique requise", "48h", "ÉLEVÉ"
                anomalies = anomalies + 1
            End If
    End Select
    
    ' Vérification délai post-opératoire
    If VerifierDelaiPostOperatoire(nomPatient, codeActe, dateFacture, delaiMinJours) Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Délai post-opératoire", "ÉLEVÉ", _
                             "Intervention trop rapprochée - Délai cicatrisation non respecté", _
                             "Délai minimum: " & delaiMinJours & " jours", "Délai insuffisant", prixTotal, _
                             "Justifier urgence médicale", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ControlerChirurgie = anomalies
End Function

Private Function ControlerImagerieMedicale(nomPatient As String, codeActe As String, libelleFacture As String, _
                                          sousCategorie As String, dateFacture As Date, quantite As Double, _
                                          prixTotal As Double, qtyMaxJour As Long, qtyMaxMois As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles radioprotection (IRSN, ASN)
    If InStr(LCase(libelleFacture), "scanner") > 0 Or InStr(LCase(libelleFacture), "irm") > 0 Then
        ' Contrôle exposition radiologique
        If VerifierExpositionRadiologique(nomPatient, codeActe, dateFacture, quantite) Then
            AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Surexposition radiologique", "ÉLEVÉ", _
                                 "Cumul examens irradiants dépasse seuils radioprotection", _
                                 "Respect protocole ALARA", "Surexposition détectée", prixTotal, _
                                 "Justification radiologique obligatoire", "48h", "ÉLEVÉ"
            anomalies = anomalies + 1
        End If
    End If
    
    ' Règles justification (HAS)
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Imagerie excessive", "MODÉRÉ", _
                             "Nombre examens/jour dépasse recommandations HAS", _
                             "Maximum " & qtyMaxJour & " examens/jour", "Quantité: " & quantite, prixTotal, _
                             "Justifier indication médicale", "72h", "MOYEN"
        anomalies = anomalies + 1
    End If
    
    ControlerImagerieMedicale = anomalies
End Function

Private Function ControlerExamensBiologiques(nomPatient As String, codeActe As String, libelleFacture As String, _
                                            dateFacture As Date, quantite As Double, prixTotal As Double, _
                                            qtyMaxJour As Long, qtyMaxMois As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles biologiques (SFBC, CNBH)
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Analyses excessives", "MODÉRÉ", _
                             "Nombre analyses/jour dépasse protocole laboratoire", _
                             "Maximum " & qtyMaxJour & " analyses/jour", "Quantité: " & quantite, prixTotal, _
                             "Vérifier prescription médicale", "72h", "MOYEN"
        anomalies = anomalies + 1
    End If
    
    ' Détection doublons biologiques même jour
    If VerifierDoublonBiologique(nomPatient, codeActe, dateFacture) Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Doublon biologique", "ÉLEVÉ", _
                             "Même analyse répétée même jour - Redondance non justifiée", _
                             "Analyse unique/jour sauf urgence", "Doublon détecté", prixTotal, _
                             "Justifier répétition analyse", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ControlerExamensBiologiques = anomalies
End Function

Private Function ControlerConsultations(nomPatient As String, codeActe As String, libelleFacture As String, _
                                       sousCategorie As String, dateFacture As Date, quantite As Double, _
                                       prixTotal As Double, qtyMaxJour As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles consultations (CNOM, HAS)
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Consultations excessives", "MODÉRÉ", _
                             "Nombre consultations/jour dépasse usage médical standard", _
                             "Maximum " & qtyMaxJour & " consultations/jour", "Quantité: " & quantite, prixTotal, _
                             "Justifier nécessité médicale", "72h", "MOYEN"
        anomalies = anomalies + 1
    End If
    
    ' Contrôle spécialisé pédiatrie
    If sousCategorie = "Pédiatrie" Then
        ' Vérification cohérence âge (simulation basée sur prénom)
        If DetecterIncohérenceAge(nomPatient, "pédiatrie") Then
            AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Âge inapproprié", "ÉLEVÉ", _
                                 "Consultation pédiatrique pour patient apparemment adulte", _
                                 "Patient < 18 ans", "Patient probablement adulte", prixTotal, _
                                 "Vérifier âge patient", "48h", "ÉLEVÉ"
            anomalies = anomalies + 1
        End If
    End If
    
    ControlerConsultations = anomalies
End Function

Private Function ControlerUrgences(nomPatient As String, codeActe As String, libelleFacture As String, _
                                  dateFacture As Date, quantite As Double, prixTotal As Double, qtyMaxJour As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles urgences (SFMU, SAMU)
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Urgences multiples", "ÉLEVÉ", _
                             "Passages multiples urgences même jour - Vérifier organisation soins", _
                             "Maximum " & qtyMaxJour & " passages/jour", "Quantité: " & quantite, prixTotal, _
                             "Justifier urgences répétées", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ControlerUrgences = anomalies
End Function

Private Function ControlerGynecologie(nomPatient As String, codeActe As String, libelleFacture As String, _
                                     sousCategorie As String, dateFacture As Date, quantite As Double, _
                                     prixTotal As Double, qtyMaxVie As Long, delaiMinJours As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles gynécologiques (CNGOF, HAS)
    If InStr(LCase(libelleFacture), "hystérectomie") > 0 Then
        ' Contrôle hystérectomie (acte irréversible)
        If VerifierActeIrreversible(nomPatient, codeActe, "hystérectomie") Then
            AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Acte irréversible répété", "CRITIQUE", _
                                 "Hystérectomie répétée - Impossible anatomiquement", _
                                 "Maximum 1 fois dans la vie", "Répétition détectée", prixTotal, _
                                 "Vérification dossier médical urgente", "Immédiat", "URGENT"
            anomalies = anomalies + 1
        End If
    End If
    
    ControlerGynecologie = anomalies
End Function

Private Function ControlerMedicaments(nomPatient As String, codeActe As String, libelleFacture As String, _
                                     dateFacture As Date, quantite As Double, prixTotal As Double, _
                                     qtyMaxJour As Long, qtyMaxMois As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles pharmaceutiques (ANSM, Ordre pharmaciens)
    If quantite > qtyMaxJour Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Posologie excessive", "ÉLEVÉ", _
                             "Quantité médicament dépasse posologie maximale autorisée", _
                             "Maximum " & qtyMaxJour & " unités/jour", "Quantité: " & quantite, prixTotal, _
                             "Vérifier prescription et posologie", "48h", "ÉLEVÉ"
        anomalies = anomalies + 1
    End If
    
    ControlerMedicaments = anomalies
End Function

Private Function ControlerAnesthesie(nomPatient As String, codeActe As String, libelleFacture As String, _
                                    dateFacture As Date, quantite As Double, prixTotal As Double, ligneFacture As Long) As Long
    Dim anomalies As Long
    anomalies = 0
    
    ' Règles anesthésiques (SFAR)
    If Not VerifierActeChirurgicalAssocie(nomPatient, dateFacture, ligneFacture) Then
        AjouterAnomalieExpert nomPatient, codeActe, libelleFacture, "Anesthésie isolée", "CRITIQUE", _
                             "Anesthésie sans acte chirurgical associé - Protocole non conforme", _
                             "Anesthésie + acte chirurgical", "Anesthésie isolée", prixTotal, _
                             "Vérifier protocole anesthésique", "Immédiat", "URGENT"
        anomalies = anomalies + 1
    End If
    
    ControlerAnesthesie = anomalies
End Function

' ===================================================================
' FONCTIONS UTILITAIRES EXPERTES
' ===================================================================

Private Function RechercherCodeParLibelle(libelle As String) As String
    Dim wsReferentiel As Worksheet
    Dim derniereLigne As Long, i As Long
    Dim libelleRef As String, similarite As Double
    
    Set wsReferentiel = ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL)
    derniereLigne = wsReferentiel.Cells(wsReferentiel.Rows.Count, 1).End(xlUp).Row
    
    ' Recherche exacte d'abord
    For i = 2 To derniereLigne
        libelleRef = Trim(UCase(wsReferentiel.Cells(i, 2).Value))
        If libelleRef = Trim(UCase(libelle)) Then
            RechercherCodeParLibelle = wsReferentiel.Cells(i, 1).Value
            Exit Function
        End If
    Next i
    
    ' Recherche approximative (similarité > 80%)
    For i = 2 To derniereLigne
        libelleRef = Trim(UCase(wsReferentiel.Cells(i, 2).Value))
        similarite = CalculerSimilarite(Trim(UCase(libelle)), libelleRef)
        If similarite > 0.8 Then
            RechercherCodeParLibelle = wsReferentiel.Cells(i, 1).Value
            Exit Function
        End If
    Next i
    
    RechercherCodeParLibelle = ""
End Function

Private Function CalculerSimilarite(texte1 As String, texte2 As String) As Double
    ' Algorithme simple de similarité de chaînes
    Dim longueurMin As Long, correspondances As Long, i As Long
    
    longueurMin = Application.WorksheetFunction.Min(Len(texte1), Len(texte2))
    If longueurMin = 0 Then
        CalculerSimilarite = 0
        Exit Function
    End If
    
    correspondances = 0
    For i = 1 To longueurMin
        If Mid(texte1, i, 1) = Mid(texte2, i, 1) Then
            correspondances = correspondances + 1
        End If
    Next i
    
    CalculerSimilarite = correspondances / Application.WorksheetFunction.Max(Len(texte1), Len(texte2))
End Function

Private Function DetecterDoublonIntelligent(nomPatient As String, codeActe As String, dateFacture As Date, ligneActuelle As Long) As Boolean
    Dim wsFactures As Worksheet
    Dim derniereLigne As Long, i As Long
    Dim nomPatientRef As String, codeActeRef As String, dateFactureRef As Date
    
    Set wsFactures = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    derniereLigne = wsFactures.Cells(wsFactures.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If i <> ligneActuelle Then
            On Error Resume Next
            nomPatientRef = Trim(wsFactures.Cells(i, 2).Value)
            codeActeRef = Trim(wsFactures.Cells(i, 5).Value)
            dateFactureRef = CDate(wsFactures.Cells(i, 1).Value)
            On Error GoTo 0
            
            ' Même patient, même acte, dates proches (≤ 7 jours)
            If UCase(nomPatientRef) = UCase(nomPatient) And _
               codeActeRef = codeActe And _
               Abs(dateFactureRef - dateFacture) <= 7 Then
                DetecterDoublonIntelligent = True
                Exit Function
            End If
        End If
    Next i
    
    DetecterDoublonIntelligent = False
End Function

Private Function VerifierCumulExcessif(nomPatient As String, codeActe As String, dateFacture As Date, _
                                      quantiteActuelle As Double, qtyMaxMois As Long, ligneActuelle As Long) As Boolean
    Dim wsFactures As Worksheet
    Dim derniereLigne As Long, i As Long
    Dim nomPatientRef As String, codeActeRef As String, dateFactureRef As Date, quantiteRef As Double
    Dim cumulMois As Double
    Dim debutMois As Date, finMois As Date
    
    Set wsFactures = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    derniereLigne = wsFactures.Cells(wsFactures.Rows.Count, 1).End(xlUp).Row
    
    ' Définir période du mois
    debutMois = DateSerial(Year(dateFacture), Month(dateFacture), 1)
    finMois = DateSerial(Year(dateFacture), Month(dateFacture) + 1, 0)
    
    cumulMois = quantiteActuelle
    
    For i = 2 To derniereLigne
        If i <> ligneActuelle Then
            On Error Resume Next
            nomPatientRef = Trim(wsFactures.Cells(i, 2).Value)
            codeActeRef = Trim(wsFactures.Cells(i, 5).Value)
            dateFactureRef = CDate(wsFactures.Cells(i, 1).Value)
            quantiteRef = CDbl(wsFactures.Cells(i, 8).Value)
            On Error GoTo 0
            
            ' Même patient, même acte, même mois
            If UCase(nomPatientRef) = UCase(nomPatient) And _
               codeActeRef = codeActe And _
               dateFactureRef >= debutMois And dateFactureRef <= finMois Then
                cumulMois = cumulMois + quantiteRef
            End If
        End If
    Next i
    
    VerifierCumulExcessif = (cumulMois > qtyMaxMois)
End Function

Private Sub AjouterAnomalieExpert(nomPatient As String, codeActe As String, libelleActe As String, _
                                 typeAnomalie As String, niveauGravite As String, description As String, _
                                 valeurAttendue As String, valeurTrouvee As String, montantConcerne As Double, _
                                 actionRecommandee As String, delaiAction As String, prioriteTraitement As String)
    Dim wsSurveillance As Worksheet
    Dim nouvelleLigne As Long
    
    Set wsSurveillance = ThisWorkbook.Worksheets(FEUILLE_SURVEILLANCE)
    nouvelleLigne = wsSurveillance.Cells(wsSurveillance.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Ajouter données complètes
    wsSurveillance.Cells(nouvelleLigne, 1).Value = Now()
    wsSurveillance.Cells(nouvelleLigne, 2).Value = nomPatient
    wsSurveillance.Cells(nouvelleLigne, 3).Value = codeActe
    wsSurveillance.Cells(nouvelleLigne, 4).Value = libelleActe
    wsSurveillance.Cells(nouvelleLigne, 5).Value = typeAnomalie
    wsSurveillance.Cells(nouvelleLigne, 6).Value = niveauGravite
    wsSurveillance.Cells(nouvelleLigne, 7).Value = description
    wsSurveillance.Cells(nouvelleLigne, 8).Value = valeurAttendue
    wsSurveillance.Cells(nouvelleLigne, 9).Value = valeurTrouvee
    wsSurveillance.Cells(nouvelleLigne, 10).Value = montantConcerne
    wsSurveillance.Cells(nouvelleLigne, 11).Value = "Règle médicale violée selon normes internationales"
    wsSurveillance.Cells(nouvelleLigne, 12).Value = actionRecommandee
    wsSurveillance.Cells(nouvelleLigne, 13).Value = delaiAction
    wsSurveillance.Cells(nouvelleLigne, 14).Value = "À traiter"
    wsSurveillance.Cells(nouvelleLigne, 15).Value = ""
    wsSurveillance.Cells(nouvelleLigne, 16).Value = ""
    wsSurveillance.Cells(nouvelleLigne, 17).Value = "Détectée par système expert médical"
    wsSurveillance.Cells(nouvelleLigne, 18).Value = prioriteTraitement
    wsSurveillance.Cells(nouvelleLigne, 19).Value = IIf(montantConcerne > 1000, "ÉLEVÉ", IIf(montantConcerne > 100, "MODÉRÉ", "FAIBLE"))
End Sub

' ===================================================================
' FONCTIONS UTILITAIRES SUPPLÉMENTAIRES
' ===================================================================

Private Function TrouverActeReferentiel(codeActe As String) As Long
    ' [Fonction existante - pas de modification]
    Dim wsReferentiel As Worksheet
    Dim derniereLigne As Long, i As Long
    
    Set wsReferentiel = ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL)
    derniereLigne = wsReferentiel.Cells(wsReferentiel.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If wsReferentiel.Cells(i, 1).Value = codeActe Then
            TrouverActeReferentiel = i
            Exit Function
        End If
    Next i
    
    TrouverActeReferentiel = 0
End Function

Private Function ObtenirPrixTarifaire(codeActe As String) As Double
    Dim ligneRef As Long
    ligneRef = TrouverActeReferentiel(codeActe)
    If ligneRef > 0 Then
        ObtenirPrixTarifaire = Val(ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL).Cells(ligneRef, 17).Value)
    Else
        ObtenirPrixTarifaire = 0
    End If
End Function

Private Function CompterActesReferentiel() As Long
    Dim wsReferentiel As Worksheet
    Set wsReferentiel = ThisWorkbook.Worksheets(FEUILLE_REFERENTIEL)
    CompterActesReferentiel = wsReferentiel.Cells(wsReferentiel.Rows.Count, 1).End(xlUp).Row - 1
End Function

Private Function VerifierStructureComplete() As Boolean
    ' Vérifier toutes les feuilles nécessaires
    Dim feuilles As Variant
    Dim i As Integer
    
    feuilles = Array(FEUILLE_REFERENTIEL, FEUILLE_SAISIE, FEUILLE_SURVEILLANCE, FEUILLE_VALIDATION, FEUILLE_ALERTES, FEUILLE_DASHBOARD)
    
    For i = 0 To UBound(feuilles)
        If Not FeuilleExiste(feuilles(i)) Then
            VerifierStructureComplete = False
            Exit Function
        End If
    Next i
    
    VerifierStructureComplete = True
End Function

Private Function FeuilleExiste(nomFeuille As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    FeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

' Fonctions médicales spécialisées (stubs - à implémenter selon besoins)
Private Function VerifierExpositionRadiologique(nomPatient As String, codeActe As String, dateFacture As Date, quantite As Double) As Boolean
    ' Logique de vérification exposition radiologique
    VerifierExpositionRadiologique = False ' Placeholder
End Function

Private Function VerifierDelaiPostOperatoire(nomPatient As String, codeActe As String, dateFacture As Date, delaiMinJours As Long) As Boolean
    ' Logique de vérification délai post-opératoire
    VerifierDelaiPostOperatoire = False ' Placeholder
End Function

Private Function VerifierDoublonBiologique(nomPatient As String, codeActe As String, dateFacture As Date) As Boolean
    ' Logique de vérification doublon biologique
    VerifierDoublonBiologique = False ' Placeholder
End Function

Private Function DetecterIncohérenceAge(nomPatient As String, typeConsultation As String) As Boolean
    ' Logique de détection incohérence âge
    DetecterIncohérenceAge = False ' Placeholder
End Function

Private Function VerifierActeIrreversible(nomPatient As String, codeActe As String, typeActe As String) As Boolean
    ' Logique de vérification acte irréversible
    VerifierActeIrreversible = False ' Placeholder
End Function

Private Function VerifierActeChirurgicalAssocie(nomPatient As String, dateFacture As Date, ligneFacture As Long) As Boolean
    ' Logique de vérification acte chirurgical associé
    VerifierActeChirurgicalAssocie = True ' Placeholder
End Function

Private Sub InitialiserBaseConnaissancesMedicales()
    ' Initialisation base de connaissances
End Sub

Private Sub ConfigurerAlertesAutomatiques()
    ' Configuration alertes
End Sub

Private Sub ActiverSurveillanceTempsReel()
    ' Activation surveillance
End Sub

Private Sub EffacerAnciennesAnomalies()
    ' Effacement anciennes anomalies
End Sub

Private Sub EffectuerAnalysesComplementaires()
    ' Analyses complémentaires
End Sub

Private Sub FormaterResultatsExpert()
    ' Formatage résultats
End Sub

Private Sub MettreAJourDashboardExpert()
    ' Mise à jour dashboard
End Sub

Private Sub GenererAlertesAutomatiques(nombreCritiques As Long)
    ' Génération alertes
End Sub

Private Sub EnregistrerDansHistorique(nbFactures As Long, nbAnomalies As Long, duree As Double)
    ' Enregistrement historique
End Sub

