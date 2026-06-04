Attribute VB_Name = "Module1"
Option Explicit

'================================================================
' 禁食・アレルギー検証システム v43
' Mac版Excel対応 / Dictionary非依存 / 高速化対応
'
' 【v43変更点】
'   - 検証結果の「コース」「番号」が、昼夕でコースが違う利用者で
'     取り違わる不具合を修正
'   - 原因: BuildNumberCourseDictionary が辞書キーを利用者名のみで作り、
'     先に出てきた行(通常は昼)だけを採用していたため、夕の行にも
'     昼のコースが入っていた(例: 伊東恵美子 昼k2/夕a → 夕がk2に化ける)
'   - 対策: 辞書キーを「利用者名 + 昼夕」に変更し、出力側も昼夕付きで引く
'
' 【v42変更点】
'   - _弁当別禁食マスタ シートを新設
'     弁当種類×メニュー名の組み合わせで NG材料を追加展開
'     全弁当共通の _メニュー材料補完 と分離し、弁当種類限定の補完を実現
'   - LoadBentoSpecificRules / ExpandFoodNameWithBentoRule 関数を追加
'
' 【v41変更点】
'   - 代替肉/豆製品の自動抑制(IsPlantBasedMeat/IsTofuProduct)を
'     _メニュー材料補完での明示登録時はオーバーライドするよう修正
'   - 例: 「茄子と大豆ミートのカレー風味」に鶏ひき肉が混入している場合、
'     _メニュー材料補完に「鶏肉|ひき肉」を登録すれば肉禁の利用者にヒット
'   - HasExplicitIngredient ヘルパー関数を追加
'   - v40 のデバッグ用 Debug.Print(橋章/れんこん/ExpandFoodName)を削除
'
' 【v40変更点】
'   - ExpandFoodName に完全一致最優先ロジックを追加
'     「メニュー名フル」が _メニュー材料補完 に登録されていれば、
'     部分一致は無視して完全一致のみ採用
'   - IsFalsePositiveByContext: メニュー名側の文脈ワードヒットも判定対象に
'     (foodNameForMatch を引数追加)
'
' 【v39変更点】
'   - ポケット限定指定を汎用化: 「魚(Aポケットのみ)」「卵(Bだけ)」等の
'     どのキーワードでも動作するように
'   - キーワード→ポケットのDictionaryで管理(ParsePocketRestrictions)
'
' 【v38変更点】
'   - 検証結果シートをA4横印刷向けに自動整形
'
' 【v37変更点】
'   - 全ボタンをSheet1(操作パネル)に集約
'   - 各ボタンの下に使い方の説明書きを配置
'
' 【v36変更点】
'   - 1b透析食の1日行数を9行→8行に修正
'   - 1c健康ボリューム食の1日行数を10行→12行に修正
'   - パターン2を全面刷新: G列「エネルギー」ヘッダー起点
'
' 【v34変更点】
'   - メニュー貼付シート列構造を9列に拡張(D~L列、A''/B'追加)
'   - 健康ボリューム食用パターン1c追加(A,A',A'',B,B',C,D,E)
'   - 透析食「0」値混入対策(数値0/"0"を空扱い)
'
' 【v33変更点】
'   - 「_スキップリスト」シートで名前+弁当種類による除外を可能に
'   - 検証結果に昼夕境目・日付境目の太罫線を追加(視認性向上)
'
' 【v32変更点】
'   - 全メニュー対応(普通食以外も取込・検証可能に)
'   - 「_設定」シートで「普通食のみ/全メニュー」モードを切替
'   - 「6.メニュー範囲切替」ボタンでトグル切替
'
' 【v31変更点】
'   - 全角カッコ()を半角に正規化してから処理(NormalizeParen追加)
'   - NormalizeBentoTypeに「野田市配食サービス(使捨)」全/半角、
'     「普通食(回収)」全/半角を追加
'   - 「魚ダメ(揚げ物はOK)」のOK判定をメニュー名から正しく除外
'     (ポケットレベルでOK判定するよう順序修正)
'   - 「肉(Aポケットのみ)」指定時、A以外のポケットを完全に除外
'     (ポケット限定でない肉同義語マッチもポケット判定を厳格化)
'
' 【v30変更点】
'   - v29: 「日付シール」検出機能(該当者を毎日出力、行を水色)
'   - v29: 「野田市配食サービス」セル(回収容器)をオレンジ色
'   - v30: 「前日履歴として保存」機能
'   - v30: 検証結果N列「前回」(前日夕or同日昼に名前があれば「有」)
'
' 【運用フロー】
'   1. 必要なメニュー貼付シートに貼り付け or「メニューフォルダから取込」
'   2. 「_調理指示表貼付」シートに調理指示表を貼り付け
'   3. ボタン: 1.フォルダ取込 → 3.調理指示表取込 → 4.検証実行
'   4. 翌日のために 5.前日履歴として保存 を押す
'================================================================

' ポケット列の定義(メニュー貼付シート)
' D=ポケットA, E=ポケットA', F=ポケットB, G=ポケットC,
' H=ポケットC', I=ポケットD, J=ポケットE


'================================================================
' メイン:検証実行
'================================================================
Public Sub 禁食アレルギー検証実行()

    Dim wsMenu As Worksheet
    Dim wsAllergy As Worksheet
    Dim wsResult As Worksheet
    
    Dim menuData As Variant
    Dim allergyData As Variant
    Dim resultArr() As Variant
    
    Dim menuRows As Long, allergyRows As Long
    Dim i As Long, j As Long, k As Long
    Dim resultCount As Long
    
    Dim pocketNames As Variant
    Dim pocketCols As Variant
    
    Dim allergyDate As Variant
    Dim allergyMealTime As String
    Dim allergyUser As String
    Dim allergyBento As String
    Dim allergyContent As String
    Dim allergyWords As Variant
    Dim w As Long
    Dim word As String
    
    Dim menuDate As Variant
    Dim menuMealTime As String
    Dim menuBento As String
    Dim foodName As String
    
    Dim dateMatch As Boolean
    
    Dim synonymData() As String
    Dim synonymCount As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call EnsureSheet("メニュー貼付")
    Call EnsureSheet("禁食アレルギー貼付")
    Call EnsureSheet("検証結果")
    Call EnsureSheet("_同義語マスタ")
    Call EnsureSheet("_メニュー材料補完")
    Call EnsureSheet("_前日履歴")
    Call EnsureSheet("_設定")
    Call EnsureSheet("_スキップリスト")
    Call EnsureSheet("_除外ペアマスタ")
    Call EnsureSheet("_弁当別禁食マスタ")
    
    Call LoadSynonyms(synonymData, synonymCount)
    
    ' 除外ペア読み込み(偽陽性防止用)
    Dim exclusionData() As String
    Dim exclusionCount As Long
    Call LoadExclusionPairs(exclusionData, exclusionCount)
    
    Dim ingredientDict As Object
    Set ingredientDict = CreateObject("Scripting.Dictionary")
    Call LoadIngredientSupplement(ingredientDict)
    
    ' v33: スキップリスト読み込み
    Dim skipList() As String
    Dim skipCount As Long
    Call LoadSkipList(skipList, skipCount)

    'v42: 弁当別禁食マスタ読み込み
    Dim bentoRules() As String
    Dim bentoRuleCount As Long
    Call LoadBentoSpecificRules(bentoRules, bentoRuleCount)
    
    Set wsMenu = ThisWorkbook.Worksheets("メニュー貼付")
    Set wsAllergy = ThisWorkbook.Worksheets("禁食アレルギー貼付")
    Set wsResult = ThisWorkbook.Worksheets("検証結果")
    
    pocketNames = Array("A", "A'", "A''", "B", "B'", "C", "C'", "D", "E")
    pocketCols = Array(4, 5, 6, 7, 8, 9, 10, 11, 12)
    
    menuRows = wsMenu.Cells(wsMenu.Rows.Count, "A").End(xlUp).Row
    If menuRows < 2 Then
        Call CleanupAndExit
        MsgBox "メニュー貼付シートにデータがありません。" & vbCrLf & _
               "先に「メニュー取込」を実行してください。", vbExclamation
        Exit Sub
    End If
    menuData = wsMenu.Range("A1:L" & menuRows).Value
    
    allergyRows = wsAllergy.Cells(wsAllergy.Rows.Count, "A").End(xlUp).Row
    If allergyRows < 2 Then
        Call CleanupAndExit
        MsgBox "禁食アレルギー貼付シートにデータがありません。" & vbCrLf & _
               "先に「調理指示表取込」を実行してください。", vbExclamation
        Exit Sub
    End If
    allergyData = wsAllergy.Range("A1:G" & allergyRows).Value
    
    ReDim resultArr(1 To allergyRows * 7, 1 To 11)
    resultCount = 0
    
    Dim allergySize As String
    Dim allergyQty As Variant
    
    For i = 2 To allergyRows
        
        If Trim(CStr(allergyData(i, 1) & "")) = "" Then GoTo NextAllergy
        If Trim(CStr(allergyData(i, 3) & "")) = "" Then GoTo NextAllergy
        
        allergyDate = allergyData(i, 1)
        allergyMealTime = Trim(CStr(allergyData(i, 2) & ""))
        allergyUser = Trim(CStr(allergyData(i, 3) & ""))
        allergyBento = Trim(CStr(allergyData(i, 4) & ""))
        allergyContent = Trim(CStr(allergyData(i, 5) & ""))
        allergySize = Trim(CStr(allergyData(i, 6) & ""))
        allergyQty = allergyData(i, 7)
        
        If allergyContent = "" Then GoTo NextAllergy
        
        ' v33: スキップリスト判定(名前+弁当種類で除外)
        If IsInSkipList(allergyUser, allergyBento, skipList, skipCount) Then
            GoTo NextAllergy
        End If
        
        Dim processedContent As String
        processedContent = FilterBentoConditional(allergyContent, allergyBento)
        
        ' v29: 「日付シール」フラグを抽出
        Dim hasDateSeal As Boolean
        hasDateSeal = (InStr(1, processedContent, "日付シール", vbBinaryCompare) > 0)
        
        allergyWords = SplitAllergyContent(processedContent)
        
        ' v29: ワード配列から「日付シール」を除去(メニュー判定対象から外す)
        If hasDateSeal Then
            Dim cleanWords() As String
            Dim cwn As Long, cwi As Long
            cwn = 0
            ReDim cleanWords(0 To UBound(allergyWords))
            For cwi = LBound(allergyWords) To UBound(allergyWords)
                Dim cww As String
                cww = Trim(CStr(allergyWords(cwi)))
                If cww <> "" And InStr(1, cww, "日付シール") = 0 Then
                    cleanWords(cwn) = cww
                    cwn = cwn + 1
                End If
            Next cwi
            If cwn = 0 Then
                allergyWords = Split("", ",")
            Else
                ReDim Preserve cleanWords(0 To cwn - 1)
                allergyWords = cleanWords
            End If
        End If
        
        Dim meatPocketRestrict As String
        meatPocketRestrict = ParseMeatPocketRestriction(processedContent)
        
        ' v39: 汎用ポケット限定指定をDictionaryで取得(キーワード→ポケットリスト)
        Dim pocketRestrictDict As Object
        Set pocketRestrictDict = CreateObject("Scripting.Dictionary")
        Call ParsePocketRestrictions(processedContent, pocketRestrictDict)
        
        Dim okKeywords As String
        okKeywords = ExtractOkKeywords(processedContent)
        
        For j = 2 To menuRows
            
            If Trim(CStr(menuData(j, 1) & "")) = "" Then GoTo NextMenu
            
            menuDate = menuData(j, 1)
            menuMealTime = Trim(CStr(menuData(j, 2) & ""))
            menuBento = Trim(CStr(menuData(j, 3) & ""))
            
            dateMatch = CompareDate(allergyDate, menuDate)
            
            If dateMatch And menuMealTime = allergyMealTime And _
               NormalizeBentoType(menuBento) = NormalizeBentoType(allergyBento) Then
                
                For k = 0 To UBound(pocketNames)
                    
                    foodName = Trim(CStr(menuData(j, pocketCols(k)) & ""))
                    
                    If foodName <> "" Then
                        
                        Dim foodNameForMatch As String
                        foodNameForMatch = ExpandFoodName(foodName, ingredientDict)
                        'v42: 弁当種類限定の追加展開
                        foodNameForMatch = ExpandFoodNameWithBentoRule( _
                            foodNameForMatch, allergyBento, bentoRules, bentoRuleCount)
                        
                        ' 橋 章さん専用：牛・豚は禁食ではなく刻み判断用。ひき肉・コロッケ等は抽出しない
                        If IsHashiAkiraNoKizamiNeed(allergyUser, allergyContent, foodNameForMatch) Then
                         GoTo NextPocket
                        End If
                        If okKeywords <> "" Then
                            If IsFoodMatchesOkKeywords(foodNameForMatch, okKeywords, synonymData, synonymCount) Then
                                GoTo NextPocket
                            End If
                        End If
                        
                        ' (1) 分割した各禁食ワードで直接マッチ判定
                        For w = LBound(allergyWords) To UBound(allergyWords)
                            
                            word = Trim(CStr(allergyWords(w)))
                            If word <> "" Then
                                Dim matched As Boolean
                                Dim matchedKeyword As String
                                matched = False
                                matchedKeyword = ""
                                
                               If InStr(1, NormalizeKana(foodNameForMatch), NormalizeKana(word), vbBinaryCompare) > 0 Then
    ' ★ 除外ペアマスタによる偽陽性チェック
    If IsFalsePositiveByContext(word, allergyContent, foodNameForMatch, exclusionData, exclusionCount) Then
        GoTo SkipDirectMatch
    End If
                                    
                                    If IsMeatKeyword(word) Then
                                        If meatPocketRestrict <> "" Then
                                            If InStr(1, meatPocketRestrict, pocketNames(k)) = 0 Then
                                                GoTo SkipDirectMatch
                                            End If
                                        End If
                                        If IsPlantBasedMeat(foodName) Then
                                            If InStr(1, allergyContent, "畑のお肉") = 0 Then
                                                ' v41: _メニュー材料補完で該当ワードが明示登録されていれば抑制しない
                                                If Not HasExplicitIngredient(foodName, word, ingredientDict) Then
                                                    GoTo SkipDirectMatch
                                                End If
                                            End If
                                        End If
                                        If HasMeatShapeRestriction(allergyContent) Then
                                            If IsProcessedOrGroundMeat(foodNameForMatch) Then
                                                GoTo SkipDirectMatch
                                            End If
                                        End If
                                    End If
                                    If word = "豆" Then
                                        If IsTofuProduct(foodName) Then
                                            ' v41: _メニュー材料補完で該当ワードが明示登録されていれば抑制しない
                                            If Not HasExplicitIngredient(foodName, word, ingredientDict) Then
                                                GoTo SkipDirectMatch
                                            End If
                                        End If
                                    End If
                                    matched = True
                                    matchedKeyword = word
                                End If
SkipDirectMatch:
                                
                                If Not matched Then
                                    Dim synonymList As String
                                    synonymList = GetSynonyms(word, synonymData, synonymCount)
                                    If synonymList <> "" Then
                                        Dim syns As Variant
                                        syns = Split(synonymList, "|")
                                        Dim si As Long
                                        For si = LBound(syns) To UBound(syns)
                                            Dim syn As String
                                            syn = Trim(CStr(syns(si)))
                                            If syn <> "" Then
                                                If InStr(1, NormalizeKana(foodNameForMatch), NormalizeKana(syn), vbBinaryCompare) > 0 Then
    ' ★ 除外ペアマスタによる偽陽性チェック
    If IsFalsePositiveByContext(word, allergyContent, foodNameForMatch, exclusionData, exclusionCount) Then
        GoTo SkipSynMatch1
    End If
                                                    If word = "肉" Then
                                                        If meatPocketRestrict <> "" Then
                                                            If InStr(1, meatPocketRestrict, pocketNames(k)) = 0 Then
                                                                GoTo SkipSynMatch1
                                                            End If
                                                        End If
                                                        If IsPlantBasedMeat(foodName) Then
                                                            If InStr(1, allergyContent, "畑のお肉") = 0 Then
                                                                ' v41: _メニュー材料補完で該当ワードが明示登録されていれば抑制しない
                                                                If Not HasExplicitIngredient(foodName, word, ingredientDict) Then
                                                                    GoTo SkipSynMatch1
                                                                End If
                                                            End If
                                                        End If
                                                        If HasMeatShapeRestriction(allergyContent) Then
                                                            If IsProcessedOrGroundMeat(foodNameForMatch) Then
                                                                GoTo SkipSynMatch1
                                                            End If
                                                        End If
                                                    End If
                                                    matched = True
                                                    matchedKeyword = word & "(→" & syn & ")"
                                                    Exit For
SkipSynMatch1:
                                                End If
                                            End If
                                        Next si
                                    End If
                                End If
                                
                                If matched Then
                                    ' v39: 汎用ポケット限定チェック
                                    ' wordに対応するポケット限定があれば、該当ポケット以外を除外
                                    Dim genericRestrict As String
                                    genericRestrict = GetPocketRestriction(word, "", pocketRestrictDict)
                                    If genericRestrict <> "" Then
                                        If Not IsPocketAllowedByRestriction(genericRestrict, CStr(pocketNames(k))) Then
                                            matched = False  ' 該当ポケットでないので不採用
                                            GoTo SkipMatchAdd
                                        End If
                                    End If
                                    
                                    Dim dupAdded As Boolean
                                    dupAdded = False
                                    Dim dupR As Long
                                    For dupR = 1 To resultCount
                                        If resultArr(dupR, 3) = allergyUser And _
                                           resultArr(dupR, 7) = pocketNames(k) And _
                                           resultArr(dupR, 8) = foodName And _
                                           CompareDate(resultArr(dupR, 1), allergyDate) And _
                                           resultArr(dupR, 2) = allergyMealTime Then
                                            dupAdded = True
                                            Exit For
                                        End If
                                    Next dupR
                                    
                                    If Not dupAdded Then
                                        resultCount = resultCount + 1
                                        resultArr(resultCount, 1) = allergyDate
                                        resultArr(resultCount, 2) = allergyMealTime
                                        resultArr(resultCount, 3) = allergyUser
                                        resultArr(resultCount, 4) = allergyBento
                                        resultArr(resultCount, 5) = allergyContent
                                        ' v29: 日付シールあれば付記
                                        If hasDateSeal Then
                                            resultArr(resultCount, 6) = matchedKeyword & " / 日付シール"
                                        Else
                                            resultArr(resultCount, 6) = matchedKeyword
                                        End If
                                        resultArr(resultCount, 7) = pocketNames(k)
                                        resultArr(resultCount, 8) = foodName
                                        resultArr(resultCount, 9) = "要対応"
                                        resultArr(resultCount, 10) = allergySize
                                        resultArr(resultCount, 11) = allergyQty
                                    End If
SkipMatchAdd:
                                End If
                            End If
                        Next w
                        
                        ' (2) 文中スキャン
                        Dim mi As Long
                        
                        Dim hasMeatPrefix As Boolean
                        Dim meatTargets As String
                        hasMeatPrefix = False
                        meatTargets = ""
                        Call AnalyzeMeatRestriction(allergyContent, hasMeatPrefix, meatTargets)
                        
                        For mi = 1 To synonymCount
                            Dim masterWord As String
                            Dim masterSyns As String
                            masterWord = synonymData(1, mi)
                            masterSyns = synonymData(2, mi)
                            
                            Dim effectiveSyns As String
                            effectiveSyns = masterSyns
                            If masterWord = "肉" Then
                                If hasMeatPrefix Then
                                    If meatTargets = "" Then
                                        GoTo NextSynonym
                                    End If
                                    effectiveSyns = meatTargets
                                End If
                            End If
                            
                            If InStr(1, NormalizeKana(allergyContent), NormalizeKana(masterWord), vbBinaryCompare) > 0 Then
                                
                                If IsFalsePositiveByContext(masterWord, allergyContent, foodNameForMatch, exclusionData, exclusionCount) Then
                                    GoTo NextSynonym
                                End If
                                
                                Dim mSyns As Variant
                                mSyns = Split(effectiveSyns, "|")
                                Dim mj As Long
                                For mj = LBound(mSyns) To UBound(mSyns)
                                    Dim mSyn As String
                                    mSyn = Trim(CStr(mSyns(mj)))
                                    If mSyn <> "" Then
                                        If InStr(1, NormalizeKana(foodNameForMatch), NormalizeKana(mSyn), vbBinaryCompare) > 0 Then
                                          
                                            If masterWord = "肉" Then
                                                If meatPocketRestrict <> "" Then
                                                    If InStr(1, meatPocketRestrict, pocketNames(k)) = 0 Then
                                                        GoTo SkipMeatMatch
                                                    End If
                                                End If
                                                If IsPlantBasedMeat(foodName) Then
                                                    If InStr(1, allergyContent, "畑のお肉") = 0 Then
                                                        ' v41: _メニュー材料補完で該当ワードが明示登録されていれば抑制しない
                                                        If Not HasExplicitIngredient(foodName, masterWord, ingredientDict) Then
                                                            GoTo SkipMeatMatch
                                                        End If
                                                    End If
                                                End If
                                                If HasMeatShapeRestriction(allergyContent) Then
                                                    If IsProcessedOrGroundMeat(foodNameForMatch) Then
                                                        GoTo SkipMeatMatch
                                                    End If
                                                End If
                                            End If
                                            If masterWord = "豆" Then
                                                If IsTofuProduct(foodName) Then
                                                    ' v41: _メニュー材料補完で該当ワードが明示登録されていれば抑制しない
                                                    If Not HasExplicitIngredient(foodName, masterWord, ingredientDict) Then
                                                        GoTo SkipMeatMatch
                                                    End If
                                                End If
                                            End If
                                            
                                            ' v39: 汎用ポケット限定チェック(文中スキャン版)
                                            ' masterWord(例「魚」)にポケット限定があれば、該当ポケット以外を除外
                                            Dim genericRestrict2 As String
                                            genericRestrict2 = GetPocketRestriction("", masterWord, pocketRestrictDict)
                                            If genericRestrict2 <> "" Then
                                                If Not IsPocketAllowedByRestriction(genericRestrict2, CStr(pocketNames(k))) Then
                                                    GoTo SkipMeatMatch
                                                End If
                                            End If
                                            
                                            Dim alreadyAdded As Boolean
                                            alreadyAdded = False
                                            Dim chkR As Long
                                            For chkR = 1 To resultCount
                                                If resultArr(chkR, 3) = allergyUser And _
                                                   resultArr(chkR, 7) = pocketNames(k) And _
                                                   resultArr(chkR, 8) = foodName And _
                                                   CompareDate(resultArr(chkR, 1), allergyDate) And _
                                                   resultArr(chkR, 2) = allergyMealTime Then
                                                    alreadyAdded = True
                                                    Exit For
                                                End If
                                            Next chkR
                                            
                                            If Not alreadyAdded Then
                                                resultCount = resultCount + 1
                                                resultArr(resultCount, 1) = allergyDate
                                                resultArr(resultCount, 2) = allergyMealTime
                                                resultArr(resultCount, 3) = allergyUser
                                                resultArr(resultCount, 4) = allergyBento
                                                resultArr(resultCount, 5) = allergyContent
                                                ' v29: 日付シールあれば付記
                                                If hasDateSeal Then
                                                    resultArr(resultCount, 6) = masterWord & "(→" & mSyn & ") / 日付シール"
                                                Else
                                                    resultArr(resultCount, 6) = masterWord & "(→" & mSyn & ")"
                                                End If
                                                resultArr(resultCount, 7) = pocketNames(k)
                                                resultArr(resultCount, 8) = foodName
                                                resultArr(resultCount, 9) = "要対応"
                                                resultArr(resultCount, 10) = allergySize
                                                resultArr(resultCount, 11) = allergyQty
                                            End If
                                            Exit For
                                        End If
SkipMeatMatch:
                                    End If
                                Next mj
                            End If
NextSynonym:
                        Next mi
                    End If
NextPocket:
                Next k
                
                ' v29: 日付シール対応 - 既存ヒットなし or 全てスキップでも、
                ' この日その昼夕に既にヒット行があるかを再確認。
                ' なければ「日付シール専用行」を1行追加する
                ' (= 同日中に他の禁食該当があれば、上で「/ 日付シール」付きで既に追加済み。
                '    禁食該当がなくても日付シール対象者なので、専用行を追加して結果に出す)
                If hasDateSeal Then
                    Dim dsAlreadyHit As Boolean
                    Dim dsR As Long
                    dsAlreadyHit = False
                    For dsR = 1 To resultCount
                        If resultArr(dsR, 3) = allergyUser And _
                           resultArr(dsR, 2) = allergyMealTime Then
                            If CompareDate(resultArr(dsR, 1), allergyDate) Then
                                dsAlreadyHit = True
                                Exit For
                            End If
                        End If
                    Next dsR
                    
                    If Not dsAlreadyHit Then
                        resultCount = resultCount + 1
                        resultArr(resultCount, 1) = allergyDate
                        resultArr(resultCount, 2) = allergyMealTime
                        resultArr(resultCount, 3) = allergyUser
                        resultArr(resultCount, 4) = allergyBento
                        resultArr(resultCount, 5) = allergyContent
                        resultArr(resultCount, 6) = "日付シール"
                        resultArr(resultCount, 7) = ""
                        resultArr(resultCount, 8) = ""
                        resultArr(resultCount, 9) = "日付シール"
                        resultArr(resultCount, 10) = allergySize
                        resultArr(resultCount, 11) = allergyQty
                    End If
                End If
            End If
NextMenu:
        Next j
NextAllergy:
    Next i
    
    ' 結果書き込み
    wsResult.Cells.Clear
    With wsResult
        .Range("A1").Value = "日付"
        .Range("B1").Value = "昼夕"
        .Range("C1").Value = "番号"
        .Range("D1").Value = "コース"
        .Range("E1").Value = "利用者名"
        .Range("F1").Value = "弁当種類"
        .Range("G1").Value = "禁食・アレルギー内容"
        .Range("H1").Value = "該当ワード"
        .Range("I1").Value = "該当ポケット"
        .Range("J1").Value = "該当食材"
        .Range("K1").Value = "サイズ"
        .Range("L1").Value = "個数"
        .Range("M1").Value = "前回"
        With .Range("A1:M1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 255)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    If resultCount > 0 Then
        Dim numCourseDict As Object
        Set numCourseDict = CreateObject("Scripting.Dictionary")
        Call BuildNumberCourseDictionary(numCourseDict)
        
        ' v30: 前回判定用 Dictionary 構築
        Dim prevMealDict As Object
        Set prevMealDict = CreateObject("Scripting.Dictionary")
        Call BuildPreviousMealDict(prevMealDict)
        
        Dim outArr() As Variant
        ReDim outArr(1 To resultCount, 1 To 13)
        For i = 1 To resultCount
            Dim userName As String
            userName = Trim(CStr(resultArr(i, 3) & ""))
            Dim numCourseStr As String
            numCourseStr = ""
            Dim mealKey As String
            mealKey = userName & "|" & Trim(CStr(resultArr(i, 2) & ""))  ' v43: 昼夕も含めて引く
            If numCourseDict.Exists(mealKey) Then
                numCourseStr = numCourseDict(mealKey)
            End If
            Dim numVal As String, courseVal As String
            numVal = ""
            courseVal = ""
            If numCourseStr <> "" Then
                Dim parts As Variant
                parts = Split(numCourseStr, "|")
                If UBound(parts) >= 0 Then numVal = CStr(parts(0))
                If UBound(parts) >= 1 Then courseVal = CStr(parts(1))
            End If
            
            outArr(i, 1) = resultArr(i, 1)
            outArr(i, 2) = resultArr(i, 2)
            outArr(i, 3) = numVal
            outArr(i, 4) = courseVal
            outArr(i, 5) = resultArr(i, 3)
            outArr(i, 6) = resultArr(i, 4)
            outArr(i, 7) = resultArr(i, 5)
            outArr(i, 8) = resultArr(i, 6)
            outArr(i, 9) = resultArr(i, 7)
            outArr(i, 10) = resultArr(i, 8)
            outArr(i, 11) = resultArr(i, 10)   ' サイズ
            outArr(i, 12) = resultArr(i, 11)   ' 個数
            outArr(i, 13) = CheckPreviousMeal(prevMealDict, userName, resultArr(i, 1), CStr(resultArr(i, 2) & ""))
        Next i
        wsResult.Range("A2").Resize(resultCount, 13).Value = outArr
        wsResult.Range("A2").Resize(resultCount, 13).Borders.LineStyle = xlContinuous
        
        ' M列「前回」を赤字・太字・中央寄せ
        With wsResult.Range("M2").Resize(resultCount, 1)
            .Font.Bold = True
            .Font.Color = RGB(192, 0, 0)
            .HorizontalAlignment = xlCenter
        End With
        wsResult.Range("A2").Resize(resultCount, 1).NumberFormat = "m/d"
        
        ' v29: 色付け処理
        Call ApplyResultCellColors(wsResult, resultCount)
        
        ' v33: 昼夕境目・日付境目の太罫線
        Call ApplyResultBoundaryBorders(wsResult, resultCount)
        
        ' v38: A4横印刷向けに整形
        Call FormatResultForPrint(wsResult, resultCount)
    End If
    
    wsResult.Columns("A:M").AutoFit
    wsResult.Activate
    wsResult.Range("A1").Select
    
    Call CleanupAndExit
    
    If resultCount = 0 Then
        MsgBox "該当する禁食・アレルギーはありませんでした。", vbInformation, "検証完了"
    Else
        MsgBox resultCount & " 件の要対応項目が見つかりました。", vbInformation, "検証完了"
    End If
    Exit Sub
    
ErrorHandler:
    Call CleanupAndExit
    MsgBox "検証実行でエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub


'================================================================
' 診断: メニューフォルダの状態をMsgBoxで表示
'================================================================
Public Sub メニューフォルダ診断()
    Dim folderPath As String
    Dim fileName As String
    Dim msg As String
    Dim cnt As Long
    
    folderPath = ThisWorkbook.Path
    If folderPath = "" Then
        MsgBox "マクロブックを保存してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    folderPath = folderPath & "メニュー" & Application.PathSeparator
    
    msg = "確認したパス:" & vbCrLf & folderPath & vbCrLf & vbCrLf
    
    If Dir(folderPath, vbDirectory) = "" Then
        msg = msg & "★フォルダが見つかりません★"
        MsgBox msg, vbExclamation, "メニューフォルダ診断"
        Exit Sub
    End If
    
    msg = msg & "【見つかったxlsxファイル】" & vbCrLf
    cnt = 0
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            cnt = cnt + 1
            Dim dst As String
            dst = DetermineMenuSheetName(fileName)
            If dst = "" Then dst = "(判定不能)"
            msg = msg & "  " & cnt & ". " & fileName & vbCrLf & _
                  "     → " & dst & vbCrLf
        End If
        fileName = Dir()
    Loop
    
    If cnt = 0 Then
        msg = msg & "(フォルダにxlsxファイルがありません)"
    End If
    
    MsgBox msg, vbInformation, "メニューフォルダ診断"
End Sub


'================================================================
' 取り込み[0]: フォルダ内のメニューファイルを一括自動取込
'================================================================
Public Sub メニューフォルダから取込()
    
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim destSheetName As String
    Dim bentoKind As String
    Dim totalCopied As Long
    Dim totalFiles As Long
    Dim copiedSummary As String
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    folderPath = ThisWorkbook.Path
    If folderPath = "" Then
        Call CleanupAndExit
        Application.DisplayAlerts = True
        MsgBox "マクロブックを保存してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    folderPath = folderPath & "メニュー" & Application.PathSeparator
    
    If Dir(folderPath, vbDirectory) = "" Then
        Call CleanupAndExit
        Application.DisplayAlerts = True
        MsgBox "「メニュー」サブフォルダが見つかりません。" & vbCrLf & vbCrLf & _
               "次のパスにフォルダを作成してください:" & vbCrLf & _
               folderPath & vbCrLf & vbCrLf & _
               "そこに各メニュー表のExcelファイルを入れてから再実行してください。", vbExclamation
        Exit Sub
    End If
    
    totalCopied = 0
    totalFiles = 0
    copiedSummary = ""
    
    fileName = Dir(folderPath & "*.xlsx")
    
    ' v32: 「_設定」シートのB1から「普通食のみ/全メニュー」を読み取る
    Dim enableOnlyFutsu As Boolean
    enableOnlyFutsu = IsOnlyFutsuMode()
    
    Application.StatusBar = "フォルダ取込開始..."
    
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            totalFiles = totalFiles + 1
            filePath = folderPath & fileName
            
            Application.StatusBar = "処理中: " & fileName
            DoEvents
            
            destSheetName = DetermineMenuSheetName(fileName)
            bentoKind = DetermineMenuBentoKind(fileName)
            
            If enableOnlyFutsu And destSheetName <> "_メニュー_普通" Then
                copiedSummary = copiedSummary & "  ・" & fileName & " → スキップ(普通食のみ取込中)" & vbCrLf
                GoTo NextFile
            End If
            
            If destSheetName <> "" Then
                Call EnsureSheet(destSheetName)
                Set wsDest = ThisWorkbook.Worksheets(destSheetName)
                wsDest.Cells.Clear
                
                DoEvents
                
                Set wbSrc = Nothing
                On Error Resume Next
                
                Dim pvw As ProtectedViewWindow
                Set pvw = Nothing
                Workbooks.Open fileName:=filePath, _
                    UpdateLinks:=0, _
                    ReadOnly:=True, _
                    IgnoreReadOnlyRecommended:=True, _
                    Notify:=False, _
                    AddToMru:=False
                
                Dim wb As Workbook
                For Each wb In Workbooks
                    If LCase(wb.FullName) = LCase(filePath) Then
                        Set wbSrc = wb
                        Exit For
                    End If
                Next wb
                
                If wbSrc Is Nothing Then
                    For Each pvw In Application.ProtectedViewWindows
                        If LCase(pvw.Workbook.FullName) = LCase(filePath) Then
                            pvw.Edit
                            Exit For
                        End If
                    Next pvw
                    
                    For Each wb In Workbooks
                        If LCase(wb.FullName) = LCase(filePath) Then
                            Set wbSrc = wb
                            Exit For
                        End If
                    Next wb
                End If
                
                Dim openErr As Long
                openErr = Err.Number
                On Error GoTo ErrHandler
                
                If wbSrc Is Nothing Then
                    copiedSummary = copiedSummary & "  ・" & fileName & " → ファイルが開けません(エラー" & openErr & ")" & vbCrLf
                    GoTo NextFile
                End If
                
                DoEvents
                
                If Not wbSrc Is Nothing Then
                    Dim destRow As Long
                    destRow = 1
                    Dim sheetsCopied As Long
                    sheetsCopied = 0
                    
                    For Each wsSrc In wbSrc.Worksheets
                        Dim srcLastRow As Long
                        Dim srcLastCol As Long
                        srcLastRow = 0
                        srcLastCol = 0
                        On Error Resume Next
                        Dim lastA As Long, lastB As Long, lastG As Long, lastJ As Long
                        lastA = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                        lastB = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row
                        lastG = wsSrc.Cells(wsSrc.Rows.Count, 7).End(xlUp).Row
                        lastJ = wsSrc.Cells(wsSrc.Rows.Count, 10).End(xlUp).Row
                        srcLastRow = lastA
                        If lastB > srcLastRow Then srcLastRow = lastB
                        If lastG > srcLastRow Then srcLastRow = lastG
                        If lastJ > srcLastRow Then srcLastRow = lastJ
                        srcLastCol = 14
                        On Error GoTo ErrHandler
                        
                        If srcLastRow >= 3 Then
                            Dim srcVals As Variant
                            srcVals = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(srcLastRow, srcLastCol)).Value
                            wsDest.Cells(destRow, 1).Resize(srcLastRow, srcLastCol).Value = srcVals
                            destRow = destRow + srcLastRow + 1
                            sheetsCopied = sheetsCopied + 1
                            DoEvents
                        End If
                    Next wsSrc
                    
                    wbSrc.Close SaveChanges:=False
                    
                    If sheetsCopied > 0 Then
                        totalCopied = totalCopied + 1
                        copiedSummary = copiedSummary & "  ・" & fileName & " → " & destSheetName & " (" & sheetsCopied & "シート)" & vbCrLf
                    Else
                        copiedSummary = copiedSummary & "  ・" & fileName & " → " & destSheetName & " (シートなし、スキップ)" & vbCrLf
                    End If
                End If
            Else
                copiedSummary = copiedSummary & "  ・" & fileName & " → 判定不能、スキップ" & vbCrLf
            End If
        End If
        
NextFile:
        fileName = Dir()
    Loop
    
    Application.DisplayAlerts = True
    Call CleanupAndExit
    
    If totalCopied = 0 Then
        MsgBox "「メニュー」フォルダにファイルが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "確認したパス:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
               "メニューExcelファイルを入れてから再実行してください。", vbExclamation
        Exit Sub
    End If
    
    MsgBox totalCopied & " / " & totalFiles & " 個のメニューファイルを貼付シートにコピーしました。" & vbCrLf & vbCrLf & _
           copiedSummary & vbCrLf & _
           "続けて「メニュー取込」を実行します。", vbInformation, "フォルダから取込完了"
    
    Call メニュー取込
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    Call CleanupAndExit
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "フォルダ取込でエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "処理中のファイル: " & fileName, vbCritical, "エラー"
End Sub


'================================================================
' 補助:ファイル名からメニュー種類を判定(取込先シート名)
'================================================================
Private Function DetermineMenuSheetName(ByVal fileName As String) As String
    Dim fn As String
    fn = fileName
    
    If InStr(fn, "塩分") > 0 Then
        DetermineMenuSheetName = "_メニュー_たんぱく塩分"
    ElseIf InStr(fn, "透析") > 0 Then
        DetermineMenuSheetName = "_メニュー_透析食"
    ElseIf InStr(fn, "幸") > 0 Then
        DetermineMenuSheetName = "_メニュー_幸たんぱく"
    ElseIf InStr(fn, "健康") > 0 Or InStr(fn, "ボリューム") > 0 Or InStr(fn, "ホ_リューム") > 0 Then
        DetermineMenuSheetName = "_メニュー_健康ボリューム"
    ElseIf InStr(fn, "やわらか") > 0 Then
        DetermineMenuSheetName = "_メニュー_やわらか"
    ElseIf InStr(fn, "ムース") > 0 Then
        DetermineMenuSheetName = "_メニュー_ムース"
    ElseIf InStr(fn, "カロリー") > 0 Then
        DetermineMenuSheetName = "_メニュー_カロリー"
    ElseIf InStr(fn, "消化") > 0 Then
        DetermineMenuSheetName = "_メニュー_消化"
    ElseIf InStr(fn, "普通") > 0 Then
        DetermineMenuSheetName = "_メニュー_普通"
    Else
        DetermineMenuSheetName = ""
    End If
End Function


Private Function DetermineMenuBentoKind(ByVal fileName As String) As String
    Dim fn As String
    fn = fileName
    
    If InStr(fn, "塩分") > 0 Then
        DetermineMenuBentoKind = "日替 たんぱく・塩分調整食"
    ElseIf InStr(fn, "透析") > 0 Then
        DetermineMenuBentoKind = "透析食"
    ElseIf InStr(fn, "幸") > 0 Then
        DetermineMenuBentoKind = "幸たんぱく食"
    ElseIf InStr(fn, "健康") > 0 Or InStr(fn, "ボリューム") > 0 Or InStr(fn, "ホ_リューム") > 0 Then
        DetermineMenuBentoKind = "健康ボリューム食"
    ElseIf InStr(fn, "やわらか") > 0 Then
        DetermineMenuBentoKind = "やわらか食"
    ElseIf InStr(fn, "ムース") > 0 Then
        DetermineMenuBentoKind = "ムースセット食"
    ElseIf InStr(fn, "カロリー") > 0 Then
        DetermineMenuBentoKind = "宅配クックのカロリー食"
    ElseIf InStr(fn, "消化") > 0 Then
        DetermineMenuBentoKind = "消化にやさしい食"
    ElseIf InStr(fn, "普通") > 0 Then
        DetermineMenuBentoKind = "普通食"
    Else
        DetermineMenuBentoKind = ""
    End If
End Function


'================================================================
' 取り込み : メニュー貼付シート(各種) → 「メニュー貼付」へ整形
'================================================================
Public Sub メニュー取込()

    Dim wsMenu As Worksheet
    Dim outArr() As Variant
    Dim outCount As Long
    Dim totalFound As Long
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call EnsureSheet("メニュー貼付")
    Set wsMenu = ThisWorkbook.Worksheets("メニュー貼付")
    
    ReDim outArr(1 To 20000, 1 To 12)
    outCount = 0
    totalFound = 0
    
    Dim curStep As String
    
    ' v32: 「_設定」シートのB1から「普通食のみ/全メニュー」を読み取る
    Dim enableOnlyFutsu As Boolean
    enableOnlyFutsu = IsOnlyFutsuMode()
    
    curStep = "_メニュー_普通"
    Call ProcessMenuSheetPattern1("_メニュー_普通", "普通食", 2, 9, "1a", outArr, outCount, totalFound)
    
    If Not enableOnlyFutsu Then
        curStep = "_メニュー_幸たんぱく"
        Call ProcessMenuSheetPattern1("_メニュー_幸たんぱく", "幸たんぱく食", 2, 9, "1a", outArr, outCount, totalFound)
        curStep = "_メニュー_健康ボリューム"
        Call ProcessMenuSheetPattern1("_メニュー_健康ボリューム", "健康ボリューム食", 2, 9, "1c", outArr, outCount, totalFound)
        curStep = "_メニュー_たんぱく塩分"
        Call ProcessMenuSheetPattern1("_メニュー_たんぱく塩分", "日替 たんぱく・塩分調整食", 2, 10, "1b", outArr, outCount, totalFound)
        curStep = "_メニュー_透析食"
        Call ProcessMenuSheetPattern1("_メニュー_透析食", "透析食", 2, 9, "1b", outArr, outCount, totalFound)
        curStep = "_メニュー_やわらか"
        Call ProcessMenuSheetPattern2("_メニュー_やわらか", "やわらか食", outArr, outCount, totalFound)
        curStep = "_メニュー_ムース"
        Call ProcessMenuSheetPattern2("_メニュー_ムース", "ムースセット食", outArr, outCount, totalFound)
        curStep = "_メニュー_カロリー"
        Call ProcessMenuSheetPattern2("_メニュー_カロリー", "宅配クックのカロリー食", outArr, outCount, totalFound)
        curStep = "_メニュー_消化"
        Call ProcessMenuSheetPattern2("_メニュー_消化", "消化にやさしい食", outArr, outCount, totalFound)
    End If
    
    curStep = "結果書き込み"
    
    If totalFound = 0 Then
        Call CleanupAndExit
        MsgBox "どのメニューシートにもデータが見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    
    curStep = "結果書き込み: シートクリア"
    wsMenu.Cells.Clear
    
    curStep = "結果書き込み: ヘッダー"
    With wsMenu
        .Range("A1:L1").Value = Array("日付", "昼夕", "弁当種類", _
            "ポケットA", "ポケットA'", "ポケットA''", "ポケットB", "ポケットB'", _
            "ポケットC", "ポケットC'", "ポケットD", "ポケットE")
        With .Range("A1:L1")
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
        End With
    End With
    
    If outCount > 0 Then
        curStep = "結果書き込み: writeArr確保 outCount=" & outCount
        Dim writeArr() As Variant
        ReDim writeArr(1 To outCount, 1 To 12)
        Dim a As Long, b As Long
        
        curStep = "結果書き込み: writeArrコピー"
        For a = 1 To outCount
            For b = 1 To 12
                Dim srcVal As Variant
                srcVal = outArr(a, b)
                If IsError(srcVal) Then
                    writeArr(a, b) = "#ERR"
                ElseIf IsObject(srcVal) Then
                    writeArr(a, b) = ""
                ElseIf IsNull(srcVal) Or IsEmpty(srcVal) Then
                    writeArr(a, b) = ""
                Else
                    writeArr(a, b) = srcVal
                End If
            Next b
        Next a
        
        curStep = "結果書き込み: シートへ書込(一括)"
        On Error Resume Next
        wsMenu.Range("A2").Resize(outCount, 12).Value = writeArr
        Dim writeErr As Long
        writeErr = Err.Number
        On Error GoTo ErrHandler
        
        If writeErr <> 0 Then
            curStep = "結果書き込み: 1行ずつ書込にフォールバック (一括失敗 Err=" & writeErr & ")"
            Dim wr As Long, wc As Long
            For wr = 1 To outCount
                For wc = 1 To 12
                    On Error Resume Next
                    wsMenu.Cells(wr + 1, wc).Value = writeArr(wr, wc)
                    On Error GoTo ErrHandler
                Next wc
            Next wr
        End If
        
        curStep = "結果書き込み: 日付書式"
        On Error Resume Next
        wsMenu.Range("A2").Resize(outCount, 1).NumberFormat = "m/d"
        On Error GoTo ErrHandler
    End If
    
    curStep = "結果書き込み: AutoFit"
    On Error Resume Next
    wsMenu.Columns("A:L").AutoFit
    On Error GoTo ErrHandler
    
    Call CleanupAndExit
    MsgBox outCount & " 件のメニューを取り込みました。" & vbCrLf & _
           "(" & totalFound & " シートから取り込み)", vbInformation, "メニュー取込完了"
    Exit Sub
    
ErrHandler:
    Call CleanupAndExit
    MsgBox "メニュー取込でエラーが発生しました。" & vbCrLf & _
           "処理中の箇所: " & curStep & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "outCount: " & outCount, vbCritical, "エラー"
End Sub


'================================================================
' パターン1(普通食タイプ)
'================================================================
Private Sub ProcessMenuSheetPattern1(ByVal sheetName As String, ByVal bentoType As String, _
                                      ByVal lunchCol As Long, ByVal dinnerCol As Long, _
                                      ByVal subPattern As String, _
                                      ByRef outArr() As Variant, ByRef outCount As Long, _
                                      ByRef totalFound As Long)
    
    Dim wsSrc As Worksheet
    Dim srcLastRow As Long
    Dim r As Long, rr As Long, i As Long
    Dim weekStart As Date, weekEnd As Date
    Dim dayCount As Long
    Dim a1Val As Variant
    Dim weekHeaderRow As Long, weekDataStartRow As Long
    Dim blockRow As Long, weekDayIdx As Long
    Dim curDate As Date
    Dim dishLunch(1 To 8) As String     ' v34: 8品まで対応(1c用)
    Dim dishDinner(1 To 8) As String
    Dim hasLunch As Boolean, hasDinner As Boolean
    
    On Error Resume Next
    Set wsSrc = Nothing
    Set wsSrc = ThisWorkbook.Worksheets(sheetName)
    If wsSrc Is Nothing Then Exit Sub
    On Error GoTo 0
    
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    Dim srcLastRowB As Long
    srcLastRowB = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If srcLastRowB > srcLastRow Then srcLastRow = srcLastRowB
    If srcLastRow < 6 Then Exit Sub
    
    r = 1
    Do While r <= srcLastRow
        
        a1Val = wsSrc.Cells(r, 1).Value
        Call ParseWeekRange(a1Val, weekStart, weekEnd, dayCount)
        
        If weekStart = 0 Then
            r = r + 1
        Else
            weekHeaderRow = r
            totalFound = totalFound + 1
            
            weekDataStartRow = 0
            For rr = weekHeaderRow + 1 To weekHeaderRow + 60
                If rr > srcLastRow Then Exit For
                Dim aVal2 As String, bVal2 As String, dVal2 As String
                aVal2 = Trim(CStr(wsSrc.Cells(rr, 1).Value & ""))
                bVal2 = Trim(CStr(wsSrc.Cells(rr, 2).Value & ""))
                dVal2 = Trim(CStr(wsSrc.Cells(rr, dinnerCol).Value & ""))
                
                If aVal2 = "メニュー日" Or aVal2 = "商品名" Or aVal2 = "栄養価" Then
                    GoTo NextScanRow
                End If
                If bVal2 = "商品名" Or bVal2 = "≪昼≫" Or bVal2 = "内容量(g)" Then
                    GoTo NextScanRow
                End If
                If dVal2 = "商品名" Or dVal2 = "≪夕≫" Or dVal2 = "内容量(g)" Then
                    GoTo NextScanRow
                End If
                
                If (bVal2 <> "" And bVal2 <> "ご飯" And InStr(bVal2, "メニューは食材") = 0) Or _
                   (dVal2 <> "" And dVal2 <> "ご飯" And dVal2 <> "夕" And InStr(dVal2, "メニューは食材") = 0) Then
                    Dim aRaw As Variant
                    aRaw = wsSrc.Cells(rr, 1).Value
                    Dim aIsBlockTop As Boolean
                    aIsBlockTop = False
                    If IsEmpty(aRaw) Or aRaw = "" Then aIsBlockTop = True
                    If IsNumeric(aRaw) Then aIsBlockTop = True
                    If VarType(aRaw) = vbDate Then aIsBlockTop = True
                    
                    If aIsBlockTop Then
                        weekDataStartRow = rr
                        Exit For
                    End If
                End If
NextScanRow:
            Next rr
            
            If weekDataStartRow = 0 Then
                r = weekHeaderRow + 1
                GoTo NextWeekLoopP1
            End If
            
            blockRow = weekDataStartRow
            weekDayIdx = 0
            
            Do While blockRow <= srcLastRow And weekDayIdx < dayCount
                
                Dim ws2 As Date, we2 As Date, dc2 As Long
                Call ParseWeekRange(wsSrc.Cells(blockRow, 1).Value, ws2, we2, dc2)
                If ws2 > 0 And blockRow > weekHeaderRow Then Exit Do
                
                curDate = weekStart + weekDayIdx
                
                hasLunch = False
                hasDinner = False
                ' v34: パターンごとに読み取り行数を変える
                '   1c (健康ボリューム食) = 8品
                '   1a/1b = 6品
                Dim dishCount As Long
                If subPattern = "1c" Then
                    dishCount = 8
                Else
                    dishCount = 6
                End If
                For i = 1 To 8
                    dishLunch(i) = ""
                    dishDinner(i) = ""
                Next i
                
                ' v35: 1c(健康ボリューム食)はブロック先頭1行目がタイトル行なのでスキップ
                '   読み取り開始位置を blockRow+1 にする(1a/1bは blockRow のまま)
                Dim readStartRow As Long
                If subPattern = "1c" Then
                    readStartRow = blockRow + 1
                Else
                    readStartRow = blockRow
                End If
                
                For i = 1 To dishCount
                    Dim rawL As Variant, rawD As Variant
                    rawL = wsSrc.Cells(readStartRow + i - 1, lunchCol).Value
                    rawD = wsSrc.Cells(readStartRow + i - 1, dinnerCol).Value
                    ' v34: 数値0や"0"を空扱い(透析食の「0」混入対策)
                    If IsNumeric(rawL) Then
                        If CDbl(rawL) = 0 Then rawL = ""
                    End If
                    If IsNumeric(rawD) Then
                        If CDbl(rawD) = 0 Then rawD = ""
                    End If
                    dishLunch(i) = Trim(CStr(rawL & ""))
                    dishDinner(i) = Trim(CStr(rawD & ""))
                    If dishLunch(i) <> "" Then hasLunch = True
                    If dishDinner(i) <> "" Then hasDinner = True
                Next i
                
                If hasLunch Then
                    outCount = outCount + 1
                    outArr(outCount, 1) = curDate
                    outArr(outCount, 2) = "昼"
                    outArr(outCount, 3) = bentoType
                    Call AssignPocketsToOutArr(outArr, outCount, dishLunch, subPattern)
                End If
                
                If hasDinner Then
                    outCount = outCount + 1
                    outArr(outCount, 1) = curDate
                    outArr(outCount, 2) = "夕"
                    outArr(outCount, 3) = bentoType
                    Call AssignPocketsToOutArr(outArr, outCount, dishDinner, subPattern)
                End If
                
                ' v36: パターンごとの1日行数
                '   1c (健康ボリューム食) = 12行/日(タイトル1+おかず8+ご飯1+空2)
                '   1b (透析食) = 8行/日 ※注: 1bでも たんぱく塩分は9行/日のため、
                '                         シート名で個別判定
                '   1a (普通食、幸たんぱく) = 9行/日
                If subPattern = "1c" Then
                    blockRow = blockRow + 12
                ElseIf subPattern = "1b" And InStr(sheetName, "透析") > 0 Then
                    blockRow = blockRow + 8
                Else
                    blockRow = blockRow + 9
                End If
                weekDayIdx = weekDayIdx + 1
            Loop
            
            r = blockRow
        End If
        
NextWeekLoopP1:
    Loop
End Sub


' v34: ポケット振り分け (9ポケット対応)
' outArr の列構成:
'   1=日付, 2=昼夕, 3=弁当種類,
'   4=A, 5=A', 6=A'', 7=B, 8=B', 9=C, 10=C', 11=D, 12=E
'
'   1a (普通食、幸たんぱく): dishes(1..6) → A, A', B, C, D, E
'   1b (透析食、たんぱく塩分): dishes(1..6) → A, A', B, C, C', D
'   1c (健康ボリューム食): dishes(1..8) → A, A', A'', B, B', C, D, E
Private Sub AssignPocketsToOutArr(ByRef outArr() As Variant, ByVal idx As Long, _
                                   ByRef dishes() As String, ByVal subPattern As String)
    ' まず全ポケット列を空欄初期化
    Dim c As Long
    For c = 4 To 12
        outArr(idx, c) = ""
    Next c
    
    Select Case subPattern
        Case "1b"
            ' A, A', B, C, C', D
            outArr(idx, 4) = dishes(1)   ' A
            outArr(idx, 5) = dishes(2)   ' A'
            outArr(idx, 7) = dishes(3)   ' B
            outArr(idx, 9) = dishes(4)   ' C
            outArr(idx, 10) = dishes(5)  ' C'
            outArr(idx, 11) = dishes(6)  ' D
            
        Case "1c"
            ' A, A', A'', B, B', C, D, E (健康ボリューム食)
            outArr(idx, 4) = dishes(1)   ' A
            outArr(idx, 5) = dishes(2)   ' A'
            outArr(idx, 6) = dishes(3)   ' A''
            outArr(idx, 7) = dishes(4)   ' B
            outArr(idx, 8) = dishes(5)   ' B'
            outArr(idx, 9) = dishes(6)   ' C
            outArr(idx, 11) = dishes(7)  ' D
            outArr(idx, 12) = dishes(8)  ' E
            
        Case Else  ' "1a"
            ' A, A', B, C, D, E (普通食)
            outArr(idx, 4) = dishes(1)   ' A
            outArr(idx, 5) = dishes(2)   ' A'
            outArr(idx, 7) = dishes(3)   ' B
            outArr(idx, 9) = dishes(4)   ' C
            outArr(idx, 11) = dishes(5)  ' D
            outArr(idx, 12) = dishes(6)  ' E
    End Select
End Sub

'================================================================
' パターン2(やわらか・ムース・カロリー・消化食タイプ)
' v40: A列の日番号を独立に追跡、A1から月番号抽出、
'      G列ヘッダーは「エネルギー*」で前方一致
'================================================================
Private Sub ProcessMenuSheetPattern2(ByVal sheetName As String, ByVal bentoType As String, _
                                      ByRef outArr() As Variant, ByRef outCount As Long, _
                                      ByRef totalFound As Long)
    
    Dim wsSrc As Worksheet
    Dim srcLastRow As Long
    Dim r As Long, i As Long
    Dim curDate As Date
    Dim defaultMonth As Long, defaultYear As Long
    Dim lastDayNum As Long
    Dim curMeal As String
    Dim dishCount As Long
    Dim dishes(1 To 8) As String
    Dim foundData As Boolean
    
    On Error Resume Next
    Set wsSrc = Nothing
    Set wsSrc = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSrc Is Nothing Then Exit Sub
    
    ' 範囲を広めに取得
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, "G").End(xlUp).Row
    Dim srcLastRowA As Long
    srcLastRowA = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If srcLastRowA > srcLastRow Then srcLastRow = srcLastRowA
    Dim srcLastRowB As Long
    srcLastRowB = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If srcLastRowB > srcLastRow Then srcLastRow = srcLastRowB
    If srcLastRow < 5 Then Exit Sub
    
    ' 弁当種類別の品数決定
    Select Case True
        Case InStr(sheetName, "やわらか") > 0
            dishCount = 5
        Case InStr(sheetName, "ムース") > 0
            dishCount = 3
        Case InStr(sheetName, "カロリー") > 0
            dishCount = 4
        Case InStr(sheetName, "消化") > 0
            dishCount = 4
        Case Else
            dishCount = 4
    End Select
    
    ' === A1から月番号を抽出 ===
    defaultMonth = 0
    Dim a1V As Variant
    a1V = wsSrc.Cells(1, 1).Value
    If IsNumeric(a1V) Then
        If CLng(a1V) >= 1 And CLng(a1V) <= 12 Then
            defaultMonth = CLng(a1V)
        End If
    Else
        Dim a1Str As String
        a1Str = Trim(CStr(a1V & ""))
        Dim posTsuki As Long
        posTsuki = InStr(a1Str, "月")
        If posTsuki > 1 Then
            Dim numStr As String
            numStr = Trim(Left(a1Str, posTsuki - 1))
            If IsNumeric(numStr) Then
                Dim n As Long
                n = CLng(numStr)
                If n >= 1 And n <= 12 Then defaultMonth = n
            End If
        End If
    End If
    If defaultMonth = 0 Then defaultMonth = Month(Date)
    
    defaultYear = Year(Date)
    If defaultMonth < Month(Date) - 6 Then defaultYear = defaultYear + 1
    
    lastDayNum = 0
    foundData = False
    
    ' === メインループ: A列とG列を上から順にスキャン ===
    For r = 1 To srcLastRow
        
        ' (1) A列に1～31の整数があれば日番号を更新
        Dim aRaw As Variant
        aRaw = wsSrc.Cells(r, 1).Value
        If IsValidDayNumberLoose(aRaw) Then
            lastDayNum = CLng(aRaw)
        End If
        
        ' (2) G列に「エネルギー」で始まるヘッダーがあれば食事ブロックを処理
        Dim gVal As String
        gVal = Trim(CStr(wsSrc.Cells(r, 7).Value & ""))
        
        If Left(gVal, 5) = "エネルギー" Then
            
            ' 食区分(昼/夕)はヘッダー+1～+4行のB列から探す
            curMeal = ""
            Dim k As Long
            For k = 1 To 4
                If r + k > srcLastRow Then Exit For
                Dim bVal As String
                bVal = Trim(CStr(wsSrc.Cells(r + k, 2).Value & ""))
                If bVal = "昼" Or bVal = "夕" Then
                    curMeal = bVal
                    Exit For
                End If
            Next k
            
            ' ヘッダー+1～+4行のA列に日番号があれば、優先的に採用
            For k = 1 To 4
                If r + k > srcLastRow Then Exit For
                Dim aRaw2 As Variant
                aRaw2 = wsSrc.Cells(r + k, 1).Value
                If IsValidDayNumberLoose(aRaw2) Then
                    lastDayNum = CLng(aRaw2)
                    Exit For
                End If
            Next k
            
            If lastDayNum > 0 And curMeal <> "" Then
                ' 日付計算
                curDate = 0
                On Error Resume Next
                curDate = DateSerial(defaultYear, defaultMonth, lastDayNum)
                On Error GoTo 0
                
                If curDate > 0 Then
                    For i = 1 To 8
                        dishes(i) = ""
                    Next i
                    
                    Dim collectedCount As Long
                    collectedCount = 0
                    Dim scanRow As Long
                    For scanRow = r + 2 To r + 2 + dishCount + 5
                        If scanRow > srcLastRow Then Exit For
                        Dim itemVal As String
                        itemVal = Trim(CStr(wsSrc.Cells(scanRow, 7).Value & ""))
                        
                        ' 次の食事ヘッダーに到達したら抜ける
                        If Left(itemVal, 5) = "エネルギー" Then Exit For
                        ' 数値だけのセル(栄養価)はスキップ
                        If IsNumeric(itemVal) Then GoTo NextScan
                        ' 栄養価のラベル文字列もスキップ
                        If itemVal = "" Or itemVal = "0" Or _
                           InStr(itemVal, "蛋白質") > 0 Or _
                           InStr(itemVal, "脂質") > 0 Or _
                           InStr(itemVal, "炭水化物") > 0 Or _
                           InStr(itemVal, "ナトリウム") > 0 Or _
                           InStr(itemVal, "カリウム") > 0 Or _
                           InStr(itemVal, "食塩相当量") > 0 Or _
                           InStr(itemVal, "リン") > 0 Then
                            GoTo NextScan
                        End If
                        
                        collectedCount = collectedCount + 1
                        If collectedCount <= dishCount Then
                            dishes(collectedCount) = itemVal
                        Else
                            Exit For
                        End If
NextScan:
                    Next scanRow
                    
                    If collectedCount > 0 Then
                        outCount = outCount + 1
                        outArr(outCount, 1) = curDate
                        outArr(outCount, 2) = curMeal
                        outArr(outCount, 3) = bentoType
                        Dim pc As Long
                        For pc = 4 To 12
                            outArr(outCount, pc) = ""
                        Next pc
                        ' 品数別の振り分け
                        '  3品 (ムース)         → A, B, C
                        '  4品 (カロリー/消化)  → A, B, C, D
                        '  5品 (やわらか)       → A, B, C, D, E
                        If dishCount >= 1 Then outArr(outCount, 4) = dishes(1)
                        If dishCount >= 2 Then outArr(outCount, 7) = dishes(2)
                        If dishCount >= 3 Then outArr(outCount, 9) = dishes(3)
                        If dishCount >= 4 Then outArr(outCount, 11) = dishes(4)
                        If dishCount >= 5 Then outArr(outCount, 12) = dishes(5)
                        foundData = True
                    End If
                End If
            End If
        End If
    Next r
    
    If foundData Then totalFound = totalFound + 1
End Sub


'================================================================
' 取り込み : 調理指示表 → 禁食アレルギー貼付
'================================================================
Public Sub 調理指示表取込()

    Dim wsSrc As Worksheet
    Dim wsAllergy As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim outArr() As Variant
    Dim outCount As Long
    
    Dim curDate As Date
    Dim curMeal As String
    Dim curBlock As String
    Dim inDataRows As Boolean
    
    Dim cellA As String
    Dim cellE As String
    Dim cellF As String
    Dim cellH As String
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call EnsureSheet("_調理指示表貼付")
    Call EnsureSheet("禁食アレルギー貼付")
    
    Set wsSrc = ThisWorkbook.Worksheets("_調理指示表貼付")
    Set wsAllergy = ThisWorkbook.Worksheets("禁食アレルギー貼付")
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        Call CleanupAndExit
        MsgBox "「_調理指示表貼付」シートに調理指示表を貼り付けてから実行してください。", vbExclamation
        Exit Sub
    End If
    
    ReDim outArr(1 To lastRow * 5, 1 To 7)
    outCount = 0
    
    inDataRows = False
    curDate = 0
    curMeal = ""
    curBlock = ""
    
    Dim sizeLabels(1 To 50) As String
    Dim sizeLabelsLoaded As Boolean
    sizeLabelsLoaded = False
    
    Dim wsDiag As Worksheet
    Call EnsureSheet("_診断")
    Set wsDiag = ThisWorkbook.Worksheets("_診断")
    wsDiag.Cells.Clear
    wsDiag.Range("A1").Value = "行"
    wsDiag.Range("B1").Value = "A列"
    wsDiag.Range("C1").Value = "C列"
    wsDiag.Range("D1").Value = "E列"
    wsDiag.Range("E1").Value = "F列"
    wsDiag.Range("F1").Value = "H列"
    wsDiag.Range("G1").Value = "判定"
    wsDiag.Range("H1").Value = "状態"
    wsDiag.Range("A1:H1").Font.Bold = True
    
    Dim diagRow As Long
    diagRow = 2
    
    For r = 1 To lastRow
        
        cellA = Trim(CStr(wsSrc.Cells(r, 1).Value & ""))
        cellE = Trim(CStr(wsSrc.Cells(r, 5).Value & ""))
        cellF = Trim(CStr(wsSrc.Cells(r, 6).Value & ""))
        cellH = Trim(CStr(wsSrc.Cells(r, 8).Value & ""))
        Dim cellC As String
        cellC = Trim(CStr(wsSrc.Cells(r, 3).Value & ""))
        
        Dim judge As String
        Dim status As String
        judge = ""
        status = ""
        
        If InStr(cellH, "嫌いなもの一覧表") > 0 Then
            curDate = ParseDateFromHeader(cellA)
            curMeal = NormalizeMealTime(cellE)
            curBlock = "嫌いなもの"
            inDataRows = False
            sizeLabelsLoaded = False
            Dim sli As Long
            For sli = 1 To 50
                sizeLabels(sli) = ""
            Next sli
            judge = "嫌いなものブロック開始"
            status = "日付=" & curDate & " 昼夕=" & curMeal
            GoTo WriteDiag
        ElseIf InStr(cellH, "禁食一覧表") > 0 Then
            curDate = ParseDateFromHeader(cellA)
            curMeal = NormalizeMealTime(cellE)
            curBlock = "禁食"
            inDataRows = False
            sizeLabelsLoaded = False
            For sli = 1 To 50
                sizeLabels(sli) = ""
            Next sli
            judge = "禁食ブロック開始"
            status = "日付=" & curDate & " 昼夕=" & curMeal
            GoTo WriteDiag
        ElseIf InStr(cellH, "アレルギー一覧表") > 0 Then
            curDate = ParseDateFromHeader(cellA)
            curMeal = NormalizeMealTime(cellE)
            curBlock = "アレルギー"
            inDataRows = False
            sizeLabelsLoaded = False
            For sli = 1 To 50
                sizeLabels(sli) = ""
            Next sli
            judge = "アレルギーブロック開始"
            status = "日付=" & curDate & " 昼夕=" & curMeal
            GoTo WriteDiag
        End If
        
        If cellA = "行" Then
            inDataRows = True
            Dim slCol As Long
            Dim slLabelCount As Long
            slLabelCount = 0
            If r + 1 <= lastRow Then
                For slCol = 7 To 50
                    Dim labelVal As String
                    labelVal = Trim(CStr(wsSrc.Cells(r + 1, slCol).Value & ""))
                    If labelVal <> "" And Not IsNumeric(labelVal) Then
                        sizeLabels(slCol) = labelVal
                        slLabelCount = slLabelCount + 1
                    End If
                Next slCol
            End If
            sizeLabelsLoaded = (slLabelCount > 0)
            judge = "ヘッダー行"
            status = "inDataRows=True サイズラベル数=" & slLabelCount
            GoTo WriteDiag
        End If
        
        If cellF = "合計数" Then
            judge = "合計数スキップ"
            GoTo WriteDiag
        End If
        
        If inDataRows And cellA <> "" And IsNumeric(cellA) Then
            
            Dim userName As String
            Dim bentoKind As String
            Dim allergyContent As String
            
            userName = Trim(CStr(wsSrc.Cells(r, 3).Value & ""))
            bentoKind = Trim(CStr(wsSrc.Cells(r, 5).Value & ""))
            allergyContent = Trim(CStr(wsSrc.Cells(r, 6).Value & ""))
            
            userName = CleanUserName(userName)
            
            Dim sizeOutList(1 To 20) As String
            Dim sizeQtyList(1 To 20) As Long
            Dim sizeOutCount As Long
            sizeOutCount = 0
            
            If sizeLabelsLoaded Then
                Dim szCol As Long
                For szCol = 7 To 50
                    If sizeLabels(szCol) <> "" Then
                        Dim szVal As Variant
                        szVal = wsSrc.Cells(r, szCol).Value
                        If Not IsEmpty(szVal) And Not IsNull(szVal) Then
                            Dim szStr As String
                            szStr = Trim(CStr(szVal & ""))
                            If szStr <> "" And IsNumeric(szStr) Then
                                Dim szQty As Double
                                szQty = CDbl(szStr)
                                If szQty >= 1 Then
                                    sizeOutCount = sizeOutCount + 1
                                    sizeOutList(sizeOutCount) = sizeLabels(szCol)
                                    sizeQtyList(sizeOutCount) = CLng(szQty)
                                End If
                            End If
                        End If
                    End If
                Next szCol
            End If
            
            If sizeOutCount = 0 Then
                sizeOutCount = 1
                sizeOutList(1) = ""
                sizeQtyList(1) = 1
            End If
            
            judge = "データ候補"
            status = "名前=[" & userName & "] 禁食=[" & allergyContent & "] サイズ数=" & sizeOutCount & " curDate=" & curDate
            
            If userName <> "" And allergyContent <> "" And curDate > 0 Then
                Dim soi As Long
                For soi = 1 To sizeOutCount
                    outCount = outCount + 1
                    outArr(outCount, 1) = curDate
                    outArr(outCount, 2) = curMeal
                    outArr(outCount, 3) = userName
                    outArr(outCount, 4) = bentoKind
                    outArr(outCount, 5) = allergyContent
                    outArr(outCount, 6) = sizeOutList(soi)
                    outArr(outCount, 7) = sizeQtyList(soi)
                Next soi
                judge = "採用"
                status = "outCount=" & outCount & " 名前=[" & userName & "] サイズ数=" & sizeOutCount
            Else
                judge = "不採用"
                If userName = "" Then status = "名前空"
                If allergyContent = "" Then status = status & " 禁食空"
                If curDate = 0 Then status = status & " 日付未確定"
            End If
        Else
            If cellA <> "" Or cellC <> "" Or cellE <> "" Or cellF <> "" Or cellH <> "" Then
                judge = "対象外"
                status = "inDataRows=" & inDataRows & " IsNum=" & IsNumeric(cellA)
            End If
        End If
        
WriteDiag:
        If cellA <> "" Or cellC <> "" Or cellE <> "" Or cellF <> "" Or cellH <> "" Or judge <> "" Then
            wsDiag.Cells(diagRow, 1).Value = r
            wsDiag.Cells(diagRow, 2).Value = cellA
            wsDiag.Cells(diagRow, 3).Value = cellC
            wsDiag.Cells(diagRow, 4).Value = cellE
            wsDiag.Cells(diagRow, 5).Value = cellF
            wsDiag.Cells(diagRow, 6).Value = cellH
            wsDiag.Cells(diagRow, 7).Value = judge
            wsDiag.Cells(diagRow, 8).Value = status
            diagRow = diagRow + 1
        End If
        
    Next r
    
    wsDiag.Columns("A:H").AutoFit
    
    wsAllergy.Cells.Clear
    With wsAllergy
        .Range("A1:G1").Value = Array("日付", "昼夕", "利用者名", _
            "弁当種類", "禁食・アレルギー内容", "サイズ", "個数")
        With .Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(255, 230, 220)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    If outCount > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To outCount, 1 To 7)
        Dim a As Long, b As Long
        For a = 1 To outCount
            For b = 1 To 7
                writeArr(a, b) = outArr(a, b)
            Next b
        Next a
        wsAllergy.Range("A2").Resize(outCount, 7).Value = writeArr
        wsAllergy.Range("A2").Resize(outCount, 1).NumberFormat = "m/d"
        wsAllergy.Range("A2").Resize(outCount, 7).Borders.LineStyle = xlContinuous
    End If
    
    wsAllergy.Columns("A:G").AutoFit
    
    Call CleanupAndExit
    If outCount = 0 Then
        MsgBox "0件の取り込みになりました。" & vbCrLf & vbCrLf & _
               "「_診断」シートに各行の判定結果を出力しました。", vbExclamation, "調理指示表取込結果"
    Else
        MsgBox outCount & " 件の禁食・アレルギーデータを取り込みました。", vbInformation, "調理指示表取込完了"
    End If
    Exit Sub
    
ErrHandler:
    Call CleanupAndExit
    MsgBox "調理指示表取込でエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub


'================================================================
' v33 ADD: 「_スキップリスト」シートから除外リストを読み込み
'================================================================
Public Sub LoadSkipList(ByRef skipList() As String, ByRef skipCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    skipCount = 0
    ReDim skipList(1 To 2, 1 To 100)
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_スキップリスト")
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For r = 2 To lastRow
        Dim nm As String, bt As String
        nm = Trim(CStr(ws.Cells(r, 1).Value & ""))
        bt = Trim(CStr(ws.Cells(r, 2).Value & ""))
        If nm <> "" Then
            skipCount = skipCount + 1
            If skipCount > UBound(skipList, 2) Then
                ReDim Preserve skipList(1 To 2, 1 To skipCount + 50)
            End If
            skipList(1, skipCount) = nm
            skipList(2, skipCount) = bt
        End If
    Next r
    On Error GoTo 0
End Sub


'================================================================
' v33 ADD: 利用者名+弁当種類がスキップリストに該当するか判定
'   名前は完全一致、弁当種類は部分一致(空欄ならその名前全件スキップ)
'================================================================
Public Function IsInSkipList(ByVal userName As String, ByVal bentoType As String, _
                              ByRef skipList() As String, ByVal skipCount As Long) As Boolean
    Dim i As Long
    IsInSkipList = False
    If skipCount = 0 Then Exit Function
    If userName = "" Then Exit Function
    
    For i = 1 To skipCount
        If skipList(1, i) = userName Then
            ' 弁当種類が空欄なら、その名前は全件スキップ
            If skipList(2, i) = "" Then
                IsInSkipList = True
                Exit Function
            End If
            ' 部分一致
            If InStr(1, bentoType, skipList(2, i), vbBinaryCompare) > 0 Then
                IsInSkipList = True
                Exit Function
            End If
        End If
    Next i
End Function


'================================================================
' v38 ADD: 検証結果シートをA4横印刷向けに整形
'   - フォントサイズ拡大、行高拡大、列幅最適化、折り返し、印刷設定
'================================================================
Public Sub FormatResultForPrint(ByVal wsResult As Worksheet, ByVal resultCount As Long)
    
    On Error Resume Next
    
    ' === フォント設定 ===
    With wsResult.Range("A1:M1")  ' ヘッダー
        .Font.Name = "MS Pゴシック"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    If resultCount > 0 Then
        With wsResult.Range("A2").Resize(resultCount, 13)  ' データ行
            .Font.Name = "MS Pゴシック"
            .Font.Size = 11
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        ' 「前回」N列の赤字太字は維持
        With wsResult.Range("M2").Resize(resultCount, 1)
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
        End With
        
       
        
        ' 中央寄せにすべき列(日付、昼夕、番号、コース、ポケット、サイズ、個数)
        With wsResult.Range("A2").Resize(resultCount, 1)
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("B2").Resize(resultCount, 1)
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("C2").Resize(resultCount, 1)
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("D2").Resize(resultCount, 1)
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("I2").Resize(resultCount, 1)  ' 該当ポケット
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("K2").Resize(resultCount, 1)  ' サイズ
            .HorizontalAlignment = xlCenter
        End With
        With wsResult.Range("L2").Resize(resultCount, 1)  ' 個数
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    ' === 列幅(A4横210mm想定で、印刷時に1ページに収まる比率) ===
    ' 合計目安 100 程度を9列+5列(計14列)に振り分け
    wsResult.Columns("A").ColumnWidth = 6     ' 日付
    wsResult.Columns("B").ColumnWidth = 5     ' 昼夕
    wsResult.Columns("C").ColumnWidth = 5     ' 番号
    wsResult.Columns("D").ColumnWidth = 5     ' コース
    wsResult.Columns("E").ColumnWidth = 11    ' 利用者名
    wsResult.Columns("F").ColumnWidth = 13    ' 弁当種類
    wsResult.Columns("G").ColumnWidth = 20    ' 禁食・アレルギー内容(K削減分を回す)
    wsResult.Columns("H").ColumnWidth = 14    ' 該当ワード(K削減分を回す)
    wsResult.Columns("I").ColumnWidth = 5     ' 該当ポケット
    wsResult.Columns("J").ColumnWidth = 17    ' 該当食材(K削減分を回す)
    wsResult.Columns("K").ColumnWidth = 6     ' サイズ
    wsResult.Columns("L").ColumnWidth = 5     ' 個数
    wsResult.Columns("M").ColumnWidth = 5     ' 前回
    
    ' === 行の高さ ===
    wsResult.Rows(1).RowHeight = 32  ' ヘッダー
    If resultCount > 0 Then
        wsResult.Rows("2:" & (resultCount + 1)).RowHeight = 28
    End If
    
    ' === 印刷設定 ===
    With wsResult.PageSetup
        .Orientation = xlLandscape          ' 横向き
        .PaperSize = xlPaperA4              ' A4
        .Zoom = False                       ' 拡大率指定をやめる
        .FitToPagesWide = 1                 ' 横1ページに収める
        .FitToPagesTall = False             ' 縦は何ページでも可
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintTitleRows = "$1:$1"           ' 1行目を各ページに繰り返し印刷
        .PrintGridlines = False
        .CenterHeader = "&""MS Pゴシック,太字""&14 禁食・アレルギー検証結果"
        .RightHeader = "&""MS Pゴシック""&10 印刷日: &D"
        .CenterFooter = "&""MS Pゴシック""&10 - &P / &N -"
    End With
    
    ' 印刷範囲を自動設定(データ部のみ)
    If resultCount > 0 Then
        wsResult.PageSetup.PrintArea = "$A$1:$M$" & (resultCount + 1)
    End If
    
    On Error GoTo 0
End Sub


'================================================================
' v33 ADD: 検証結果シートの昼夕境目・日付境目に太罫線
'================================================================
Public Sub ApplyResultBoundaryBorders(ByVal wsResult As Worksheet, ByVal resultCount As Long)
    Dim r As Long
    Dim curDate As String, curMeal As String
    Dim prevDate As String, prevMeal As String
    Dim isBoundary As Boolean
    
    If resultCount <= 1 Then Exit Sub
    
    prevDate = ""
    prevMeal = ""
    
    For r = 2 To resultCount + 1
        Dim dateVal As Variant
        dateVal = wsResult.Cells(r, 1).Value
        If IsDate(dateVal) Then
            curDate = Format(CDate(dateVal), "yyyy/mm/dd")
        Else
            curDate = Trim(CStr(dateVal & ""))
        End If
        curMeal = Trim(CStr(wsResult.Cells(r, 2).Value & ""))
        
        isBoundary = False
        If r > 2 Then
            If curDate <> prevDate Or curMeal <> prevMeal Then
                isBoundary = True
            End If
        End If
        
        If isBoundary Then
            With wsResult.Range(wsResult.Cells(r, 1), wsResult.Cells(r, 13)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(0, 0, 0)
            End With
        End If
        
        prevDate = curDate
        prevMeal = curMeal
    Next r
End Sub


'================================================================
' v32 ADD: 「_設定」シートのB1セルから現在のメニュー範囲モードを取得
'   戻り値: True  = 普通食のみモード
'           False = 全メニューモード
'   (デフォルトは「普通食のみ」)
'================================================================
Public Function IsOnlyFutsuMode() As Boolean
    Dim ws As Worksheet
    Dim v As String
    
    IsOnlyFutsuMode = True  ' デフォルト
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_設定")
    If ws Is Nothing Then Exit Function
    
    v = Trim(CStr(ws.Range("B1").Value & ""))
    On Error GoTo 0
    
    If InStr(v, "全メニュー") > 0 Then
        IsOnlyFutsuMode = False
    Else
        IsOnlyFutsuMode = True
    End If
End Function


'================================================================
' v32 ADD: メニュー範囲モードをトグル切替
'   「普通食のみ」⇔「全メニュー」を切り替えてMsgBoxで現在モード表示
'================================================================
Public Sub メニュー範囲切替()
    Dim ws As Worksheet
    Dim newMode As String
    Dim msg As String
    
    On Error GoTo ErrHandler
    
    Call EnsureSheet("_設定")
    Set ws = ThisWorkbook.Worksheets("_設定")
    
    If IsOnlyFutsuMode() Then
        newMode = "全メニュー"
        msg = "「全メニュー」モードに切り替えました。" & vbCrLf & vbCrLf & _
              "次回の取込から、普通食以外のメニュー" & vbCrLf & _
              "(幸たんぱく、たんぱく塩分、透析、やわらか、" & vbCrLf & _
              " ムース、カロリー、消化)も処理されます。"
    Else
        newMode = "普通食のみ"
        msg = "「普通食のみ」モードに切り替えました。" & vbCrLf & vbCrLf & _
              "次回の取込では、普通食以外はスキップされます。"
    End If
    
    ws.Range("A1").Value = "メニュー範囲"
    ws.Range("B1").Value = newMode
    ws.Range("A1:B1").Font.Bold = True
    With ws.Range("B1")
        If newMode = "全メニュー" Then
            .Interior.Color = RGB(200, 255, 200)  ' 緑(全)
        Else
            .Interior.Color = RGB(255, 240, 180)  ' 黄(普通食のみ)
        End If
    End With
    ws.Columns("A:B").AutoFit
    
    MsgBox msg, vbInformation, "メニュー範囲切替"
    Exit Sub
    
ErrHandler:
    MsgBox "切替でエラー: " & Err.Description, vbExclamation
End Sub


Public Sub 前日履歴として保存()
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long
    Dim ans As VbMsgBoxResult
    
    On Error GoTo ErrHandler
    
    Call EnsureSheet("禁食アレルギー貼付")
    Set wsSrc = ThisWorkbook.Worksheets("禁食アレルギー貼付")
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "「禁食アレルギー貼付」シートにデータがありません。" & vbCrLf & _
               "先に「調理指示表取込」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ans = MsgBox("現在の「禁食アレルギー貼付」シートの内容を" & vbCrLf & _
                 "「_前日履歴」シートに保存しますか?" & vbCrLf & vbCrLf & _
                 "(既存の前日履歴は上書きされます)", vbYesNo + vbQuestion, "前日履歴として保存")
    If ans <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call EnsureSheet("_前日履歴")
    Set wsDst = ThisWorkbook.Worksheets("_前日履歴")
    
    wsDst.Cells.Clear
    wsDst.Range("A1:G1").Value = Array("日付", "昼夕", "利用者名", _
        "弁当種類", "禁食・アレルギー内容", "サイズ", "個数")
    With wsDst.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    Dim srcData As Variant
    srcData = wsSrc.Range("A2:G" & lastRow).Value
    wsDst.Range("A2").Resize(lastRow - 1, 7).Value = srcData
    wsDst.Range("A2").Resize(lastRow - 1, 1).NumberFormat = "m/d"
    wsDst.Range("A2").Resize(lastRow - 1, 7).Borders.LineStyle = xlContinuous
    wsDst.Columns("A:G").AutoFit
    
    Call CleanupAndExit
    
    MsgBox (lastRow - 1) & " 件のデータを「_前日履歴」シートに保存しました。", vbInformation, "保存完了"
    Exit Sub
    
ErrHandler:
    Call CleanupAndExit
    MsgBox "前日履歴の保存でエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub


'================================================================
' v30 ADD: 「前回」判定用 Dictionary を構築
'   キー: 利用者名 + "|" + 日付(yyyy/mm/dd) + "|" + 昼夕
'================================================================
Public Sub BuildPreviousMealDict(ByRef dict As Object)
    Dim wsCurrent As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String, mealTime As String
    Dim dateVal As Variant
    Dim dateKey As String
    Dim k As String
    
    On Error Resume Next
    Set wsCurrent = Nothing
    Set wsCurrent = ThisWorkbook.Worksheets("禁食アレルギー貼付")
    On Error GoTo 0
    
    If Not wsCurrent Is Nothing Then
        lastRow = wsCurrent.Cells(wsCurrent.Rows.Count, "A").End(xlUp).Row
        For r = 2 To lastRow
            dateVal = wsCurrent.Cells(r, 1).Value
            mealTime = Trim(CStr(wsCurrent.Cells(r, 2).Value & ""))
            nm = Trim(CStr(wsCurrent.Cells(r, 3).Value & ""))
            If nm <> "" And mealTime <> "" Then
                dateKey = ""
                If IsDate(dateVal) Then dateKey = Format(CDate(dateVal), "yyyy/mm/dd")
                If dateKey <> "" Then
                    k = nm & "|" & dateKey & "|" & mealTime
                    If Not dict.Exists(k) Then dict.Add k, True
                End If
            End If
        Next r
    End If
    
    On Error Resume Next
    Set wsHistory = Nothing
    Set wsHistory = ThisWorkbook.Worksheets("_前日履歴")
    On Error GoTo 0
    
    If Not wsHistory Is Nothing Then
        lastRow = wsHistory.Cells(wsHistory.Rows.Count, "A").End(xlUp).Row
        For r = 2 To lastRow
            dateVal = wsHistory.Cells(r, 1).Value
            mealTime = Trim(CStr(wsHistory.Cells(r, 2).Value & ""))
            nm = Trim(CStr(wsHistory.Cells(r, 3).Value & ""))
            If nm <> "" And mealTime <> "" Then
                dateKey = ""
                If IsDate(dateVal) Then dateKey = Format(CDate(dateVal), "yyyy/mm/dd")
                If dateKey <> "" Then
                    k = nm & "|" & dateKey & "|" & mealTime
                    If Not dict.Exists(k) Then dict.Add k, True
                End If
            End If
        Next r
    End If
End Sub


'================================================================
' v30 ADD: 1行分の「前回」判定 → "有" または ""
'   昼の場合: 前日夕にその人の名前があれば「有」
'   夕の場合: 同日昼にその人の名前があれば「有」
'================================================================
Public Function CheckPreviousMeal(ByRef dict As Object, ByVal userName As String, _
                                    ByVal currentDate As Variant, ByVal currentMeal As String) As String
    Dim checkDate As Date
    Dim checkMeal As String
    Dim k As String
    
    CheckPreviousMeal = ""
    If userName = "" Then Exit Function
    If Not IsDate(currentDate) Then Exit Function
    
    Select Case currentMeal
        Case "昼"
            checkDate = CDate(currentDate) - 1
            checkMeal = "夕"
        Case "夕"
            checkDate = CDate(currentDate)
            checkMeal = "昼"
        Case Else
            Exit Function
    End Select
    
    k = userName & "|" & Format(checkDate, "yyyy/mm/dd") & "|" & checkMeal
    If dict.Exists(k) Then CheckPreviousMeal = "有"
End Function


'================================================================
' v29 ADD: 検証結果シートの色付け
'================================================================
Public Sub ApplyResultCellColors(ByVal wsResult As Worksheet, ByVal resultCount As Long)
    Dim r As Long
    Dim wordVal As String
    Dim bentoVal As String
    
    Dim clrDateSeal As Long
    Dim clrNoda As Long
    Dim clrDateSealFont As Long
    
    clrDateSeal = RGB(204, 236, 255)
    clrNoda = RGB(255, 192, 0)
    clrDateSealFont = RGB(0, 102, 204)
    
    If resultCount <= 0 Then Exit Sub
    
    For r = 2 To resultCount + 1
        wordVal = Trim(CStr(wsResult.Cells(r, 8).Value & ""))
        bentoVal = Trim(CStr(wsResult.Cells(r, 6).Value & ""))
        
        If InStr(1, wordVal, "日付シール", vbBinaryCompare) > 0 Then
            wsResult.Range(wsResult.Cells(r, 1), wsResult.Cells(r, 13)).Interior.Color = clrDateSeal
            With wsResult.Cells(r, 8)
                .Font.Color = clrDateSealFont
                .Font.Bold = True
            End With
        End If
        
        If bentoVal = "野田市配食サービス" Then
            With wsResult.Cells(r, 6)
                .Interior.Color = clrNoda
                .Font.Bold = True
            End With
        End If
    Next r
End Sub


'================================================================
' 補助:アプリケーション設定を元に戻す
'================================================================
Private Sub CleanupAndExit()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub


'================================================================
' 補助:シートが存在しなければ作成
'================================================================
Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim found As Boolean
    
    found = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            found = True
            Exit For
        End If
    Next ws
    
    If Not found Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        
        Select Case sheetName
            Case "メニュー貼付"
                ws.Range("A1:L1").Value = Array("日付", "昼夕", "弁当種類", _
                    "ポケットA", "ポケットA'", "ポケットA''", "ポケットB", "ポケットB'", _
                    "ポケットC", "ポケットC'", "ポケットD", "ポケットE")
                ws.Range("A1:L1").Font.Bold = True
                ws.Range("A1:L1").Interior.Color = RGB(220, 230, 241)
                
            Case "禁食アレルギー貼付"
                ws.Range("A1:G1").Value = Array("日付", "昼夕", "利用者名", _
                    "弁当種類", "禁食・アレルギー内容", "サイズ", "個数")
                ws.Range("A1:G1").Font.Bold = True
                ws.Range("A1:G1").Interior.Color = RGB(255, 230, 220)
                
            Case "_前日履歴"
                ws.Range("A1:G1").Value = Array("日付", "昼夕", "利用者名", _
                    "弁当種類", "禁食・アレルギー内容", "サイズ", "個数")
                ws.Range("A1:G1").Font.Bold = True
                ws.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
                
            Case "_設定"
                ' v32: メニュー範囲モード保存用
                ws.Range("A1").Value = "メニュー範囲"
                ws.Range("B1").Value = "普通食のみ"
                ws.Range("A1:B1").Font.Bold = True
                ws.Range("B1").Interior.Color = RGB(255, 240, 180)
                ws.Range("A2").Value = "使い方"
                ws.Range("B2").Value = "「6.メニュー範囲切替」ボタンで普通食のみ/全メニューを切替"
                ws.Columns("A:B").AutoFit
                
            Case "_スキップリスト"
                ' v33: 検証対象から除外する利用者+弁当種類のリスト
                ws.Range("A1:B1").Value = Array("利用者名", "弁当種類(部分一致・空欄なら全件)")
                ws.Range("A1:B1").Font.Bold = True
                ws.Range("A1:B1").Interior.Color = RGB(255, 220, 220)
                ws.Range("A2").Value = "岡本勲"
                ws.Range("B2").Value = "健康ボリューム食"
                ws.Range("D1").Value = "使い方"
                ws.Range("D2").Value = "・A列: 利用者名(完全一致)"
                ws.Range("D3").Value = "・B列: 弁当種類のキーワード(部分一致)"
                ws.Range("D4").Value = "・B列空欄ならその利用者を全件除外"
                ws.Columns("A:B").AutoFit
                ws.Columns("D:D").AutoFit
                
            Case "_除外ペアマスタ"
                ws.Range("A1:B1").Value = Array("親キー", "偽陽性文脈ワード(|区切り)")
                ws.Range("A1:B1").Font.Bold = True
                ws.Range("A1:B1").Interior.Color = RGB(255, 200, 200)
                Dim exDefaults As Variant
                exDefaults = Array( _
                    Array("イモ", "固いもの|硬いもの|かたいもの"), _
                    Array("いも", "固いもの|硬いもの|かたいもの"), _
                    Array("芋", "固いもの|硬いもの|かたいもの"))
                Dim ei As Long
                For ei = LBound(exDefaults) To UBound(exDefaults)
                    ws.Cells(ei + 2, 1).Value = exDefaults(ei)(0)
                    ws.Cells(ei + 2, 2).Value = exDefaults(ei)(1)
                Next ei
                ws.Range("D1").Value = "使い方"
                ws.Range("D2").Value = "・A列: 検証対象の親キー(同義語マスタのキー or 単語)"
                ws.Range("D3").Value = "・B列: そのキーが含まれていても誤マッチとなる文脈ワード"
                ws.Range("D4").Value = "・例: 「イモ」が「固いもの」の中に含まれてヒットするのを防ぐ"
                ws.Columns("A:B").AutoFit
                ws.Columns("D:D").AutoFit
                
            Case "_弁当別禁食マスタ"
                ws.Range("A1:C1").Value = Array("弁当種類(部分一致)", "メニュー名(部分一致)", "NG材料(|区切り)")
                ws.Range("A1:C1").Font.Bold = True
                ws.Range("A1:C1").Interior.Color = RGB(255, 220, 180)
                ws.Range("D1").Value = "使い方"
                ws.Range("D2").Value = "・A列: 弁当種類のキーワード(部分一致) 例: ムース"
                ws.Range("D3").Value = "・B列: メニュー名のキーワード(部分一致) 例: ビーフシチュー"
                ws.Range("D4").Value = "・C列: その弁当×メニューの組み合わせで追加するNG材料(|区切り)"
                ws.Range("D5").Value = "・普通食など他の弁当種類には影響しない"
                ws.Range("A2").Value = "ムースセット食"
                ws.Range("B2").Value = "ビーフシチュー"
                ws.Range("C2").Value = "バター"
                ws.Columns("A:C").AutoFit
                ws.Columns("D:D").AutoFit
                
            Case "検証結果"
                ws.Range("A1:N1").Value = Array("日付", "昼夕", "番号", "コース", _
                    "利用者名", "弁当種類", "禁食・アレルギー内容", "該当ワード", _
                    "該当ポケット", "該当食材", "判定", "サイズ", "個数", "前回")
                ws.Range("A1:N1").Font.Bold = True
                ws.Range("A1:N1").Interior.Color = RGB(200, 220, 255)
                
            Case "_調理指示表貼付"
                
            Case "_メニュー生データ"
                
            Case "_メニュー_普通", "_メニュー_幸たんぱく", "_メニュー_健康ボリューム", _
                 "_メニュー_たんぱく塩分", "_メニュー_透析食", _
                 "_メニュー_やわらか", "_メニュー_ムース", "_メニュー_カロリー", "_メニュー_消化"
                
            Case "_同義語マスタ"
                ws.Range("A1:B1").Value = Array("禁食ワード", "同義語(|区切り)")
                ws.Range("A1:B1").Font.Bold = True
                ws.Range("A1:B1").Interior.Color = RGB(255, 255, 200)
                Dim defaults As Variant
                defaults = Array( _
                    Array("魚", "銀ひらす|さば|鯖|サバ|さわら|サワラ|鮭|サーモン|鯵|アジ|あじ|鰯|イワシ|いわし|鯛|タイ|たら|タラ|鱈|鰈|カレイ|かれい|ぶり|ブリ|鰤|ホッケ|ほっけ|赤魚|あかうお|まぐろ|マグロ|まあじ|金目鯛|きんめだい|にしん|ニシン|鮒|ふな|めばる|メバル|かつお|カツオ|鰹|ししゃも|シシャモ|あゆ|アユ|鮎|うなぎ|ウナギ|穴子|あなご|秋刀魚|サンマ|さんま|鰆|タチウオ|たちうお|ひらめ|ヒラメ|平目|カラスガレイ|フライ★"), _
                    Array("肉", "豚|牛|鶏|鳥|チキン|ポーク|ビーフ|ハム|ソーセージ|ウインナー|ベーコン|ミートボール|ミートソース|ミートローフ|ハンバーグ|メンチ|つくね|しゅうまい|シュウマイ|そぼろ|チャーシュー|生姜焼き|角煮|から揚げ|唐揚げ"), _
                    Array("卵", "玉子|たまご|タマゴ|卵焼|玉子焼|オムレツ|スクランブルエッグ"), _
                    Array("えび", "エビ|海老|シュリンプ"), _
                    Array("かに", "カニ|蟹"), _
                    Array("貝", "あさり|アサリ|しじみ|シジミ|はまぐり|ハマグリ|帆立|ホタテ|ほたて|つぶ貝|サザエ"), _
                    Array("牛乳", "ミルク|生クリーム|チーズ|バター|ヨーグルト"), _
                    Array("そば", "蕎麦|ソバ"), _
                    Array("青魚", "さば|サバ|鯖|あじ|アジ|鯵|いわし|イワシ|鰯|さんま|サンマ|秋刀魚|かつお|カツオ|鰹"), _
                    Array("きのこ", "しいたけ|シイタケ|椎茸|まいたけ|マイタケ|舞茸|しめじ|シメジ|エリンギ|なめこ|ナメコ|マッシュルーム|えのき|エノキ"), _
                    Array("揚げ物", "フライ|から揚げ|唐揚げ|天ぷら|てんぷら|コロッケ|カツ"), _
                    Array("漬物", "漬け|たくあん|タクアン|沢庵|キムチ|奈良漬"))
                Dim di As Long
                For di = LBound(defaults) To UBound(defaults)
                    ws.Cells(di + 2, 1).Value = defaults(di)(0)
                    ws.Cells(di + 2, 2).Value = defaults(di)(1)
                Next di
                ws.Columns("A:B").AutoFit
                
            Case "_メニュー材料補完"
                ws.Range("A1:B1").Value = Array("メニュー名(部分一致)", "隠し材料(|区切り)")
                ws.Range("A1:B1").Font.Bold = True
                ws.Range("A1:B1").Interior.Color = RGB(220, 255, 220)
                Dim ingDefaults As Variant
                ingDefaults = Array( _
                    Array("麻婆茄子", "鶏肉|ひき肉"), _
                    Array("麻婆豆腐", "豚肉|ひき肉"), _
                    Array("肉じゃが", "豚肉"), _
                    Array("肉味噌", "豚肉|ひき肉"), _
                    Array("ハンバーグ", "ひき肉"), _
                    Array("ミートボール", "ひき肉"), _
                    Array("メンチカツ", "ひき肉"), _
                    Array("つくね", "鶏肉|ひき肉"), _
                    Array("しゅうまい", "豚肉|ひき肉"), _
                    Array("シュウマイ", "豚肉|ひき肉"), _
                    Array("そぼろ", "鶏肉|ひき肉"))
                Dim ii As Long
                For ii = LBound(ingDefaults) To UBound(ingDefaults)
                    ws.Cells(ii + 2, 1).Value = ingDefaults(ii)(0)
                    ws.Cells(ii + 2, 2).Value = ingDefaults(ii)(1)
                Next ii
                ws.Columns("A:B").AutoFit
        End Select
    End If
End Sub



'================================================================
' v31 ADD: 全角カッコを半角に正規化
'   全角「（」→半角「(」、全角「）」→半角「)」
'   これにより、後続の処理は半角カッコだけを見ればよい
'================================================================
Private Function NormalizeParen(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, ChrW(65288), "(")  ' （ FULLWIDTH LEFT PARENTHESIS
    t = Replace(t, ChrW(65289), ")")  ' ） FULLWIDTH RIGHT PARENTHESIS
    NormalizeParen = t
End Function


'================================================================
' 補助:禁食内容を「、」「,」で分割
'================================================================
Private Function SplitAllergyContent(ByVal content As String) As Variant
    Dim s As String
    s = content
    
    s = RemoveOkSegments(s)
    
    s = Replace(s, "、", ",")
    s = Replace(s, ",", ",")
    s = Replace(s, "・", ",")
    s = Replace(s, " ", ",")
    s = Replace(s, " ", ",")
    s = Replace(s, "。", ",")
    s = Replace(s, "(", ",")
    s = Replace(s, ")", ",")
    s = Replace(s, "(", ",")
    s = Replace(s, ")", ",")
    
    Dim parts As Variant
    parts = Split(s, ",")
    
    Dim out() As String
    ReDim out(0 To UBound(parts))
    Dim n As Long
    n = 0
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim w As String
        w = Trim(CStr(parts(i)))
        w = StripNgSuffix(w)
        w = StripQuantityPrefix(w)
        w = StripCategorySuffix(w)
        If w <> "" Then
            out(n) = w
            n = n + 1
        End If
    Next i
    
    If n = 0 Then
        SplitAllergyContent = Split("", ",")
    Else
        ReDim Preserve out(0 To n - 1)
        SplitAllergyContent = out
    End If
End Function


Private Function StripNgSuffix(ByVal w As String) As String
    Dim s As String
    s = w
    Dim suffixes As Variant
    suffixes = Array("ダメ", "だめ", "ダメ", "NG", "ng", "Ng", "なし", "ナシ", "不可", "×", "ばつ")
    Dim i As Long
    For i = LBound(suffixes) To UBound(suffixes)
        Dim suf As String
        suf = CStr(suffixes(i))
        If Len(s) >= Len(suf) Then
            If Right(s, Len(suf)) = suf Then
                s = Left(s, Len(s) - Len(suf))
                s = Trim(s)
                Exit For
            End If
        End If
    Next i
    StripNgSuffix = s
End Function


Private Function StripQuantityPrefix(ByVal w As String) As String
    Dim s As String
    s = w
    Dim prefixes As Variant
    prefixes = Array("1人分", "2人分", "3人分", "1人分", "2人分", "3人分", _
                     "1つ", "2つ", "3つ", "1つ", "2つ", "3つ", _
                     "1個", "2個", "3個", "1個", "2個", "3個")
    Dim i As Long
    For i = LBound(prefixes) To UBound(prefixes)
        Dim pre As String
        pre = CStr(prefixes(i))
        If Len(s) >= Len(pre) Then
            If Left(s, Len(pre)) = pre Then
                s = Mid(s, Len(pre) + 1)
                s = Trim(s)
                Exit For
            End If
        End If
    Next i
    StripQuantityPrefix = s
End Function


Private Function StripCategorySuffix(ByVal w As String) As String
    Dim s As String
    s = w
    If Right(s, 1) = "類" Then
        Dim base As String
        base = Left(s, Len(s) - 1)
        If Len(base) >= 1 And Len(base) <= 3 Then
            s = base
        End If
    End If
    StripCategorySuffix = s
End Function


Private Function RemoveOkSegments(ByVal s As String) As String
    Dim result As String
    Dim parenContent As String
    
    ' v31: 全角カッコを半角に正規化してから処理
    result = NormalizeParen(s)
    
    Dim posOpen As Long, posClose As Long
    Dim p As Long
    p = 1
    Do
        posOpen = 0
        Dim p1 As Long, p2 As Long
        p1 = InStr(p, result, "(")
        p2 = InStr(p, result, "(")
        If p1 > 0 And (p2 = 0 Or p1 < p2) Then
            posOpen = p1
        ElseIf p2 > 0 Then
            posOpen = p2
        End If
        
        If posOpen = 0 Then Exit Do
        
        posClose = 0
        Dim c1 As Long, c2 As Long
        c1 = InStr(posOpen, result, ")")
        c2 = InStr(posOpen, result, ")")
        If c1 > 0 And (c2 = 0 Or c1 < c2) Then
            posClose = c1
        ElseIf c2 > 0 Then
            posClose = c2
        End If
        
        If posClose = 0 Or posClose <= posOpen Then Exit Do
        
        parenContent = Mid(result, posOpen + 1, posClose - posOpen - 1)
        
        If InStr(1, parenContent, "OK") > 0 Or InStr(1, parenContent, "ok") > 0 _
           Or InStr(1, parenContent, "Ok") > 0 Or InStr(1, parenContent, "オーケー") > 0 Then
            result = Left(result, posOpen - 1) & "," & Mid(result, posClose + 1)
            p = posOpen + 1
        Else
            p = posClose + 1
        End If
    Loop
    
    RemoveOkSegments = result
End Function


Public Function ExtractOkKeywords(ByVal allergyContent As String) As String
    Dim s As String
    ' v31: 全角カッコを半角に正規化してから処理
    s = NormalizeParen(allergyContent)
    Dim result As String
    result = ""
    
    Dim posOpen As Long, posClose As Long
    Dim p As Long
    p = 1
    Do
        posOpen = 0
        Dim p1 As Long, p2 As Long
        p1 = InStr(p, s, "(")
        p2 = InStr(p, s, "(")
        If p1 > 0 And (p2 = 0 Or p1 < p2) Then
            posOpen = p1
        ElseIf p2 > 0 Then
            posOpen = p2
        End If
        
        If posOpen = 0 Then Exit Do
        
        posClose = 0
        Dim c1 As Long, c2 As Long
        c1 = InStr(posOpen, s, ")")
        c2 = InStr(posOpen, s, ")")
        If c1 > 0 And (c2 = 0 Or c1 < c2) Then
            posClose = c1
        ElseIf c2 > 0 Then
            posClose = c2
        End If
        
        If posClose = 0 Or posClose <= posOpen Then Exit Do
        
        Dim parenContent As String
        parenContent = Mid(s, posOpen + 1, posClose - posOpen - 1)
        
        If InStr(1, parenContent, "OK") > 0 Or InStr(1, parenContent, "ok") > 0 _
           Or InStr(1, parenContent, "Ok") > 0 Or InStr(1, parenContent, "オーケー") > 0 Then
            Dim ok As String
            ok = parenContent
            ok = Replace(ok, "OK", "")
            ok = Replace(ok, "ok", "")
            ok = Replace(ok, "Ok", "")
            ok = Replace(ok, "オーケー", "")
            ok = Replace(ok, "は", "")
            ok = Replace(ok, "、", "|")
            ok = Replace(ok, "・", "|")
            ok = Replace(ok, ",", "|")
            
            Dim oks As Variant
            oks = Split(ok, "|")
            Dim oi As Long
            For oi = LBound(oks) To UBound(oks)
                Dim okw As String
                okw = Trim(CStr(oks(oi)))
                If okw <> "" Then
                    If result <> "" Then result = result & "|"
                    result = result & okw
                End If
            Next oi
        End If
        
        p = posClose + 1
    Loop
    
    ExtractOkKeywords = result
End Function


Public Function FilterBentoConditional(ByVal allergyContent As String, ByVal currentBento As String) As String
    Dim s As String
    ' v31: 全角カッコを半角に正規化してから処理
    s = NormalizeParen(allergyContent)
    
    Dim conditions As Variant
    conditions = Array( _
        Array("消化にやさしい", "消化"), _
        Array("消化食", "消化"), _
        Array("やわらか食", "やわらか"), _
        Array("やわらか", "やわらか"), _
        Array("ムース食", "ムース"), _
        Array("ムース", "ムース"), _
        Array("透析食", "透析"), _
        Array("透析", "透析"))
    
    Dim ci As Long
    For ci = LBound(conditions) To UBound(conditions)
        Dim keyword As String
        Dim bentoMarker As String
        keyword = conditions(ci)(0)
        bentoMarker = conditions(ci)(1)
        
        Dim posKw As Long
        posKw = InStr(1, s, keyword, vbBinaryCompare)
        Do While posKw > 0
            Dim posOpen As Long
            posOpen = 0
            Dim charAfter As String
            charAfter = ""
            If posKw + Len(keyword) <= Len(s) Then
                charAfter = Mid(s, posKw + Len(keyword), 1)
            End If
            
            If charAfter = "(" Or charAfter = "(" Then
                posOpen = posKw + Len(keyword)
                
                Dim posClose As Long
                Dim pc1 As Long, pc2 As Long
                pc1 = InStr(posOpen, s, ")")
                pc2 = InStr(posOpen, s, ")")
                posClose = 0
                If pc1 > 0 And (pc2 = 0 Or pc1 < pc2) Then
                    posClose = pc1
                ElseIf pc2 > 0 Then
                    posClose = pc2
                End If
                
                If posClose > posOpen Then
                    If InStr(1, currentBento, bentoMarker) > 0 Then
                        Dim innerContent As String
                        innerContent = Mid(s, posOpen + 1, posClose - posOpen - 1)
                        s = Left(s, posKw - 1) & "," & innerContent & "," & Mid(s, posClose + 1)
                    Else
                        s = Left(s, posKw - 1) & Mid(s, posClose + 1)
                    End If
                    posKw = InStr(1, s, keyword, vbBinaryCompare)
                Else
                    posKw = InStr(posKw + 1, s, keyword, vbBinaryCompare)
                End If
            Else
                posKw = InStr(posKw + 1, s, keyword, vbBinaryCompare)
            End If
        Loop
    Next ci
    
    FilterBentoConditional = s
End Function


Public Function IsFoodMatchesOkKeywords(ByVal foodName As String, ByVal okKeywords As String, _
                                          ByRef synData() As String, ByVal synCount As Long) As Boolean
    IsFoodMatchesOkKeywords = False
    If foodName = "" Or okKeywords = "" Then Exit Function
    
    Dim oks As Variant
    oks = Split(okKeywords, "|")
    Dim oi As Long
    For oi = LBound(oks) To UBound(oks)
        Dim okw As String
        okw = Trim(CStr(oks(oi)))
        If okw = "" Then GoTo NextOk
        
        If InStr(1, foodName, okw, vbBinaryCompare) > 0 Then
            IsFoodMatchesOkKeywords = True
            Exit Function
        End If
        
        Dim syns As String
        syns = GetSynonyms(okw, synData, synCount)
        If syns <> "" Then
            Dim synArr As Variant
            synArr = Split(syns, "|")
            Dim si As Long
            For si = LBound(synArr) To UBound(synArr)
                Dim syn As String
                syn = Trim(CStr(synArr(si)))
                If syn <> "" Then
                    If InStr(1, foodName, syn, vbBinaryCompare) > 0 Then
                        IsFoodMatchesOkKeywords = True
                        Exit Function
                    End If
                End If
            Next si
        End If
NextOk:
    Next oi
End Function


Private Function CompareDate(ByVal d1 As Variant, ByVal d2 As Variant) As Boolean
    Dim s1 As String, s2 As String
    On Error Resume Next
    If IsDate(d1) And IsDate(d2) Then
        CompareDate = (CDate(d1) = CDate(d2))
        Exit Function
    End If
    s1 = Trim(CStr(d1 & ""))
    s2 = Trim(CStr(d2 & ""))
    CompareDate = (s1 = s2)
    On Error GoTo 0
End Function


Private Function ParseWeekStartDate(ByVal titleVal As Variant) As Date
    Dim s As String
    Dim posY As Long, posM As Long, posD As Long
    Dim y As Long, m As Long, d As Long
    
    On Error Resume Next
    ParseWeekStartDate = 0
    
    If IsDate(titleVal) Then
        ParseWeekStartDate = CDate(titleVal)
        Exit Function
    End If
    
    If IsNumeric(titleVal) Then
        Dim sv As Double
        sv = CDbl(titleVal)
        If sv >= 1 And sv <= 109574 Then
            ParseWeekStartDate = CDate(sv)
            Exit Function
        End If
    End If
    
    s = Trim(CStr(titleVal & ""))
    If s = "" Then Exit Function
    
    posY = InStr(s, "年")
    posM = InStr(s, "月")
    posD = InStr(s, "日")
    
    If posY = 0 Or posM = 0 Or posD = 0 Then
        Exit Function
    End If
    
    If posY < posM And posM < posD Then
        y = CLng(Mid(s, 1, posY - 1))
        m = CLng(Mid(s, posY + 1, posM - posY - 1))
        d = CLng(Mid(s, posM + 1, posD - posM - 1))
        
        If y > 1900 And y < 2200 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            ParseWeekStartDate = DateSerial(y, m, d)
        End If
    End If
    On Error GoTo 0
End Function


Private Sub ParseWeekRange(ByVal titleVal As Variant, ByRef startDate As Date, ByRef endDate As Date, ByRef dayCount As Long)
    Dim s As String
    Dim posY As Long, posM1 As Long, posD1 As Long
    Dim posTilde As Long, posM2 As Long, posD2 As Long
    Dim y As Long, m1 As Long, d1 As Long, m2 As Long, d2 As Long
    Dim restStr As String
    
    On Error Resume Next
    startDate = 0
    endDate = 0
    dayCount = 0
    
    s = Trim(CStr(titleVal & ""))
    If s = "" Then Exit Sub
    
    posY = InStr(s, "年")
    posM1 = InStr(s, "月")
    posD1 = InStr(s, "日")
    If posY = 0 Or posM1 = 0 Or posD1 = 0 Then Exit Sub
    If Not (posY < posM1 And posM1 < posD1) Then Exit Sub
    
    y = CLng(Mid(s, 1, posY - 1))
    m1 = CLng(Mid(s, posY + 1, posM1 - posY - 1))
    d1 = CLng(Mid(s, posM1 + 1, posD1 - posM1 - 1))
    
    If y < 1900 Or y > 2200 Or m1 < 1 Or m1 > 12 Or d1 < 1 Or d1 > 31 Then Exit Sub
    
    startDate = DateSerial(y, m1, d1)
    
    posTilde = InStr(posD1, s, ChrW(65374))
    If posTilde = 0 Then posTilde = InStr(posD1, s, ChrW(126))
    If posTilde = 0 Then posTilde = InStr(posD1, s, ChrW(12316))
    
    If posTilde > 0 Then
        restStr = Mid(s, posTilde + 1)
        posM2 = InStr(restStr, "月")
        posD2 = InStr(restStr, "日")
        If posM2 > 0 And posD2 > 0 And posM2 < posD2 Then
            m2 = CLng(Trim(Mid(restStr, 1, posM2 - 1)))
            d2 = CLng(Trim(Mid(restStr, posM2 + 1, posD2 - posM2 - 1)))
            If m2 >= 1 And m2 <= 12 And d2 >= 1 And d2 <= 31 Then
                If m2 < m1 Then
                    endDate = DateSerial(y + 1, m2, d2)
                Else
                    endDate = DateSerial(y, m2, d2)
                End If
            End If
        End If
    End If
    
    If endDate = 0 Then endDate = startDate
    
    dayCount = CLng(endDate - startDate) + 1
    If dayCount < 1 Then dayCount = 1
    If dayCount > 14 Then dayCount = 7
    
    On Error GoTo 0
End Sub


Private Function CalcDateInWeek(ByVal weekStart As Date, ByVal dayNum As Long) As Date
    Dim i As Long
    
    On Error Resume Next
    
    For i = 0 To 7
        Dim cd As Date
        cd = weekStart + i
        If Day(cd) = dayNum Then
            CalcDateInWeek = cd
            Exit Function
        End If
    Next i
    
    CalcDateInWeek = DateSerial(Year(weekStart), Month(weekStart), dayNum)
    
    On Error GoTo 0
End Function


Public Sub LoadIngredientSupplement(ByRef dict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_メニュー材料補完")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For r = 2 To lastRow
        Dim menuKey As String
        Dim hiddenIngredients As String
        menuKey = Trim(CStr(ws.Cells(r, 1).Value & ""))
        hiddenIngredients = Trim(CStr(ws.Cells(r, 2).Value & ""))
        
        If menuKey <> "" And hiddenIngredients <> "" Then
            If Not dict.Exists(menuKey) Then
                dict.Add menuKey, hiddenIngredients
            End If
        End If
    Next r
End Sub

Public Function ExpandFoodName(ByVal foodName As String, ByRef dict As Object) As String
    ExpandFoodName = foodName
    If foodName = "" Then Exit Function
    If dict Is Nothing Then Exit Function
    
    Dim addedIngredients As String
    addedIngredients = ""
    
    ' v40: 完全一致が存在する場合、それを最優先して部分一致は無視する
    '      例: 「豆腐と野菜のそぼろあん」が登録されていれば、「そぼろ」の部分一致は無視
    If dict.Exists(foodName) Then
        addedIngredients = " " & Replace(CStr(dict(foodName)), "|", " ")
        ExpandFoodName = foodName & addedIngredients
        Exit Function
    End If
    
    Dim k As Variant
    For Each k In dict.Keys
        Dim menuKey As String
        menuKey = CStr(k)
        If menuKey <> "" Then
            If InStr(1, foodName, menuKey, vbBinaryCompare) > 0 Then
                Dim ingredients As String
                ingredients = CStr(dict(k))
                addedIngredients = addedIngredients & " " & Replace(ingredients, "|", " ")
            End If
        End If
    Next k
    
    If addedIngredients <> "" Then
        ExpandFoodName = foodName & addedIngredients
    End If
End Function


Public Sub BuildNumberCourseDictionary(ByRef dict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim curMeal As String   ' v43: 現在の昼夕ブロックを追跡
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_調理指示表貼付")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    curMeal = ""
    For r = 1 To lastRow
        Dim cellH As String
        cellH = Trim(CStr(ws.Cells(r, 8).Value & ""))
        ' v43: ブロックヘッダー(○○一覧表)で昼夕を更新
        If InStr(cellH, "一覧表") > 0 Then
            curMeal = NormalizeMealTime(Trim(CStr(ws.Cells(r, 5).Value & "")))
        End If
        
        Dim nm As String, num As String, course As String
        nm = Trim(CStr(ws.Cells(r, 3).Value & ""))
        num = Trim(CStr(ws.Cells(r, 2).Value & ""))
        course = Trim(CStr(ws.Cells(r, 4).Value & ""))
        
        If nm <> "" And nm <> "名前" Then
            ' v43: 名前+昼夕 をキーにする(昼夕でコースが違っても正しく引ける)
            Dim k As String
            k = nm & "|" & curMeal
            If Not dict.Exists(k) Then
                dict.Add k, num & "|" & course
            End If
        End If
    Next r
End Sub


Public Function IsMeatKeyword(ByVal word As String) As Boolean
    IsMeatKeyword = False
    If word = "" Then Exit Function
    
    Dim meatKeywords As Variant
    meatKeywords = Array("肉", "鶏", "豚", "牛", "鳥", _
                         "鶏肉", "豚肉", "牛肉", "鳥肉", _
                         "ひき肉", "挽肉", "挽き肉", _
                         "チキン", "ポーク", "ビーフ", _
                         "鴨", "ラム", "羊", "マトン")
    Dim i As Long
    For i = LBound(meatKeywords) To UBound(meatKeywords)
        If word = CStr(meatKeywords(i)) Then
            IsMeatKeyword = True
            Exit Function
        End If
    Next i
End Function


Public Function IsProcessedOrGroundMeat(ByVal foodName As String) As Boolean
    IsProcessedOrGroundMeat = False
    If foodName = "" Then Exit Function
    
    Dim keywords As Variant
    keywords = Array("ハンバーグ", "ミートボール", "メンチ", "つくね", "しゅうまい", _
                     "シュウマイ", "焼売", "そぼろ", _
                     "ひき肉", "挽肉", "挽き肉", _
                     "ハム", "ソーセージ", "ウインナー", "ウィンナー", _
                     "ベーコン", "チャーシュー")
    Dim i As Long
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, foodName, CStr(keywords(i)), vbBinaryCompare) > 0 Then
            IsProcessedOrGroundMeat = True
            Exit Function
        End If
    Next i
End Function


Public Function HasMeatShapeRestriction(ByVal allergyContent As String) As Boolean
    HasMeatShapeRestriction = False
    If allergyContent = "" Then Exit Function
    
    If InStr(1, allergyContent, "固形") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, allergyContent, "かたまり") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, allergyContent, "塊") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, allergyContent, "形ある") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, allergyContent, "形のある") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    
    Dim okKw As String
    okKw = ExtractOkKeywords(allergyContent)
    If okKw = "" Then Exit Function
    
    If InStr(1, okKw, "加工") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, okKw, "ひき肉") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
    If InStr(1, okKw, "挽肉") > 0 Then
        HasMeatShapeRestriction = True
        Exit Function
    End If
End Function


Public Function NormalizeBentoType(ByVal bento As String) As String
    Dim s As String
    s = Trim(CStr(bento & ""))
    
    ' v31: 普通食グループ - 全角/半角カッコ・「使捨」「回収」を全て普通食扱い
    If s = "普通食" Or _
       s = "野田市配食サービス" Or _
       s = "野田市配食サービス(使捨)" Or _
       s = "野田市配食サービス（使捨）" Or _
       s = "普通食(回収)" Or _
       s = "普通食（回収）" Then
        NormalizeBentoType = "普通食"
        Exit Function
    End If
    
    NormalizeBentoType = s
End Function


Public Function IsTofuProduct(ByVal foodName As String) As Boolean
    IsTofuProduct = False
    If foodName = "" Then Exit Function
    
    If InStr(1, foodName, "豆腐", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "豆乳", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "がんも", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "卯の花", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "おから", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "厚揚げ", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "油揚げ", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
    If InStr(1, foodName, "大豆ミート", vbBinaryCompare) > 0 Then
        IsTofuProduct = True
        Exit Function
    End If
End Function


Public Function IsPlantBasedMeat(ByVal foodName As String) As Boolean
    IsPlantBasedMeat = False
    If foodName = "" Then Exit Function
    
    If InStr(1, foodName, "畑のお肉", vbBinaryCompare) > 0 Then
        IsPlantBasedMeat = True
        Exit Function
    End If
    If InStr(1, foodName, "大豆ミート", vbBinaryCompare) > 0 Then
        IsPlantBasedMeat = True
        Exit Function
    End If
    If InStr(1, foodName, "ソイミート", vbBinaryCompare) > 0 Then
        IsPlantBasedMeat = True
        Exit Function
    End If
End Function

' v41: _メニュー材料補完 にメニュー名フル登録されたB列(隠し材料)に、
'       指定ワードが含まれているかを判定する。
'       含まれている場合、代替肉/豆製品の自動抑制(IsPlantBasedMeat/IsTofuProduct)
'       より優先して検出を許可する目的で使用する。
'       例: A列=「茄子と大豆ミートのカレー風味」、B列=「鶏肉|ひき肉」 のとき、
'           HasExplicitIngredient("茄子と大豆ミートのカレー風味", "鶏肉", dict) = True
Public Function HasExplicitIngredient(ByVal foodName As String, _
                                       ByVal word As String, _
                                       ByRef dict As Object) As Boolean
    HasExplicitIngredient = False
    If foodName = "" Or word = "" Then Exit Function
    If dict Is Nothing Then Exit Function
    If Not dict.Exists(foodName) Then Exit Function
    If InStr(1, CStr(dict(foodName)), word, vbBinaryCompare) > 0 Then
        HasExplicitIngredient = True
    End If
End Function
Private Function IsHashiAkiraNoKizamiNeed(ByVal userName As String, _
                                          ByVal allergyContent As String, _
                                          ByVal foodName As String) As Boolean
    Dim nm As String
    Dim c As String
    Dim f As String
    
    IsHashiAkiraNoKizamiNeed = False
    
    nm = CStr(userName & "")
    nm = Replace(nm, "　", "")
    nm = Replace(nm, " ", "")
    nm = Replace(nm, "（", "(")
    nm = Replace(nm, "）", ")")
    
    c = CStr(allergyContent & "")
    c = Replace(c, "　", "")
    c = Replace(c, " ", "")
    
    f = CStr(foodName & "")
    
    ' 橋 章さん以外には適用しない
    If InStr(1, nm, "橋章", vbBinaryCompare) = 0 Then Exit Function
    
    ' 「牛・豚・固いもの中きざみ」のような刻み判断メモだけ対象
    If InStr(1, c, "牛", vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, c, "豚", vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, c, "きざみ", vbBinaryCompare) = 0 And _
       InStr(1, c, "刻み", vbBinaryCompare) = 0 Then Exit Function
    
    ' 牛・豚が関係ない商品は対象外
    If InStr(1, f, "牛", vbBinaryCompare) = 0 And _
       InStr(1, f, "豚", vbBinaryCompare) = 0 And _
       InStr(1, f, "ビーフ", vbBinaryCompare) = 0 And _
       InStr(1, f, "ポーク", vbBinaryCompare) = 0 Then Exit Function
    
    ' ひき肉・ミンチ・コロッケ・ハンバーグ等は刻まなくてよいので抽出しない
    If IsNoKizamiNeededMeatProduct(f) Then
        IsHashiAkiraNoKizamiNeed = True
        Exit Function
    End If
End Function


Private Function IsNoKizamiNeededMeatProduct(ByVal foodName As String) As Boolean
    Dim keywords As Variant
    Dim i As Long
    
    IsNoKizamiNeededMeatProduct = False
    
    keywords = Array( _
        "ひき肉", "挽肉", "挽き肉", "ミンチ", _
        "ハンバーグ", "ミートボール", "肉団子", _
        "メンチ", "メンチカツ", _
        "コロッケ", _
        "つくね", _
        "しゅうまい", "シュウマイ", "焼売", _
        "そぼろ", _
        "餃子", "ギョーザ", _
        "春巻", "春巻き", _
        "肉味噌", "肉みそ" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, foodName, CStr(keywords(i)), vbBinaryCompare) > 0 Then
            IsNoKizamiNeededMeatProduct = True
            Exit Function
        End If
    Next i
End Function


'================================================================
' v39 ADD: 汎用ポケット限定指定パーサー
'   禁食内容を解析し、「キーワード → 限定ポケットリスト」のDictionaryを返す
'
'   例: "1つ魚ダメ(Aポケットのみ)、(奥様)"
'       → {"魚": "A"}
'
'   例: "肉(A,Bのみ)、卵(Cだけ)"
'       → {"肉": "A|B", "卵": "C"}
'
'   ロジック:
'     1. カッコ(全角/半角)を見つけたら、その直前の文字列からキーワードを抜く
'        (直前の「、」「,」「 」「。」「(」「)」または文頭まで遡る)
'     2. カッコ内に「のみ」「だけ」を含む場合、A～E(およびA',A'',B',C')を抽出
'     3. そのキーワードをDictionaryに登録
'================================================================
Public Sub ParsePocketRestrictions(ByVal allergyContent As String, ByRef dict As Object)
    Dim s As String
    Dim posOpen As Long, posClose As Long
    Dim segment As String
    Dim i As Long, ch As String
    Dim p As Long
    
    If allergyContent = "" Then Exit Sub
    
    s = NormalizeParen(allergyContent)
    s = NormalizePocketText(s)
    
    p = 1
    Do
        posOpen = InStr(p, s, "(")
        If posOpen = 0 Then Exit Do
        
        posClose = InStr(posOpen, s, ")")
        If posClose = 0 Or posClose <= posOpen Then Exit Do
        
        segment = Mid(s, posOpen + 1, posClose - posOpen - 1)
        segment = NormalizePocketText(segment)
        
        If InStr(1, segment, "のみ", vbBinaryCompare) > 0 Or _
           InStr(1, segment, "だけ", vbBinaryCompare) > 0 Then
            
            Dim restrictStr As String
            restrictStr = ""
            
            For i = 1 To Len(segment)
                ch = Mid(segment, i, 1)
                
                If ch = "A" Or ch = "B" Or ch = "C" Or ch = "D" Or ch = "E" Then
                    Dim nextCh As String, nextCh2 As String
                    Dim pktName As String
                    
                    nextCh = ""
                    nextCh2 = ""
                    If i < Len(segment) Then nextCh = Mid(segment, i + 1, 1)
                    If i + 1 < Len(segment) Then nextCh2 = Mid(segment, i + 2, 1)
                    
                    If nextCh = "'" And nextCh2 = "'" Then
                        pktName = ch & "''"
                    ElseIf nextCh = "'" Then
                        pktName = ch & "'"
                    Else
                        pktName = ch
                    End If
                    
                    If Not PocketListContains(restrictStr, pktName) Then
                        If restrictStr <> "" Then restrictStr = restrictStr & "|"
                        restrictStr = restrictStr & pktName
                    End If
                End If
            Next i
            
            Dim kw As String
            kw = ExtractKeywordBeforeParen(s, posOpen)
            kw = CleanRestrictionKeyword(kw)
            
            If kw <> "" And restrictStr <> "" Then
                If Not dict.Exists(kw) Then dict.Add kw, restrictStr
                
                ' 魚ダメ(Aポケットのみ) のような場合は、必ず「魚」でも登録
                If InStr(1, kw, "魚", vbBinaryCompare) > 0 Then
                    If Not dict.Exists("魚") Then dict.Add "魚", restrictStr
                End If
                
                If InStr(1, kw, "肉", vbBinaryCompare) > 0 Then
                    If Not dict.Exists("肉") Then dict.Add "肉", restrictStr
                End If
                
                If InStr(1, kw, "卵", vbBinaryCompare) > 0 Then
                    If Not dict.Exists("卵") Then dict.Add "卵", restrictStr
                End If
                
                If InStr(1, kw, "豆", vbBinaryCompare) > 0 Then
                    If Not dict.Exists("豆") Then dict.Add "豆", restrictStr
                End If
            End If
        End If
        
        p = posClose + 1
    Loop
End Sub

'================================================================
' v39 ADD: カッコ直前のキーワードを抽出
'   位置 parenPos の左側を、区切り文字(、,。 ()NG表現)まで遡る
'   さらに「ダメ」「だめ」「NG」「なし」「不可」「×」などのNG指示語を末尾から除去
'   さらに数量プレフィックス(1つ、2人分等)を先頭から除去
'================================================================
Public Function ExtractKeywordBeforeParen(ByVal s As String, ByVal parenPos As Long) As String
    Dim i As Long
    Dim startPos As Long
    Dim ch As String
    
    startPos = 1
    ' parenPos の直前から左に向かって、区切り文字を探す
    For i = parenPos - 1 To 1 Step -1
        ch = Mid(s, i, 1)
        If ch = "、" Or ch = "," Or ch = " " Or ch = " " Or ch = "。" Or _
           ch = "(" Or ch = ")" Or ch = vbLf Or ch = vbCr Or ch = vbTab Then
            startPos = i + 1
            Exit For
        End If
    Next i
    
    Dim kw As String
    kw = Trim(Mid(s, startPos, parenPos - startPos))
    
    ' 末尾のNG表現除去
    Dim suffixes As Variant
    suffixes = Array("ダメ", "だめ", "ダメ", "NG", "ng", "Ng", "なし", "ナシ", "不可", "×", "ばつ")
    Dim si As Long
    For si = LBound(suffixes) To UBound(suffixes)
        Dim suf As String
        suf = CStr(suffixes(si))
        If Len(kw) >= Len(suf) Then
            If Right(kw, Len(suf)) = suf Then
                kw = Trim(Left(kw, Len(kw) - Len(suf)))
                Exit For
            End If
        End If
    Next si
    
    ' 先頭の数量プレフィックス除去
    Dim prefixes As Variant
    prefixes = Array("1人分", "2人分", "3人分", "1人分", "2人分", "3人分", _
                     "1つ", "2つ", "3つ", "1つ", "2つ", "3つ", _
                     "1個", "2個", "3個", "1個", "2個", "3個")
    Dim pi As Long
    For pi = LBound(prefixes) To UBound(prefixes)
        Dim pre As String
        pre = CStr(prefixes(pi))
        If Len(kw) >= Len(pre) Then
            If Left(kw, Len(pre)) = pre Then
                kw = Trim(Mid(kw, Len(pre) + 1))
                Exit For
            End If
        End If
    Next pi
    
    ExtractKeywordBeforeParen = kw
End Function


'================================================================
' v39 ADD: 指定したワード(または親キー)に対するポケット限定文字列を取得
'   word: 検証中の禁食ワード(例「魚」「銀ひらす」)
'   masterWord: 同義語マスタの親キー(例「魚」)。直接マッチ時は ""
'   dict: ParsePocketRestrictionsで作成したDictionary
'
'   戻り値: パイプ区切りのポケットリスト、なければ ""
'================================================================
Public Function GetPocketRestriction(ByVal word As String, ByVal masterWord As String, _
                                       ByRef dict As Object) As String
    GetPocketRestriction = ""
    If dict Is Nothing Then Exit Function
    
    If word <> "" Then
        If dict.Exists(word) Then
            GetPocketRestriction = CStr(dict(word))
            Exit Function
        End If
    End If
    
    If masterWord <> "" Then
        If dict.Exists(masterWord) Then
            GetPocketRestriction = CStr(dict(masterWord))
            Exit Function
        End If
    End If
End Function


'================================================================
' 補助:禁食内容から「肉(○ポケットのみ)」等のポケット限定指定を抽出
'   ※v39で汎用版ParsePocketRestrictionsが追加されたが、
'     既存ロジックとの互換のため、この関数も「肉」キーワード用に残す
'================================================================
Public Function ParseMeatPocketRestriction(ByVal allergyContent As String) As String
    Dim s As String
    Dim posMeat As Long, posOpen As Long, posClose As Long
    Dim segment As String
    Dim restrictStr As String
    Dim i As Long, ch As String
    
    ParseMeatPocketRestriction = ""
    If allergyContent = "" Then Exit Function
    ' v31: 全角カッコを半角に正規化してから処理
    s = NormalizeParen(allergyContent)
    
    posMeat = InStr(1, s, "肉", vbBinaryCompare)
    Do While posMeat > 0
        Dim charAfterMeat As String
        charAfterMeat = ""
        If posMeat + 1 <= Len(s) Then
            charAfterMeat = Mid(s, posMeat + 1, 1)
        End If
        
        posOpen = 0
        If charAfterMeat = "(" Or charAfterMeat = "(" Then
            posOpen = posMeat + 1
        End If
        
        If posOpen > 0 Then
            Dim posClose1 As Long, posClose2 As Long
            posClose1 = InStr(posOpen, s, ")")
            posClose2 = InStr(posOpen, s, ")")
            
            posClose = 0
            If posClose1 > 0 And (posClose2 = 0 Or posClose1 < posClose2) Then
                posClose = posClose1
            ElseIf posClose2 > 0 Then
                posClose = posClose2
            End If
            
            If posClose > posOpen + 1 Then
                segment = Mid(s, posOpen + 1, posClose - posOpen - 1)
                
                If InStr(1, segment, "のみ") > 0 Or InStr(1, segment, "だけ") > 0 Then
                    For i = 1 To Len(segment)
                        ch = Mid(segment, i, 1)
                        If ch = "A" Or ch = "B" Or ch = "C" Or ch = "D" Or ch = "E" Then
                            Dim nextCh As String
                            nextCh = ""
                            If i < Len(segment) Then nextCh = Mid(segment, i + 1, 1)
                            
                            Dim pktName As String
                            If nextCh = "'" Or nextCh = "'" Then
                                pktName = ch & "'"
                            Else
                                pktName = ch
                            End If
                            
                            If InStr(1, restrictStr, pktName) = 0 Then
                                If restrictStr <> "" Then restrictStr = restrictStr & "|"
                                restrictStr = restrictStr & pktName
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        
        posMeat = InStr(posMeat + 1, s, "肉", vbBinaryCompare)
    Loop
    
    ParseMeatPocketRestriction = restrictStr
End Function


Public Sub AnalyzeMeatRestriction(ByVal allergyContent As String, _
                                   ByRef hasMeatPrefix As Boolean, _
                                   ByRef meatTargets As String)
    Dim s As String
    Dim pos As Long
    Dim prevCh As String
    Dim addedTori As Boolean, addedButa As Boolean, addedGyuu As Boolean
    Dim addedHiki As Boolean, addedKamo As Boolean, addedRam As Boolean
    
    hasMeatPrefix = False
    meatTargets = ""
    
    If allergyContent = "" Then Exit Sub
    s = allergyContent
    
    Dim hasPlantMeat As Boolean
    hasPlantMeat = False
    If InStr(1, s, "畑のお肉", vbBinaryCompare) > 0 Then hasPlantMeat = True
    
    If hasPlantMeat Then
        hasMeatPrefix = True
        meatTargets = "畑のお肉"
    End If
    
    pos = InStr(1, s, "肉", vbBinaryCompare)
    Do While pos > 0
        prevCh = ""
        If pos >= 2 Then prevCh = Mid(s, pos - 1, 1)
        Dim prevCh2 As String
        prevCh2 = ""
        If pos >= 3 Then prevCh2 = Mid(s, pos - 2, 1)
        
        If prevCh = "鶏" Or prevCh = "鳥" Then
            If Not addedTori Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "鶏|チキン|鳥"
                addedTori = True
            End If
            hasMeatPrefix = True
        ElseIf prevCh = "り" And prevCh2 = "と" Then
            If Not addedTori Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "鶏|チキン|鳥"
                addedTori = True
            End If
            hasMeatPrefix = True
        End If
        
        If prevCh = "豚" Then
            If Not addedButa Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "豚|ポーク|ハム|ソーセージ|ベーコン"
                addedButa = True
            End If
            hasMeatPrefix = True
        ElseIf prevCh = "た" And prevCh2 = "ぶ" Then
            If Not addedButa Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "豚|ポーク|ハム|ソーセージ|ベーコン"
                addedButa = True
            End If
            hasMeatPrefix = True
        End If
        
        If prevCh = "牛" Then
            If Not addedGyuu Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "牛|ビーフ"
                addedGyuu = True
            End If
            hasMeatPrefix = True
        ElseIf prevCh = "う" And prevCh2 = "ゅ" Then
            If Not addedGyuu Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "牛|ビーフ"
                addedGyuu = True
            End If
            hasMeatPrefix = True
        End If
        
        If prevCh = "挽" Then
            If Not addedHiki Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "ひき肉|挽肉|挽き肉|ミートボール|ハンバーグ|つくね"
                addedHiki = True
            End If
            hasMeatPrefix = True
        ElseIf prevCh = "き" And prevCh2 = "ひ" Then
            If Not addedHiki Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "ひき肉|挽肉|挽き肉|ミートボール|ハンバーグ|つくね"
                addedHiki = True
            End If
            hasMeatPrefix = True
        End If
        
        If prevCh = "鴨" Then
            If Not addedKamo Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "鴨"
                addedKamo = True
            End If
            hasMeatPrefix = True
        End If
        
        If prevCh = "羊" Then
            If Not addedRam Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "ラム|羊|マトン"
                addedRam = True
            End If
            hasMeatPrefix = True
        ElseIf prevCh = "ム" And prevCh2 = "ラ" Then
            If Not addedRam Then
                If meatTargets <> "" Then meatTargets = meatTargets & "|"
                meatTargets = meatTargets & "ラム|羊|マトン"
                addedRam = True
            End If
            hasMeatPrefix = True
        End If
        
        pos = InStr(pos + 1, s, "肉", vbBinaryCompare)
    Loop
End Sub


Public Sub LoadSynonyms(ByRef synData() As String, ByRef synCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    On Error Resume Next
    
    synCount = 0
    ReDim synData(1 To 2, 1 To 200)
    
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_同義語マスタ")
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For r = 2 To lastRow
        Dim k As String, v As String
        k = Trim(CStr(ws.Cells(r, 1).Value & ""))
        v = Trim(CStr(ws.Cells(r, 2).Value & ""))
        If k <> "" And v <> "" Then
            synCount = synCount + 1
            If synCount > UBound(synData, 2) Then
                ReDim Preserve synData(1 To 2, 1 To synCount + 100)
            End If
            synData(1, synCount) = k
            synData(2, synCount) = v
        End If
    Next r
    
    On Error GoTo 0
End Sub


Public Function GetSynonyms(ByVal word As String, ByRef synData() As String, ByVal synCount As Long) As String
    Dim i As Long
    Dim normWord As String
    GetSynonyms = ""
    If synCount = 0 Then Exit Function
    
    normWord = NormalizeKana(word)
    
    For i = 1 To synCount
        If NormalizeKana(synData(1, i)) = normWord Then
            GetSynonyms = synData(2, i)
            Exit Function
        End If
    Next i
End Function


Private Function IsValidDayNumber(ByVal v As Variant) As Boolean
    On Error Resume Next
    IsValidDayNumber = False
    
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Trim(CStr(v & "")) = "" Then Exit Function
    
    If Not IsNumeric(v) Then Exit Function
    
    If VarType(v) = vbDate Then Exit Function
    
    Dim n As Double
    n = CDbl(v)
    
    If n = 0 Then Exit Function
    
    If n <> Int(n) Then Exit Function
    
    If n >= 1 And n <= 31 Then
        IsValidDayNumber = True
    End If
    On Error GoTo 0
End Function


Private Function GuessBentoType(ByVal sheetName As String) As String
    Dim s As String
    s = sheetName
    
    If InStr(s, "幸たんぱく") > 0 Then
        GuessBentoType = "幸たんぱく食"
    ElseIf InStr(s, "たんぱく") > 0 And InStr(s, "塩分") > 0 Then
        GuessBentoType = "日替 たんぱく・塩分調整食"
    ElseIf InStr(s, "たんぱく") > 0 Then
        GuessBentoType = "たんぱく調整食"
    ElseIf InStr(s, "カロリー") > 0 Then
        GuessBentoType = "カロリー蓋"
    ElseIf InStr(s, "透析") > 0 Then
        GuessBentoType = "透析食"
    Else
        GuessBentoType = "普通食"
    End If
End Function


Private Function ParseDateFromHeader(ByVal headerStr As String) As Date
    Dim s As String
    Dim slash1 As Long, slash2 As Long
    Dim parenPos As Long
    Dim y As Long, m As Long, d As Long
    
    On Error Resume Next
    ParseDateFromHeader = 0
    
    s = Trim(headerStr)
    
    parenPos = InStr(s, "(")
    If parenPos = 0 Then parenPos = InStr(s, "(")
    If parenPos > 0 Then s = Trim(Left(s, parenPos - 1))
    
    slash1 = InStr(s, "/")
    If slash1 > 0 Then
        slash2 = InStr(slash1 + 1, s, "/")
        If slash2 > 0 Then
            y = CLng(Mid(s, 1, slash1 - 1))
            m = CLng(Mid(s, slash1 + 1, slash2 - slash1 - 1))
            d = CLng(Mid(s, slash2 + 1))
            If y > 1900 And y < 2200 Then
                ParseDateFromHeader = DateSerial(y, m, d)
            End If
        End If
    End If
    
    If ParseDateFromHeader = 0 And IsDate(s) Then
        ParseDateFromHeader = CDate(s)
    End If
    
    On Error GoTo 0
End Function


Private Function NormalizeMealTime(ByVal mealStr As String) As String
    Dim s As String
    s = Trim(mealStr)
    
    If InStr(s, "昼") > 0 Then
        NormalizeMealTime = "昼"
    ElseIf InStr(s, "夕") > 0 Then
        NormalizeMealTime = "夕"
    Else
        NormalizeMealTime = s
    End If
End Function


Private Function CleanUserName(ByVal s As String) As String
    Dim t As String
    t = s
    Do While Len(t) > 0 And (Right(t, 1) = " " Or Right(t, 1) = " ")
        t = Left(t, Len(t) - 1)
    Loop
    Do While Len(t) > 0 And (Left(t, 1) = " " Or Left(t, 1) = " ")
        t = Mid(t, 2)
    Loop
    CleanUserName = t
End Function
Private Function NormalizePocketText(ByVal s As String) As String
    s = Replace(s, "Ａ", "A")
    s = Replace(s, "Ｂ", "B")
    s = Replace(s, "Ｃ", "C")
    s = Replace(s, "Ｄ", "D")
    s = Replace(s, "Ｅ", "E")
    
    s = Replace(s, "’", "'")
    s = Replace(s, "‘", "'")
    s = Replace(s, "_", "'")
    s = Replace(s, "　", " ")
    
    NormalizePocketText = s
End Function


Private Function CleanRestrictionKeyword(ByVal kw As String) As String
    kw = Trim(kw)
    kw = NormalizePocketText(kw)
    
    Dim suffixes As Variant
    suffixes = Array("ダメ", "だめ", "NG", "ng", "Ng", "なし", "ナシ", "不可", "×", "ばつ")
    
    Dim i As Long
    For i = LBound(suffixes) To UBound(suffixes)
        Dim suf As String
        suf = CStr(suffixes(i))
        If Len(kw) >= Len(suf) Then
            If Right(kw, Len(suf)) = suf Then
                kw = Trim(Left(kw, Len(kw) - Len(suf)))
                Exit For
            End If
        End If
    Next i
    
    Dim prefixes As Variant
    prefixes = Array("1人分", "2人分", "3人分", "１人分", "２人分", "３人分", _
                     "1つ", "2つ", "3つ", "１つ", "２つ", "３つ", _
                     "1個", "2個", "3個", "１個", "２個", "３個")
    
    For i = LBound(prefixes) To UBound(prefixes)
        Dim pre As String
        pre = CStr(prefixes(i))
        If Len(kw) >= Len(pre) Then
            If Left(kw, Len(pre)) = pre Then
                kw = Trim(Mid(kw, Len(pre) + 1))
                Exit For
            End If
        End If
    Next i
    
    CleanRestrictionKeyword = kw
End Function


Private Function PocketListContains(ByVal listText As String, ByVal pocketName As String) As Boolean
    Dim arr As Variant
    Dim i As Long
    
    PocketListContains = False
    If listText = "" Or pocketName = "" Then Exit Function
    
    arr = Split(listText, "|")
    For i = LBound(arr) To UBound(arr)
        If Trim(CStr(arr(i))) = pocketName Then
            PocketListContains = True
            Exit Function
        End If
    Next i
End Function


Private Function IsPocketAllowedByRestriction(ByVal restrictStr As String, ByVal pocketName As String) As Boolean
    If restrictStr = "" Then
        IsPocketAllowedByRestriction = True
    Else
        IsPocketAllowedByRestriction = PocketListContains(restrictStr, pocketName)
    End If
End Function


'================================================================
' 補助:初期セットアップ(ボタン配置・ガイドシート作成)
'================================================================
Public Sub セットアップ_検証ボタン配置()
    ' v37: 全ボタンを Sheet1(操作パネル) に配置 + 説明書き
    Dim ws As Worksheet
    Dim btn As Object
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' 必要シートを全部作成
    Call EnsureSheet("メニュー貼付")
    Call EnsureSheet("禁食アレルギー貼付")
    Call EnsureSheet("検証結果")
    Call EnsureSheet("_調理指示表貼付")
    Call EnsureSheet("_メニュー_普通")
    Call EnsureSheet("_メニュー_幸たんぱく")
    Call EnsureSheet("_メニュー_健康ボリューム")
    Call EnsureSheet("_メニュー_たんぱく塩分")
    Call EnsureSheet("_メニュー_透析食")
    Call EnsureSheet("_メニュー_やわらか")
    Call EnsureSheet("_メニュー_ムース")
    Call EnsureSheet("_メニュー_カロリー")
    Call EnsureSheet("_メニュー_消化")
    Call EnsureSheet("_同義語マスタ")
    Call EnsureSheet("_メニュー材料補完")
    Call EnsureSheet("_前日履歴")
    Call EnsureSheet("_設定")
    Call EnsureSheet("_スキップリスト")
    Call EnsureSheet("_除外ペアマスタ")
    
    ' Sheet1を取得(なければ作る)
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Sheet1"
    End If
    
    ' Sheet1 を一旦クリア(既存のボタンと内容を全削除)
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 旧版で「禁食アレルギー貼付」シートに置いていたボタンも削除
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = Nothing
    Set wsOld = ThisWorkbook.Worksheets("禁食アレルギー貼付")
    On Error GoTo ErrHandler
    If Not wsOld Is Nothing Then
        Dim shpOld As Shape
        For Each shpOld In wsOld.Shapes
            If shpOld.Name = "btnMenuFolderImport" Or _
               shpOld.Name = "btnMenuImport" Or _
               shpOld.Name = "btnAllergyImport" Or _
               shpOld.Name = "btn検証実行" Or _
               shpOld.Name = "btn前日履歴保存" Or _
               shpOld.Name = "btnメニュー範囲切替" Or _
               Left(shpOld.Name, 6) = "Button" Or _
               Left(shpOld.Name, 4) = "ボタン" Then
                shpOld.Delete
            End If
        Next shpOld
    End If
    
    ' === Sheet1 の構築 ===
    ws.Activate
    
    ' タイトル
    With ws.Range("A1")
        .Value = "禁食・アレルギー検証システム 操作パネル"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
    End With
    ws.Range("A1:F1").Merge
    ws.Range("A1:F1").HorizontalAlignment = xlCenter
    ws.Rows(1).RowHeight = 30
    
    ' 説明セクション: 運用フロー
    With ws.Range("A3")
        .Value = "■ 基本の運用フロー"
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
    End With
    ws.Range("A3:F3").Merge
    
    ws.Range("A4").Value = "(1) 毎月最初: 「1.メニューフォルダから取込」でメニューを一括取込"
    ws.Range("A5").Value = "(2) 毎日: 「_調理指示表貼付」シートに調理指示表を貼り付け"
    ws.Range("A6").Value = "(3) 「3.調理指示表取込」 → 「4.検証実行」 の順に押す"
    ws.Range("A7").Value = "(4) 翌日のために 「5.前日履歴として保存」 を押す"
    ws.Range("A4:F7").Font.Color = RGB(50, 50, 50)
    
    ' ボタン配置の基準位置
    Dim btnLeft As Double, btnTop As Double
    Dim btnWidth As Double, btnHeight As Double
    Dim descCol As String
    btnLeft = ws.Range("A9").Left
    btnTop = ws.Range("A9").Top
    btnWidth = 180
    btnHeight = 32
    
    ' --- 1.メニューフォルダから取込 ---
    With ws.Range("A9")
        .Value = "【1】メニューフォルダから取込"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A10").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btnMenuFolderImport"
    btn.Caption = "1.メニューフォルダから取込"
    btn.OnAction = "メニューフォルダから取込"
    ws.Range("C10").Value = "マクロブックと同じフォルダの「メニュー」フォルダ内のメニュー表(.xlsx)を一括で取り込みます。"
    ws.Range("C11").Value = "ファイル名で弁当種類を自動判定(普通、幸、健康、塩分、透析、やわらか、ムース、カロリー、消化)。"
    
    ' --- 2.メニュー取込のみ ---
    With ws.Range("A13")
        .Value = "【2】メニュー取込のみ"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A14").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btnMenuImport"
    btn.Caption = "2.メニュー取込のみ"
    btn.OnAction = "メニュー取込"
    ws.Range("C14").Value = "「_メニュー_xxx」シートにすでに貼り付けた生データから「メニュー貼付」シートを作成。"
    ws.Range("C15").Value = "(フォルダ取込を使わず、手動で各シートに貼り付けて運用する場合に使用)"
    
    ' --- 3.調理指示表取込 ---
    With ws.Range("A17")
        .Value = "【3】調理指示表取込"
        .Font.Bold = True
        .Interior.Color = RGB(217, 234, 211)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A18").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btnAllergyImport"
    btn.Caption = "3.調理指示表取込"
    btn.OnAction = "調理指示表取込"
    ws.Range("C18").Value = "「_調理指示表貼付」シートに貼り付けた当日の調理指示表から、"
    ws.Range("C19").Value = "利用者ごとの禁食・アレルギー情報を「禁食アレルギー貼付」シートに整理。"
    
    ' --- 4.検証実行 ---
    With ws.Range("A21")
        .Value = "【4】検証実行"
        .Font.Bold = True
        .Interior.Color = RGB(252, 213, 180)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A22").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btn検証実行"
    btn.Caption = "4.検証実行"
    btn.OnAction = "禁食アレルギー検証実行"
    ws.Range("C22").Value = "メニューと禁食内容を照合し、「検証結果」シートに「要対応」項目を出力。"
    ws.Range("C23").Value = "日付シール対象者は水色行、野田市配食サービス(回収容器)はオレンジセル、"
    ws.Range("C24").Value = "前回(前日夕or同日昼)に名前があれば「前回」列に赤字で「有」を表示。"
    
    ' --- 5.前日履歴として保存 ---
    With ws.Range("A26")
        .Value = "【5】前日履歴として保存"
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A27").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btn前日履歴保存"
    btn.Caption = "5.前日履歴として保存"
    btn.OnAction = "前日履歴として保存"
    ws.Range("C27").Value = "現在の「禁食アレルギー貼付」を「_前日履歴」シートに保存。"
    ws.Range("C28").Value = "翌日の検証で「前回」列の判定に使われます。退勤前に押すのが推奨。"
    
    ' --- 6.メニュー範囲切替 ---
    With ws.Range("A30")
        .Value = "【6】メニュー範囲切替"
        .Font.Bold = True
        .Interior.Color = RGB(217, 226, 243)
    End With
    Set btn = ws.Buttons.Add(Left:=btnLeft, Top:=ws.Range("A31").Top, _
                              Width:=btnWidth, Height:=btnHeight)
    btn.Name = "btnメニュー範囲切替"
    btn.Caption = "6.メニュー範囲切替"
    btn.OnAction = "メニュー範囲切替"
    ws.Range("C31").Value = "「普通食のみ」と「全メニュー」をトグル切替。"
    ws.Range("C32").Value = "現在モード: " & IIf(IsOnlyFutsuMode(), "普通食のみ", "全メニュー")
    ws.Range("C33").Value = "テスト時は「普通食のみ」、本番運用時は「全メニュー」を推奨。"
    
    ' --- 補助シートの説明 ---
    With ws.Range("A35")
        .Value = "■ 補助シート(必要時に編集)"
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
    End With
    ws.Range("A35:F35").Merge
    
    ws.Range("A36").Value = "・_同義語マスタ"
    ws.Range("C36").Value = "禁食ワードと類似メニューの対応表(例: 魚 → さば、サーモン...)"
    
    ws.Range("A37").Value = "・_メニュー材料補完"
    ws.Range("C37").Value = "メニュー名から見えない隠し材料(例: 麻婆茄子 → 鶏肉、ひき肉)"
    
    ws.Range("A38").Value = "・_スキップリスト"
    ws.Range("C38").Value = "検証対象から除外する利用者+弁当種類のリスト(部分一致可)"
    
    ws.Range("A39").Value = "・_設定"
    ws.Range("C39").Value = "メニュー範囲モード(B1セル)を保存"
    
    ws.Range("A40").Value = "・_診断"
    ws.Range("C40").Value = "調理指示表取込時の各行の判定結果(0件のときの調査用)"
    
    ws.Range("A41").Value = "・_前日履歴"
    ws.Range("C41").Value = "ボタン5で保存される前日の禁食データ"
    
    ' 列幅調整
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 2
    ws.Columns("C").ColumnWidth = 80
    
    ws.Range("A1").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "セットアップ完了。" & vbCrLf & vbCrLf & _
           "Sheet1 を操作パネルとして整備しました。" & vbCrLf & _
           "今後はSheet1のボタンを使ってください。" & vbCrLf & vbCrLf & _
           "現在のメニュー範囲: " & IIf(IsOnlyFutsuMode(), "普通食のみ", "全メニュー"), vbInformation
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "セットアップでエラー: " & Err.Description, vbExclamation
End Sub
'================================================================
' カタカナをひらがなに変換(マッチ判定用の正規化)
'   全角カタカナ(ァ-ヶ)を全角ひらがな(ぁ-?)に変換
'   半角カタカナは事前にStrConvで全角化してから変換
'================================================================
Public Function NormalizeKana(ByVal s As String) As String
    Dim t As String
    Dim i As Long
    Dim ch As Long
    Dim result As String
    
    If s = "" Then
        NormalizeKana = ""
        Exit Function
    End If
    
    ' まず半角カタカナを全角カタカナに変換
    On Error Resume Next
    t = StrConv(s, vbWide)
    If Err.Number <> 0 Then t = s
    On Error GoTo 0
    
    ' 全角カタカナ(ァ=12449～ヶ=12534)を全角ひらがな(ぁ=12353～?=12438)に変換
    ' オフセットは -96
    result = ""
    For i = 1 To Len(t)
        ch = AscW(Mid(t, i, 1))
        If ch >= 12449 And ch <= 12534 Then
            result = result & ChrW(ch - 96)
        Else
            result = result & Mid(t, i, 1)
        End If
    Next i
    
    NormalizeKana = result
End Function
'================================================================
' 日番号判定(0や時刻も含めた緩めの判定)
'   パターン2の日番号検出に使用。0や日付シリアル値も除外。
'================================================================
Private Function IsValidDayNumberLoose(ByVal v As Variant) As Boolean
    On Error Resume Next
    IsValidDayNumberLoose = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Trim(CStr(v & "")) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If VarType(v) = vbDate Then Exit Function
    
    Dim n As Double
    n = CDbl(v)
    If n = 0 Then Exit Function
    If n <> Int(n) Then Exit Function
    If n >= 1 And n <= 31 Then IsValidDayNumberLoose = True
    On Error GoTo 0
End Function
'================================================================
' 除外ペアマスタを読み込む
'   配列の構造: exData(1, n) = 親キー、 exData(2, n) = 偽陽性文脈ワード(|区切り)
'================================================================
Public Sub LoadExclusionPairs(ByRef exData() As String, ByRef exCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    exCount = 0
    ReDim exData(1 To 2, 1 To 100)
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_除外ペアマスタ")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For r = 2 To lastRow
        Dim k As String, v As String
        k = Trim(CStr(ws.Cells(r, 1).Value & ""))
        v = Trim(CStr(ws.Cells(r, 2).Value & ""))
        If k <> "" And v <> "" Then
            exCount = exCount + 1
            If exCount > UBound(exData, 2) Then
                ReDim Preserve exData(1 To 2, 1 To exCount + 50)
            End If
            exData(1, exCount) = k
            exData(2, exCount) = v
        End If
    Next r
End Sub


'================================================================
' 偽陽性判定
'   親キー masterWord が allergyContent にヒットしたが、
'   除外ペアマスタに登録された「文脈ワード」も同時に
'   allergyContent に含まれているなら → True(偽陽性なので除外)
'   ただし、その文脈ワードを除去した残り文字列にも親キーが
'   独立して出現する場合は、偽陽性ではないとして残す。
'================================================================
Public Function IsFalsePositiveByContext(ByVal masterWord As String, _
                                          ByVal allergyContent As String, _
                                          ByVal foodNameForMatch As String, _
                                          ByRef exData() As String, _
                                          ByVal exCount As Long) As Boolean
    Dim i As Long
    Dim normContent As String
    Dim normMaster As String
    Dim normFood As String

    IsFalsePositiveByContext = False
    If masterWord = "" Or exCount = 0 Then Exit Function

    normMaster = NormalizeKana(masterWord)
    normContent = NormalizeKana(allergyContent)
    normFood = NormalizeKana(foodNameForMatch)

    For i = 1 To exCount
        If NormalizeKana(exData(1, i)) = normMaster Then
            Dim ctxList As String
            ctxList = exData(2, i)
            If ctxList <> "" Then
                Dim ctxs As Variant
                ctxs = Split(ctxList, "|")
                Dim ci As Long
                For ci = LBound(ctxs) To UBound(ctxs)
                    Dim ctx As String
                    ctx = Trim(CStr(ctxs(ci)))
                    If ctx <> "" Then
                        Dim normCtx As String
                        normCtx = NormalizeKana(ctx)
                        
                        ' メニュー名側で文脈ワードがヒット → メニュー名で独立出現チェック
                        If InStr(1, normFood, normCtx, vbBinaryCompare) > 0 Then
                            If Not IsMasterWordIndependent(masterWord, foodNameForMatch, ctx) Then
                                IsFalsePositiveByContext = True
                                Exit Function
                            End If
                        End If
                        
                        ' 禁食欄側で文脈ワードがヒット → 禁食欄で独立出現チェック(従来動作)
                        If InStr(1, normContent, normCtx, vbBinaryCompare) > 0 Then
                            If Not IsMasterWordIndependent(masterWord, allergyContent, ctx) Then
                                IsFalsePositiveByContext = True
                                Exit Function
                            End If
                        End If
                    End If
                Next ci
            End If
        End If
    Next i
End Function


'================================================================
' 親キーが文脈ワードの中だけでなく、独立して出現しているか確認
'   allergyContent から ctx を全て除去した残り文字列に
'   masterWord (or その同義語) が含まれていれば独立出現あり
'================================================================
Public Function IsMasterWordIndependent(ByVal masterWord As String, _
                                          ByVal allergyContent As String, _
                                          ByVal ctx As String) As Boolean
    Dim s As String
    Dim normMaster As String
    Dim normCtx As String
    
    IsMasterWordIndependent = False
    
    s = NormalizeKana(allergyContent)
    normMaster = NormalizeKana(masterWord)
    normCtx = NormalizeKana(ctx)
    
    Do While InStr(1, s, normCtx, vbBinaryCompare) > 0
        s = Replace(s, normCtx, "")
    Loop
    
    If InStr(1, s, normMaster, vbBinaryCompare) > 0 Then
        IsMasterWordIndependent = True
    End If
End Function

'================================================================
' v42 ADD: _弁当別禁食マスタ の読み込み
'   bentoRules(1, i) = 弁当種類キーワード（部分一致）
'   bentoRules(2, i) = メニュー名キーワード（部分一致）
'   bentoRules(3, i) = NG材料（|区切り）
'================================================================
Public Sub LoadBentoSpecificRules(ByRef bentoRules() As String, ByRef bentoRuleCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    bentoRuleCount = 0
    ReDim bentoRules(1 To 3, 1 To 10)

    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("_弁当別禁食マスタ")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        Dim bentoKw As String, menuKw As String, ngMat As String
        bentoKw = Trim(CStr(ws.Cells(r, 1).Value & ""))
        menuKw = Trim(CStr(ws.Cells(r, 2).Value & ""))
        ngMat = Trim(CStr(ws.Cells(r, 3).Value & ""))

        If bentoKw <> "" And menuKw <> "" And ngMat <> "" Then
            bentoRuleCount = bentoRuleCount + 1
            If bentoRuleCount > UBound(bentoRules, 2) Then
                ReDim Preserve bentoRules(1 To 3, 1 To bentoRuleCount + 20)
            End If
            bentoRules(1, bentoRuleCount) = bentoKw
            bentoRules(2, bentoRuleCount) = menuKw
            bentoRules(3, bentoRuleCount) = ngMat
        End If
    Next r
End Sub


'================================================================
' v42 ADD: 弁当種類×メニュー名でマッチしたNG材料を追加展開
'   _メニュー材料補完（全弁当共通）とは分離した、弁当種類限定の補完
'================================================================
Public Function ExpandFoodNameWithBentoRule( _
    ByVal foodName As String, _
    ByVal allergyBento As String, _
    ByRef bentoRules() As String, _
    ByVal bentoRuleCount As Long) As String

    ExpandFoodNameWithBentoRule = foodName
    If foodName = "" Or bentoRuleCount = 0 Then Exit Function

    Dim i As Long
    Dim added As String
    added = ""

    For i = 1 To bentoRuleCount
        If InStr(1, allergyBento, bentoRules(1, i), vbBinaryCompare) > 0 Then
            If InStr(1, foodName, bentoRules(2, i), vbBinaryCompare) > 0 Then
                added = added & " " & Replace(bentoRules(3, i), "|", " ")
            End If
        End If
    Next i

    If added <> "" Then
        ExpandFoodNameWithBentoRule = foodName & added
    End If
End Function

