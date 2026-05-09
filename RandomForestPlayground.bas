' =========================================================================
' Macro VBA: Random Forest Statistical Playground
'
' Esta macro cria duas abas:
' 1. Configuração de Random Forest: Para definir X, Y, Períodos e Hiperparâmetros
' 2. Random Forest Insights: Executa o código Python que treina os modelos,
'    faz as inferências e gera as métricas e feature importances.
' =========================================================================

Sub SetupRandomForestPlayground()
    Dim wsConfig As Worksheet
    Dim wsInsights As Worksheet

    ' Criar ou limpar aba de Configuração
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("RF_Config")
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsConfig.Name = "RF_Config"
    Else
        wsConfig.Cells.Clear
    End If
    On Error GoTo 0

    ' Preencher cabeçalhos da Configuração
    With wsConfig
        .Cells(1, 1).Value = "Random Forest Configuration"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14

        .Cells(3, 1).Value = "Parâmetro"
        .Cells(3, 2).Value = "Valor"
        .Cells(3, 3).Value = "Descrição"
        .Range("A3:C3").Font.Bold = True

        ' Dataset
        .Cells(4, 1).Value = "Data Sheet Name"
        .Cells(4, 2).Value = "base_de_dados" ' Preencher com nome da aba real
        .Cells(4, 3).Value = "Nome da aba que contém os dados brutos"

        .Cells(5, 1).Value = "Data Range"
        .Cells(5, 2).Value = "A1:AX200000"
        .Cells(5, 3).Value = "Intervalo dos dados brutos (ex: A1:AZ1000)"

        ' Variáveis
        .Cells(6, 1).Value = "Target (Y)"
        .Cells(6, 2).Value = "Nome_da_coluna_Y"
        .Cells(6, 3).Value = "Coluna a ser prevista"

        .Cells(7, 1).Value = "Problem Type"
        .Cells(7, 2).Value = "Classification" ' ou Regression
        .Cells(7, 3).Value = "Classification (Categorias) ou Regression (Números Contínuos)"

        ' Features (X)
        .Cells(8, 1).Value = "Features (X) - Separadas por vírgula"
        .Cells(8, 2).Value = "Coluna_1, Coluna_2, Coluna_3"
        .Cells(8, 3).Value = "Colunas usadas para prever Y"

        ' Divisão de tempo
        .Cells(9, 1).Value = "Time Column"
        .Cells(9, 2).Value = "Week"
        .Cells(9, 3).Value = "Coluna usada para separar períodos (ex: Semana, Mês)"

        .Cells(10, 1).Value = "Period A"
        .Cells(10, 2).Value = "W01"
        .Cells(10, 3).Value = "Valor para o Período A (Treino/Teste 1)"

        .Cells(11, 1).Value = "Period B"
        .Cells(11, 2).Value = "W02"
        .Cells(11, 3).Value = "Valor para o Período B (Treino/Teste 2)"

        ' Hyperparameters
        .Cells(12, 1).Value = "n_estimators"
        .Cells(12, 2).Value = 100
        .Cells(12, 3).Value = "Número de árvores na floresta (Best practice: 100)"

        .Cells(13, 1).Value = "max_depth"
        .Cells(13, 2).Value = 10
        .Cells(13, 3).Value = "Profundidade máxima da árvore (Best practice: 10 para evitar overfit)"

        .Cells(14, 1).Value = "test_size"
        .Cells(14, 2).Value = 0.2
        .Cells(14, 3).Value = "Proporção de dados para teste em cada período (Best practice: 0.2 ou 20%)"

        .Columns("A:C").AutoFit
    End With

    ' Adicionar botão para rodar
    Dim btn As Object
    Set btn = wsConfig.Buttons.Add(300, 10, 150, 30)
    btn.OnAction = "RunRandomForest"
    btn.Characters.Text = "Run Random Forest"

    MsgBox "Aba 'RF_Config' criada! Preencha os valores de configuração e clique no botão para rodar.", vbInformation
End Sub

Sub RunRandomForest()
    Dim wsConfig As Worksheet
    Dim wsInsights As Worksheet
    Dim dataSheet As String, dataRange As String
    Dim targetCol As String, problemType As String
    Dim featuresStr As String, timeCol As String
    Dim periodA As String, periodB As String
    Dim nEstimators As Integer, maxDepth As Integer, testSize As Double
    Dim pyScript As String

    ' Ler Configurações
    Set wsConfig = ThisWorkbook.Sheets("RF_Config")
    dataSheet = wsConfig.Cells(4, 2).Value
    dataRange = wsConfig.Cells(5, 2).Value
    targetCol = wsConfig.Cells(6, 2).Value
    problemType = wsConfig.Cells(7, 2).Value
    featuresStr = wsConfig.Cells(8, 2).Value
    timeCol = wsConfig.Cells(9, 2).Value
    periodA = wsConfig.Cells(10, 2).Value
    periodB = wsConfig.Cells(11, 2).Value
    nEstimators = wsConfig.Cells(12, 2).Value
    maxDepth = wsConfig.Cells(13, 2).Value
    testSize = wsConfig.Cells(14, 2).Value

    ' Tratar strings de features
    Dim featuresArr() As String
    Dim featuresPythonList As String
    Dim i As Integer
    featuresArr = Split(featuresStr, ",")
    featuresPythonList = "["
    For i = LBound(featuresArr) To UBound(featuresArr)
        featuresPythonList = featuresPythonList & "'" & Trim(featuresArr(i)) & "'"
        If i < UBound(featuresArr) Then featuresPythonList = featuresPythonList & ", "
    Next i
    featuresPythonList = featuresPythonList & "]"

    ' Criar ou limpar aba de Insights
    On Error Resume Next
    Set wsInsights = ThisWorkbook.Sheets("RF_Insights")
    If wsInsights Is Nothing Then
        Set wsInsights = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsInsights.Name = "RF_Insights"
    Else
        wsInsights.Cells.Clear
    End If
    On Error GoTo 0

    wsInsights.Cells(1, 1).Value = "Random Forest Insights"
    wsInsights.Cells(1, 1).Font.Bold = True
    wsInsights.Cells(1, 1).Font.Size = 14
    wsInsights.Cells(2, 1).Value = "Gerando resultados usando Python no Excel..."

    ' Construir o script Python
    pyScript = "import pandas as pd" & vbCrLf
    pyScript = pyScript & "import numpy as np" & vbCrLf
    pyScript = pyScript & "from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor" & vbCrLf
    pyScript = pyScript & "from sklearn.model_selection import train_test_split" & vbCrLf
    pyScript = pyScript & "from sklearn.metrics import accuracy_score, r2_score" & vbCrLf
    pyScript = pyScript & "from sklearn.preprocessing import LabelEncoder" & vbCrLf
    pyScript = pyScript & "import shap" & vbCrLf & vbCrLf

    ' Carregar dados do Excel
    pyScript = pyScript & "df = xl(""""'" & dataSheet & "'!" & dataRange & """", headers=True)" & vbCrLf

    ' Tratamento de nulos (Preencher com média numéricas, moda categóricas)
    pyScript = pyScript & "num_cols = df.select_dtypes(include=[np.number]).columns" & vbCrLf
    pyScript = pyScript & "cat_cols = df.select_dtypes(exclude=[np.number]).columns" & vbCrLf
    pyScript = pyScript & "df[num_cols] = df[num_cols].fillna(df[num_cols].mean())" & vbCrLf
    pyScript = pyScript & "for col in cat_cols:" & vbCrLf
    pyScript = pyScript & "    df[col] = df[col].fillna(df[col].mode()[0] if not df[col].mode().empty else 'Unknown')" & vbCrLf

    ' Variáveis configuradas
    pyScript = pyScript & "target = '" & targetCol & "'" & vbCrLf
    pyScript = pyScript & "features = " & featuresPythonList & vbCrLf
    pyScript = pyScript & "time_col = '" & timeCol & "'" & vbCrLf
    pyScript = pyScript & "period_A = '" & periodA & "'" & vbCrLf
    pyScript = pyScript & "period_B = '" & periodB & "'" & vbCrLf
    pyScript = pyScript & "prob_type = '" & problemType & "'" & vbCrLf

    ' Função auxiliar para treinar e extrair insights
    pyScript = pyScript & "def train_and_eval(data, period_name):" & vbCrLf
    pyScript = pyScript & "    if len(data) < 10:" & vbCrLf
    pyScript = pyScript & "        return {'Error': f'Not enough data for {period_name}'}" & vbCrLf

    ' Label Encoding para features categóricas se houver
    pyScript = pyScript & "    X = data[features].copy()" & vbCrLf
    pyScript = pyScript & "    for col in X.select_dtypes(include=['object']):" & vbCrLf
    pyScript = pyScript & "        le = LabelEncoder()" & vbCrLf
    pyScript = pyScript & "        X[col] = le.fit_transform(X[col].astype(str))" & vbCrLf

    ' Tratamento do Target
    pyScript = pyScript & "    y = data[target].copy()" & vbCrLf
    pyScript = pyScript & "    if prob_type == 'Classification':" & vbCrLf
    pyScript = pyScript & "        if y.dtype == 'object':" & vbCrLf
    pyScript = pyScript & "            le_y = LabelEncoder()" & vbCrLf
    pyScript = pyScript & "            y = le_y.fit_transform(y.astype(str))" & vbCrLf
    pyScript = pyScript & "        model = RandomForestClassifier(n_estimators=" & nEstimators & ", max_depth=" & maxDepth & ", random_state=42)" & vbCrLf
    pyScript = pyScript & "    else:" & vbCrLf
    pyScript = pyScript & "        model = RandomForestRegressor(n_estimators=" & nEstimators & ", max_depth=" & maxDepth & ", random_state=42)" & vbCrLf

    pyScript = pyScript & "    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=" & Replace(CStr(testSize), ",", ".") & ", random_state=42)" & vbCrLf
    pyScript = pyScript & "    model.fit(X_train, y_train)" & vbCrLf
    pyScript = pyScript & "    y_pred = model.predict(X_test)" & vbCrLf

    ' Extrair métricas e importancia
    pyScript = pyScript & "    metric = accuracy_score(y_test, y_pred) if prob_type == 'Classification' else r2_score(y_test, y_pred)" & vbCrLf
    pyScript = pyScript & "    metric_name = 'Accuracy' if prob_type == 'Classification' else 'R-squared'" & vbCrLf

    pyScript = pyScript & "    importances = model.feature_importances_" & vbCrLf
    pyScript = pyScript & "    res = pd.DataFrame({'Feature': features, f'Importance_{period_name}': importances})" & vbCrLf
    pyScript = pyScript & "    res = res.sort_values(by=f'Importance_{period_name}', ascending=False)" & vbCrLf
    pyScript = pyScript & "    res.loc[-1] = ['Metric', f'{metric_name}: {metric:.4f}']" & vbCrLf
    pyScript = pyScript & "    try:" & vbCrLf
    pyScript = pyScript & "        explainer = shap.TreeExplainer(model)" & vbCrLf
    pyScript = pyScript & "        shap_values = explainer.shap_values(X_test)" & vbCrLf
    pyScript = pyScript & "        res.loc[-2] = ['SHAP Status', 'Calculated successfully']" & vbCrLf
    pyScript = pyScript & "    except Exception as e:" & vbCrLf
    pyScript = pyScript & "        res.loc[-2] = ['SHAP Status', f'Failed: {str(e)}']" & vbCrLf
    pyScript = pyScript & "    res.index = res.index + 1" & vbCrLf
    pyScript = pyScript & "    res = res.sort_index()" & vbCrLf
    pyScript = pyScript & "    return res" & vbCrLf

    ' Separar dados por período
    pyScript = pyScript & "data_A = df[df[time_col].astype(str) == period_A]" & vbCrLf
    pyScript = pyScript & "data_B = df[df[time_col].astype(str) == period_B]" & vbCrLf

    pyScript = pyScript & "res_A = train_and_eval(data_A, 'A')" & vbCrLf
    pyScript = pyScript & "res_B = train_and_eval(data_B, 'B')" & vbCrLf

    ' Juntar resultados
    pyScript = pyScript & "if isinstance(res_A, dict):" & vbCrLf
    pyScript = pyScript & "    final_res = pd.DataFrame([res_A])" & vbCrLf
    pyScript = pyScript & "elif isinstance(res_B, dict):" & vbCrLf
    pyScript = pyScript & "    final_res = pd.DataFrame([res_B])" & vbCrLf
    pyScript = pyScript & "else:" & vbCrLf
    pyScript = pyScript & "    final_res = pd.merge(res_A, res_B, on='Feature', how='outer')" & vbCrLf

    ' Output final no Excel
    pyScript = pyScript & "final_res" & vbCrLf

    ' Como inserir o Python no Excel via VBA não possui uma propriedade nativa Cell.Formula2 = "=PY(...)" publicamente documentada da mesma forma que fórmulas normais (o objeto subjacente é restrito),
    ' a abordagem oficial para criar addins é escrever o código e solicitar avaliação ou inserir usando Range.Formula2 (alguns builds suportam =PY(...) via Formula2/Formula3, outros falham).
    ' A solução robusta é colocar a string dentro da célula se =PY for suportado, ou avisar o usuário.

    On Error Resume Next
    ' Tenta inserir como array formula PY. Em builds recentes Formula2 = "=PY(""codigo"")" pode funcionar,
    ' mas como o código tem aspas, múltiplas linhas etc, o tratamento de string é complexo.
    ' Alternativamente, podemos gravar o script gerado para o usuário colar.
    wsInsights.Cells(3, 1).Value = "Copie o código abaixo e cole na célula desejada precedido por =PY( e fechando com )"
    wsInsights.Cells(4, 1).Value = pyScript

    ' Tentar injetar diretamente se suportado
    ' Substituir aspas duplas e quebras de linha para fórmula
    ' Esta abordagem costuma ter limite de caracteres no VBA (8192 chars em Formula2)
    Dim pyFormula As String
    pyFormula = "=PY(""" & Replace(pyScript, """", """""") & """)"
    wsInsights.Cells(5, 1).Formula2 = pyFormula

    If Err.Number <> 0 Then
        wsInsights.Cells(5, 1).Value = "Aviso: Injeção automática =PY() falhou. Por favor, cole o código manualmente ou habilite canal Insider."
    End If
    On Error GoTo 0

    wsInsights.Columns("A:E").AutoFit
    MsgBox "Análise Concluída! Verifique a aba RF_Insights.", vbInformation
End Sub
