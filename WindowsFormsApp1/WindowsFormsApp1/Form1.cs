
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private static readonly string[] TeklaModelDllCandidates =
        {
            @"C:\Program Files\Tekla Structures\2024.0\bin\plugins\Tekla.Structures.Model.dll",
            @"C:\Program Files\Tekla Structures\2024.0\bin\Tekla.Structures.Model.dll",
            @"C:\Program Files\Tekla Structures\2024.0\nt\bin\plugins\Tekla.Structures.Model.dll",
            @"C:\Program Files\Tekla Structures\2024.0\nt\bin\Tekla.Structures.Model.dll"
        };

        private static readonly string[] TeklaDrawingDllCandidates =
        {
            @"C:\Program Files\Tekla Structures\2024.0\bin\plugins\Tekla.Structures.Drawing.dll",
            @"C:\Program Files\Tekla Structures\2024.0\bin\Tekla.Structures.Drawing.dll",
            @"C:\Program Files\Tekla Structures\2024.0\nt\bin\plugins\Tekla.Structures.Drawing.dll",
            @"C:\Program Files\Tekla Structures\2024.0\nt\bin\Tekla.Structures.Drawing.dll"
        };

        private static readonly string[] TeklaStructuresDllCandidates =
        {
            @"C:\Program Files\Tekla Structures\2024.0\bin\Tekla.Structures.dll",
            @"C:\Program Files\Tekla Structures\2024.0\nt\bin\Tekla.Structures.dll"
        };

        private const string ExplodedViewPrefix = "EXP_AUTO_";
        private const int MaxExplodedViews = 80;
        private const bool RotateExplodedToIsometric = true;
        private const string ExplodedViewAttributeFile = "auto-iso";
        private const int MainPartContourColor = 165; // DrawingColors.Magenta
        private const int SidePositiveContourColor = 162; // DrawingColors.Blue
        private const int SideNegativeContourColor = 161; // DrawingColors.Green
        private const int OtherPlaneContourColor = 160; // DrawingColors.Red

        public Form1()
        {
            InitializeComponent();
        }

        private void btnVerificarTekla_Click(object sender, EventArgs e)
        {
            try
            {
                Type modelType = ResolveTeklaModelType();
                if (modelType == null)
                {
                    UpdateStatus("Falha: API do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Model.Model.");
                    return;
                }

                object modelInstance = Activator.CreateInstance(modelType);
                bool? connected = InvokeBoolMethod(modelInstance, "GetConnectionStatus");
                if (!connected.HasValue)
                {
                    UpdateStatus("Falha: metodo GetConnectionStatus indisponivel.", Color.DarkRed);
                    SetOutput("Nao foi possivel ler o status da conexao de modelo.");
                    return;
                }

                if (connected.Value)
                {
                    UpdateStatus("Sucesso: comunicacao com Tekla estabelecida.", Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Falha: Tekla aberto, mas sem conexao com o modelo.", Color.DarkOrange);
                }

                StringBuilder report = new StringBuilder();
                report.AppendLine("Conexao com modelo Tekla");
                report.AppendLine("------------------------");
                report.AppendLine("Tipo carregado: " + modelType.FullName);
                report.AppendLine("GetConnectionStatus: " + (connected.Value ? "True" : "False"));
                SetOutput(report.ToString());
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao comunicar com Tekla: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void btnVerificarDesenho_Click(object sender, EventArgs e)
        {
            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    return;
                }

                bool? drawingConnection = InvokeBoolMethod(drawingHandler, "GetConnectionStatus");
                if (drawingConnection.HasValue && !drawingConnection.Value)
                {
                    UpdateStatus("Falha: sem conexao com Drawing API.", Color.DarkOrange);
                    SetOutput("DrawingHandler.GetConnectionStatus retornou False.");
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                    SetOutput("Nenhum desenho ativo encontrado no Tekla.");
                    return;
                }

                StringBuilder report = new StringBuilder();
                report.AppendLine("Conexao Drawing API: " + BoolToText(drawingConnection));
                report.AppendLine();
                report.AppendLine(
                    BuildObjectData(
                        activeDrawing,
                        "Desenho ativo",
                        new[] { "Name", "Mark", "Title1", "Title2", "Title3", "UpToDateStatus", "Identifier" }));

                SetOutput(report.ToString());
                UpdateStatus("Sucesso: desenho aberto encontrado.", Color.DarkGreen);
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao ler desenho: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void btnVerVistaSelecionada_Click(object sender, EventArgs e)
        {
            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                    SetOutput("Abra um desenho no Tekla antes de consultar a vista.");
                    return;
                }

                int selectedObjectCount;
                object selectedView = TryGetSelectedView(drawingHandler, out selectedObjectCount);

                StringBuilder report = new StringBuilder();
                report.AppendLine(
                    BuildObjectData(
                        activeDrawing,
                        "Desenho ativo",
                        new[] { "Name", "Mark", "Title1", "Title2", "Title3", "UpToDateStatus", "Identifier" }));
                report.AppendLine();
                report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                report.AppendLine();

                if (selectedView == null)
                {
                    report.AppendLine("Nenhuma vista selecionada encontrada.");
                    report.AppendLine("Selecione a moldura da vista ou um objeto dentro da vista e tente novamente.");
                    SetOutput(report.ToString());
                    UpdateStatus("Falha: nenhuma vista selecionada.", Color.DarkOrange);
                    return;
                }

                report.AppendLine(
                    BuildObjectData(
                        selectedView,
                        "Vista selecionada",
                        new[] { "Name", "Scale", "Identifier", "Shortening", "CoordinateSystem" }));

                SetOutput(report.ToString());
                UpdateStatus("Sucesso: vista selecionada encontrada.", Color.DarkGreen);
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao ler vista selecionada: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }
        private void btnGerarVistaExplodida_Click(object sender, EventArgs e)
        {
            bool ghostEnabled = chkGhostLinhas.Checked;
            bool colorizeEnabled = chkColorir.Checked;
            bool lineGuideEnabled = ghostEnabled;
            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                    SetOutput("Abra um desenho no Tekla e selecione uma vista.");
                    return;
                }

                int selectedObjectCount;
                object sourceView = TryGetSelectedView(drawingHandler, out selectedObjectCount);
                if (sourceView == null)
                {
                    UpdateStatus("Falha: nenhuma vista selecionada.", Color.DarkOrange);
                    SetOutput("Selecione a moldura de uma vista (ou um objeto dentro dela) e tente novamente.");
                    return;
                }

                object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
                if (sheet == null)
                {
                    UpdateStatus("Falha: nao foi possivel ler a folha do desenho.", Color.DarkRed);
                    SetOutput("Drawing.GetSheet retornou nulo.");
                    return;
                }

                List<ExplodedPartPlan> partPlans = CollectPartPlansFromView(sourceView);
                if (partPlans.Count == 0)
                {
                    UpdateStatus("Falha: a vista nao possui partes legiveis.", Color.DarkOrange);
                    SetOutput("Nenhum ModelIdentifier de parte foi encontrado na vista selecionada.");
                    return;
                }

                partPlans.Sort(
                    delegate (ExplodedPartPlan left, ExplodedPartPlan right)
                    {
                        return string.Compare(left.IdentifierKey, right.IdentifierKey, StringComparison.Ordinal);
                    });

                int candidateCount = partPlans.Count;
                bool truncated = false;
                int maxPartsAllowed = MaxExplodedViews;
                if (ghostEnabled)
                {
                    maxPartsAllowed = Math.Max(1, (MaxExplodedViews + 1) / 2);
                }

                if (partPlans.Count > maxPartsAllowed)
                {
                    partPlans.RemoveRange(maxPartsAllowed, partPlans.Count - maxPartsAllowed);
                    truncated = true;
                }

                MarkMainPartByLargest2D(partPlans);

                int removedExisting = DeleteExistingExplodedViews(sheet, ExplodedViewPrefix);

                double sourceScale = GetSourceViewScale(sourceView);
                object model = CreateModelInstance();
                int centersFromModel;
                bool modelLayoutUsed = ComputeExplodedOffsets(sourceView, partPlans, sourceScale, model, out centersFromModel);
                bool ghostApplied = ghostEnabled && modelLayoutUsed;
                bool guideLinesApplied = lineGuideEnabled && ghostApplied;

                object anchorOrigin = GetPropertyValue(sourceView, "Origin");
                double anchorX;
                double anchorY;
                double anchorZ;
                if (!TryGetXYZ(anchorOrigin, out anchorX, out anchorY, out anchorZ))
                {
                    UpdateStatus("Falha: nao foi possivel ler origem da vista.", Color.DarkRed);
                    SetOutput("View.Origin nao esta disponivel.");
                    return;
                }

                bool anchorFromMainPart = TryGetMainPartCenter2D(partPlans, out anchorX, out anchorY);

                object viewCoordinateSystem = GetPropertyValue(sourceView, "ViewCoordinateSystem");
                object displayCoordinateSystem = GetPropertyValue(sourceView, "DisplayCoordinateSystem");
                if (viewCoordinateSystem == null || displayCoordinateSystem == null)
                {
                    UpdateStatus("Falha: coordenadas da vista indisponiveis.", Color.DarkRed);
                    SetOutput("ViewCoordinateSystem/DisplayCoordinateSystem nao encontrados.");
                    return;
                }

                int created = 0;
                int failed = 0;
                int autoIsoAppliedCount = 0;
                int fallbackRotationCount = 0;
                int contourColoredCount = 0;
                int contourNotChangedCount = 0;
                int ghostPairsRequested = 0;
                int ghostOriginalCreated = 0;
                int ghostExplodedCreated = 0;
                int guideLinesRequested = 0;
                int guideLinesCreated = 0;
                int guideLinesFailed = 0;
                int viewSequence = 1;

                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];
                    if (ghostApplied && !plan.IsMainPart)
                    {
                        ghostPairsRequested++;

                        bool autoIsoApplied;
                        bool fallbackRotationApplied;
                        bool contourAttempted;
                        bool contourApplied;

                        string originalViewName = ExplodedViewPrefix + viewSequence.ToString("D3") + "_ORG";
                        viewSequence++;
                        bool originalCreated = TryCreateAndPlaceColoredView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan,
                            sourceScale,
                            originalViewName,
                            anchorX + plan.OriginalOffsetX,
                            anchorY + plan.OriginalOffsetY,
                            anchorZ,
                            colorizeEnabled,
                            true,
                            out autoIsoApplied,
                            out fallbackRotationApplied,
                            out contourAttempted,
                            out contourApplied);

                        if (originalCreated)
                        {
                            created++;
                            ghostOriginalCreated++;
                            if (autoIsoApplied)
                            {
                                autoIsoAppliedCount++;
                            }

                            if (fallbackRotationApplied)
                            {
                                fallbackRotationCount++;
                            }

                            if (contourAttempted)
                            {
                                if (contourApplied)
                                {
                                    contourColoredCount++;
                                }
                                else
                                {
                                    contourNotChangedCount++;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        string explodedViewName = ExplodedViewPrefix + viewSequence.ToString("D3") + "_EXP";
                        viewSequence++;
                        bool explodedCreated = TryCreateAndPlaceColoredView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan,
                            sourceScale,
                            explodedViewName,
                            anchorX + plan.OffsetX,
                            anchorY + plan.OffsetY,
                            anchorZ,
                            colorizeEnabled,
                            false,
                            out autoIsoApplied,
                            out fallbackRotationApplied,
                            out contourAttempted,
                            out contourApplied);

                        if (explodedCreated)
                        {
                            created++;
                            ghostExplodedCreated++;
                            if (autoIsoApplied)
                            {
                                autoIsoAppliedCount++;
                            }

                            if (fallbackRotationApplied)
                            {
                                fallbackRotationCount++;
                            }

                            if (contourAttempted)
                            {
                                if (contourApplied)
                                {
                                    contourColoredCount++;
                                }
                                else
                                {
                                    contourNotChangedCount++;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        if (guideLinesApplied)
                        {
                            guideLinesRequested++;
                            bool pairAvailable = originalCreated && explodedCreated;
                            if (pairAvailable
                                && TryCreateGuideLine(
                                    sheet,
                                    anchorX + plan.OriginalOffsetX,
                                    anchorY + plan.OriginalOffsetY,
                                    anchorZ,
                                    anchorX + plan.OffsetX,
                                    anchorY + plan.OffsetY,
                                    anchorZ))
                            {
                                guideLinesCreated++;
                            }
                            else
                            {
                                guideLinesFailed++;
                            }
                        }

                        continue;
                    }

                    bool standardAutoIsoApplied;
                    bool standardFallbackRotationApplied;
                    bool standardContourAttempted;
                    bool standardContourApplied;
                    string viewName = ExplodedViewPrefix + viewSequence.ToString("D3");
                    viewSequence++;
                    bool standardCreated = TryCreateAndPlaceColoredView(
                        sheet,
                        viewCoordinateSystem,
                        displayCoordinateSystem,
                        plan,
                        sourceScale,
                        viewName,
                        anchorX + plan.OffsetX,
                        anchorY + plan.OffsetY,
                        anchorZ,
                        colorizeEnabled,
                        false,
                        out standardAutoIsoApplied,
                        out standardFallbackRotationApplied,
                        out standardContourAttempted,
                        out standardContourApplied);

                    if (!standardCreated)
                    {
                        failed++;
                        continue;
                    }

                    created++;
                    if (standardAutoIsoApplied)
                    {
                        autoIsoAppliedCount++;
                    }

                    if (standardFallbackRotationApplied)
                    {
                        fallbackRotationCount++;
                    }

                    if (standardContourAttempted)
                    {
                        if (standardContourApplied)
                        {
                            contourColoredCount++;
                        }
                        else
                        {
                            contourNotChangedCount++;
                        }
                    }
                }

                InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");

                StringBuilder report = new StringBuilder();
                report.AppendLine("Vista explodida (somente drawing)");
                report.AppendLine("---------------------------------");
                report.AppendLine("Colorir: " + (colorizeEnabled ? "ligado" : "desligado"));
                report.AppendLine("Linha guia: " + (lineGuideEnabled ? (guideLinesApplied ? "ligada" : "solicitada (desativada por fallback)") : "desligada"));
                report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                report.AppendLine("Partes candidatas na vista: " + candidateCount + (truncated ? " (limitado para " + maxPartsAllowed + ")" : string.Empty));
                report.AppendLine("Vistas explodidas antigas removidas: " + removedExisting);
                report.AppendLine("Base de posicionamento: " + (modelLayoutUsed ? "modelo (posicao relativa)" : "drawing (fallback)"));
                report.AppendLine("Anchor de posicionamento: " + (anchorFromMainPart ? "centro da peca principal" : "origem da vista"));
                report.AppendLine("Centros lidos do modelo: " + centersFromModel);
                report.AppendLine("Perfil de vista solicitado: " + ExplodedViewAttributeFile);
                report.AppendLine("Vistas com perfil aplicado: " + autoIsoAppliedCount);
                report.AppendLine("Rotacao isometrica de fallback: " + fallbackRotationCount);
                report.AppendLine("Cor de contorno aplicada: " + contourColoredCount);
                report.AppendLine("Sem mudanca de cor: " + contourNotChangedCount);
                if (guideLinesApplied)
                {
                    report.AppendLine("Linhas guia solicitadas: " + guideLinesRequested);
                    report.AppendLine("Linhas guia criadas: " + guideLinesCreated);
                    report.AppendLine("Linhas guia falhas: " + guideLinesFailed);
                }
                report.AppendLine("Vistas criadas: " + created);
                report.AppendLine("Falhas: " + failed);
                report.AppendLine();
                report.AppendLine("Origem da vista fonte (sheet): X=" + anchorX.ToString("0.###") + " Y=" + anchorY.ToString("0.###"));
                report.AppendLine("Escala base usada: " + sourceScale.ToString("0.###"));
                report.AppendLine();
                report.AppendLine("Observacao: fluxo 100% drawing-only (sem leitura/escrita no modelo).");

                SetOutput(report.ToString());

                if (created > 0)
                {
                    UpdateStatus("Sucesso: vista explodida criada no desenho.", Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Falha: nenhuma vista foi criada.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao gerar vista explodida: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void btnLimparVistaExplodida_Click(object sender, EventArgs e)
        {
            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                    SetOutput("Abra um desenho no Tekla antes de limpar a vista explodida.");
                    return;
                }

                object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
                if (sheet == null)
                {
                    UpdateStatus("Falha: nao foi possivel ler a folha do desenho.", Color.DarkRed);
                    SetOutput("Drawing.GetSheet retornou nulo.");
                    return;
                }

                int removed = DeleteExistingExplodedViews(sheet, ExplodedViewPrefix);
                CommitActiveDrawingChanges(activeDrawing);

                StringBuilder report = new StringBuilder();
                report.AppendLine("Limpeza de vista explodida");
                report.AppendLine("--------------------------");
                report.AppendLine("Prefixo: " + ExplodedViewPrefix);
                report.AppendLine("Vistas removidas: " + removed);
                report.AppendLine("CommitChanges executado: sim");
                report.AppendLine();
                report.AppendLine("Observacao: somente objetos de desenho foram alterados.");
                SetOutput(report.ToString());

                if (removed > 0)
                {
                    UpdateStatus("Sucesso: vistas explodidas removidas.", Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Concluido: nenhuma vista explodida para remover.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao limpar vista explodida: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void btnExplodirPorPlanos_Click(object sender, EventArgs e)
        {
            bool autoDetectEnabled = chkDetectaAutomatico.Checked;
            bool ghostEnabled = chkGhostLinhas.Checked;
            bool colorizeEnabled = chkColorir.Checked;
            bool lineGuideEnabled = ghostEnabled;

            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                    SetOutput("Abra um desenho no Tekla e selecione uma vista.");
                    return;
                }

                int selectedObjectCount = 0;
                int autoViewCount = 0;
                int autoIsoCandidates = 0;
                object sourceView;
                string referenceMode;
                if (autoDetectEnabled)
                {
                    sourceView = TryDetectIsometricReferenceView(activeDrawing, out autoViewCount, out autoIsoCandidates);
                    referenceMode = "automatico";
                    if (sourceView == null)
                    {
                        UpdateStatus("Falha: nenhuma vista isometrica detectada.", Color.DarkOrange);
                        SetOutput("Nao foi possivel detectar automaticamente a vista de conjunto. Desligue 'Detecta automaticamente o conjunto' e selecione uma vista manualmente.");
                        return;
                    }
                }
                else
                {
                    sourceView = TryGetSelectedView(drawingHandler, out selectedObjectCount);
                    referenceMode = "manual";
                    if (sourceView == null)
                    {
                        UpdateStatus("Aguardando selecao de vista.", Color.DarkOrange);
                        SetOutput("Selecione a moldura de uma vista no desenho. Em seguida, clique novamente em 'Explodir por planos (novo)' para gerar.");
                        return;
                    }
                }

                object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
                if (sheet == null)
                {
                    UpdateStatus("Falha: nao foi possivel ler a folha do desenho.", Color.DarkRed);
                    SetOutput("Drawing.GetSheet retornou nulo.");
                    return;
                }

                List<ExplodedPartPlan> partPlans = CollectPartPlansFromView(sourceView);
                if (partPlans.Count == 0)
                {
                    UpdateStatus("Falha: a vista nao possui partes legiveis.", Color.DarkOrange);
                    SetOutput("Nenhum ModelIdentifier de parte foi encontrado na vista selecionada.");
                    return;
                }

                partPlans.Sort(
                    delegate (ExplodedPartPlan left, ExplodedPartPlan right)
                    {
                        return string.Compare(left.IdentifierKey, right.IdentifierKey, StringComparison.Ordinal);
                    });

                int candidateCount = partPlans.Count;
                bool truncated = false;
                int maxPartsAllowed = MaxExplodedViews;
                if (ghostEnabled)
                {
                    maxPartsAllowed = Math.Max(1, (MaxExplodedViews + 1) / 2);
                }

                if (partPlans.Count > maxPartsAllowed)
                {
                    partPlans.RemoveRange(maxPartsAllowed, partPlans.Count - maxPartsAllowed);
                    truncated = true;
                }

                MarkMainPartByLargest2D(partPlans);

                int removedExisting = DeleteExistingExplodedViews(sheet, ExplodedViewPrefix);

                double sourceScale = GetSourceViewScale(sourceView);
                object model = CreateModelInstance();
                int centersFromModel;
                int movedByXY;
                int movedByXZ;
                int movedByZY;
                bool modelLayoutUsed = ComputeExplodedOffsetsByPlanes(
                    sourceView,
                    partPlans,
                    sourceScale,
                    model,
                    true,
                    true,
                    true,
                    out centersFromModel,
                    out movedByXY,
                    out movedByXZ,
                    out movedByZY);

                bool fallbackLayoutUsed = false;
                if (!modelLayoutUsed)
                {
                    int fallbackCenters;
                    ComputeExplodedOffsets(sourceView, partPlans, sourceScale, model, out fallbackCenters);
                    centersFromModel = Math.Max(centersFromModel, fallbackCenters);
                    fallbackLayoutUsed = true;
                }

                bool ghostApplied = ghostEnabled && modelLayoutUsed;
                bool guideLinesApplied = lineGuideEnabled && ghostApplied;

                object anchorOrigin = GetPropertyValue(sourceView, "Origin");
                double anchorX;
                double anchorY;
                double anchorZ;
                if (!TryGetXYZ(anchorOrigin, out anchorX, out anchorY, out anchorZ))
                {
                    UpdateStatus("Falha: nao foi possivel ler origem da vista.", Color.DarkRed);
                    SetOutput("View.Origin nao esta disponivel.");
                    return;
                }

                bool anchorFromMainPart = TryGetMainPartCenter2D(partPlans, out anchorX, out anchorY);

                object viewCoordinateSystem = GetPropertyValue(sourceView, "ViewCoordinateSystem");
                object displayCoordinateSystem = GetPropertyValue(sourceView, "DisplayCoordinateSystem");
                if (viewCoordinateSystem == null || displayCoordinateSystem == null)
                {
                    UpdateStatus("Falha: coordenadas da vista indisponiveis.", Color.DarkRed);
                    SetOutput("ViewCoordinateSystem/DisplayCoordinateSystem nao encontrados.");
                    return;
                }

                int created = 0;
                int failed = 0;
                int autoIsoAppliedCount = 0;
                int fallbackRotationCount = 0;
                int contourColoredCount = 0;
                int contourNotChangedCount = 0;
                int ghostPairsRequested = 0;
                int ghostOriginalCreated = 0;
                int ghostExplodedCreated = 0;
                int guideLinesRequested = 0;
                int guideLinesCreated = 0;
                int guideLinesFailed = 0;
                int viewSequence = 1;

                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];
                    if (ghostApplied && !plan.IsMainPart)
                    {
                        ghostPairsRequested++;

                        bool autoIsoApplied;
                        bool fallbackRotationApplied;
                        bool contourAttempted;
                        bool contourApplied;

                        string originalViewName = ExplodedViewPrefix + viewSequence.ToString("D3") + "_ORG";
                        viewSequence++;
                        bool originalCreated = TryCreateAndPlaceColoredView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan,
                            sourceScale,
                            originalViewName,
                            anchorX + plan.OriginalOffsetX,
                            anchorY + plan.OriginalOffsetY,
                            anchorZ,
                            colorizeEnabled,
                            true,
                            out autoIsoApplied,
                            out fallbackRotationApplied,
                            out contourAttempted,
                            out contourApplied);

                        if (originalCreated)
                        {
                            created++;
                            ghostOriginalCreated++;
                            if (autoIsoApplied)
                            {
                                autoIsoAppliedCount++;
                            }

                            if (fallbackRotationApplied)
                            {
                                fallbackRotationCount++;
                            }

                            if (contourAttempted)
                            {
                                if (contourApplied)
                                {
                                    contourColoredCount++;
                                }
                                else
                                {
                                    contourNotChangedCount++;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        string explodedViewName = ExplodedViewPrefix + viewSequence.ToString("D3") + "_EXP";
                        viewSequence++;
                        bool explodedCreated = TryCreateAndPlaceColoredView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan,
                            sourceScale,
                            explodedViewName,
                            anchorX + plan.OffsetX,
                            anchorY + plan.OffsetY,
                            anchorZ,
                            colorizeEnabled,
                            false,
                            out autoIsoApplied,
                            out fallbackRotationApplied,
                            out contourAttempted,
                            out contourApplied);

                        if (explodedCreated)
                        {
                            created++;
                            ghostExplodedCreated++;
                            if (autoIsoApplied)
                            {
                                autoIsoAppliedCount++;
                            }

                            if (fallbackRotationApplied)
                            {
                                fallbackRotationCount++;
                            }

                            if (contourAttempted)
                            {
                                if (contourApplied)
                                {
                                    contourColoredCount++;
                                }
                                else
                                {
                                    contourNotChangedCount++;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        if (guideLinesApplied)
                        {
                            guideLinesRequested++;
                            bool pairAvailable = originalCreated && explodedCreated;
                            if (pairAvailable
                                && TryCreateGuideLine(
                                    sheet,
                                    anchorX + plan.OriginalOffsetX,
                                    anchorY + plan.OriginalOffsetY,
                                    anchorZ,
                                    anchorX + plan.OffsetX,
                                    anchorY + plan.OffsetY,
                                    anchorZ))
                            {
                                guideLinesCreated++;
                            }
                            else
                            {
                                guideLinesFailed++;
                            }
                        }

                        continue;
                    }

                    bool standardAutoIsoApplied;
                    bool standardFallbackRotationApplied;
                    bool standardContourAttempted;
                    bool standardContourApplied;
                    string viewName = ExplodedViewPrefix + viewSequence.ToString("D3");
                    viewSequence++;
                    bool standardCreated = TryCreateAndPlaceColoredView(
                        sheet,
                        viewCoordinateSystem,
                        displayCoordinateSystem,
                        plan,
                        sourceScale,
                        viewName,
                        anchorX + plan.OffsetX,
                        anchorY + plan.OffsetY,
                        anchorZ,
                        colorizeEnabled,
                        false,
                        out standardAutoIsoApplied,
                        out standardFallbackRotationApplied,
                        out standardContourAttempted,
                        out standardContourApplied);

                    if (!standardCreated)
                    {
                        failed++;
                        continue;
                    }

                    created++;
                    if (standardAutoIsoApplied)
                    {
                        autoIsoAppliedCount++;
                    }

                    if (standardFallbackRotationApplied)
                    {
                        fallbackRotationCount++;
                    }

                    if (standardContourAttempted)
                    {
                        if (standardContourApplied)
                        {
                            contourColoredCount++;
                        }
                        else
                        {
                            contourNotChangedCount++;
                        }
                    }
                }

                InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");

                StringBuilder report = new StringBuilder();
                report.AppendLine("Vista explodida por planos (somente drawing)");
                report.AppendLine("--------------------------------------------");
                report.AppendLine("Referencia da vista: " + referenceMode);
                report.AppendLine("Colorir: " + (colorizeEnabled ? "ligado" : "desligado"));
                report.AppendLine("Ghost: " + (ghostEnabled ? (ghostApplied ? "ativo" : "solicitado (desativado por fallback)") : "desligado"));
                report.AppendLine("Linha guia: " + (lineGuideEnabled ? (guideLinesApplied ? "ligada" : "solicitada (desativada por fallback)") : "desligada"));
                if (autoDetectEnabled)
                {
                    report.AppendLine("Vistas avaliadas automaticamente: " + autoViewCount);
                    report.AppendLine("Candidatas isometricas: " + autoIsoCandidates);
                }
                else
                {
                    report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                }
                report.AppendLine("Partes candidatas na vista: " + candidateCount + (truncated ? " (limitado para " + maxPartsAllowed + ")" : string.Empty));
                report.AppendLine("Vistas explodidas antigas removidas: " + removedExisting);
                report.AppendLine("Base de posicionamento: " + (modelLayoutUsed ? "modelo (explosao por planos)" : "fallback"));
                report.AppendLine("Fallback usado: " + (fallbackLayoutUsed ? "sim" : "nao"));
                report.AppendLine("Anchor de posicionamento: " + (anchorFromMainPart ? "centro da peca principal" : "origem da vista"));
                report.AppendLine("Centros lidos do modelo: " + centersFromModel);
                report.AppendLine("Movimentos aplicados por plano:");
                report.AppendLine("  xy: " + movedByXY);
                report.AppendLine("  xz: " + movedByXZ);
                report.AppendLine("  zy: " + movedByZY);
                report.AppendLine("Perfil de vista solicitado: " + ExplodedViewAttributeFile);
                report.AppendLine("Vistas com perfil aplicado: " + autoIsoAppliedCount);
                report.AppendLine("Rotacao isometrica de fallback: " + fallbackRotationCount);
                report.AppendLine("Cor de contorno aplicada: " + contourColoredCount);
                report.AppendLine("Sem mudanca de cor: " + contourNotChangedCount);
                if (guideLinesApplied)
                {
                    report.AppendLine("Linhas guia solicitadas: " + guideLinesRequested);
                    report.AppendLine("Linhas guia criadas: " + guideLinesCreated);
                    report.AppendLine("Linhas guia falhas: " + guideLinesFailed);
                }
                if (ghostApplied)
                {
                    report.AppendLine("Ghost pares secundarios: " + ghostPairsRequested);
                    report.AppendLine("Ghost vistas originais criadas: " + ghostOriginalCreated);
                    report.AppendLine("Ghost vistas deslocadas criadas: " + ghostExplodedCreated);
                }
                report.AppendLine("Vistas criadas: " + created);
                report.AppendLine("Falhas: " + failed);
                report.AppendLine();
                report.AppendLine("Origem da vista fonte (sheet): X=" + anchorX.ToString("0.###") + " Y=" + anchorY.ToString("0.###"));
                report.AppendLine("Escala base usada: " + sourceScale.ToString("0.###"));
                report.AppendLine();
                report.AppendLine("Observacao: fluxo 100% drawing-only (sem leitura/escrita no modelo).");

                SetOutput(report.ToString());

                if (created > 0)
                {
                    UpdateStatus("Sucesso: vista por planos criada no desenho.", Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Falha: nenhuma vista foi criada.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao explodir por planos: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private bool TryCreateAndPlaceColoredView(
            object sheet,
            object viewCoordinateSystem,
            object displayCoordinateSystem,
            ExplodedPartPlan plan,
            double sourceScale,
            string viewName,
            double targetX,
            double targetY,
            double targetZ,
            bool applyColoring,
            bool applyGhostStyle,
            out bool autoIsoApplied,
            out bool fallbackRotationApplied,
            out bool contourAttempted,
            out bool contourApplied)
        {
            autoIsoApplied = false;
            fallbackRotationApplied = false;
            contourAttempted = false;
            contourApplied = false;

            object newView = CreateSinglePartView(
                sheet,
                viewCoordinateSystem,
                displayCoordinateSystem,
                plan.Identifier,
                sourceScale,
                viewName,
                out autoIsoApplied);
            if (newView == null)
            {
                return false;
            }

            if (RotateExplodedToIsometric && !autoIsoApplied)
            {
                RotateViewToIsometric(newView);
                fallbackRotationApplied = true;
            }

            bool positioned = ForceViewCenterToTarget(newView, targetX, targetY, targetZ);
            if (!positioned)
            {
                InvokeBoolMethod(newView, "Delete");
                return false;
            }

            if (applyGhostStyle)
            {
                TryApplyGhostStyleToView(newView);
            }

            int contourColor = GetContourColorForPlan(plan);
            contourAttempted = applyColoring && contourColor > 0;
            if (contourAttempted)
            {
                contourApplied = TryApplyContourColorToView(newView, contourColor);
            }

            return true;
        }

        private static bool TryCreateGuideLine(
            object viewBase,
            double startX,
            double startY,
            double startZ,
            double endX,
            double endY,
            double endZ)
        {
            if (viewBase == null)
            {
                return false;
            }

            double dx = endX - startX;
            double dy = endY - startY;
            double dz = endZ - startZ;
            if (((dx * dx) + (dy * dy) + (dz * dz)) < 1e-8)
            {
                return false;
            }

            object startPoint = CreatePoint(startX, startY, startZ);
            object endPoint = CreatePoint(endX, endY, endZ);
            if (startPoint == null || endPoint == null)
            {
                return false;
            }

            Type lineType = ResolveTeklaType(
                "Tekla.Structures.Drawing.Line",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (lineType == null)
            {
                return false;
            }

            Type lineAttributesType = ResolveTeklaType(
                "Tekla.Structures.Drawing.Line+LineAttributes",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);

            object lineAttributes = null;
            if (lineAttributesType != null)
            {
                try
                {
                    lineAttributes = Activator.CreateInstance(lineAttributesType);
                }
                catch
                {
                    lineAttributes = null;
                }
            }

            if (lineAttributes != null)
            {
                ApplyGuideLineStyle(lineAttributes);
            }

            object lineObject = null;
            if (lineAttributes != null)
            {
                try
                {
                    lineObject = Activator.CreateInstance(lineType, new object[] { viewBase, startPoint, endPoint, lineAttributes });
                }
                catch
                {
                    lineObject = null;
                }
            }

            if (lineObject == null)
            {
                try
                {
                    lineObject = Activator.CreateInstance(lineType, new object[] { viewBase, startPoint, endPoint });
                }
                catch
                {
                    lineObject = null;
                }
            }

            if (lineObject == null)
            {
                return false;
            }

            if (lineAttributes != null)
            {
                SetPropertyValue(lineObject, "Attributes", lineAttributes, "Tekla.Structures.Drawing.Line");
            }

            bool? inserted = InvokeBoolMethod(lineObject, "Insert");
            if (!inserted.HasValue || !inserted.Value)
            {
                return false;
            }

            return true;
        }

        private static bool ApplyGuideLineStyle(object lineAttributes)
        {
            if (lineAttributes == null)
            {
                return false;
            }

            bool changed = false;
            object lineTypeAttributes = GetPropertyValue(lineAttributes, "Line");
            if (lineTypeAttributes != null)
            {
                changed |= TrySetLineTypeDashed(lineTypeAttributes);
                changed |= TrySetColorProperty(lineTypeAttributes, "Color", 160);
                changed |= TrySetColorProperty(lineTypeAttributes, "TrueColor", 160);
                SetPropertyValue(lineAttributes, "Line", lineTypeAttributes);
            }

            return changed;
        }

        private static bool TryApplyGhostStyleToView(object view)
        {
            if (view == null)
            {
                return false;
            }

            bool anyChange = false;

            object modelObjects = InvokeParameterlessMethod(view, "GetModelObjects");
            if (modelObjects != null)
            {
                anyChange |= TryApplyGhostStyleToEnumerator(modelObjects);
            }

            if (!anyChange)
            {
                object allObjects = InvokeParameterlessMethod(view, "GetObjects")
                    ?? InvokeParameterlessMethod(view, "GetAllObjects");
                if (allObjects != null)
                {
                    anyChange |= TryApplyGhostStyleToEnumerator(allObjects);
                }
            }

            if (anyChange)
            {
                InvokeBoolMethod(view, "Modify");
            }

            return anyChange;
        }

        private static bool TryApplyGhostStyleToEnumerator(object enumerator)
        {
            if (enumerator == null)
            {
                return false;
            }

            MethodInfo moveNext = enumerator.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = enumerator.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return false;
            }

            bool anyChange = false;
            while (true)
            {
                object moved = moveNext.Invoke(enumerator, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(enumerator, null);
                if (drawingObject == null)
                {
                    continue;
                }

                if (TryApplyGhostStyleToPartObject(drawingObject))
                {
                    anyChange = true;
                    InvokeBoolMethod(drawingObject, "Modify");
                }
            }

            return anyChange;
        }

        private static bool TryApplyGhostStyleToPartObject(object drawingObject)
        {
            if (drawingObject == null)
            {
                return false;
            }

            Type type = drawingObject.GetType();
            string fullName = type != null ? type.FullName : null;
            if (!string.Equals(fullName, "Tekla.Structures.Drawing.Part", StringComparison.Ordinal))
            {
                return false;
            }

            object partAttributes = GetPropertyValue(drawingObject, "Attributes", "Tekla.Structures.Drawing.Part")
                ?? GetPropertyValue(drawingObject, "Attributes");
            if (partAttributes == null)
            {
                return false;
            }

            bool changed = TryApplyGhostStyleToPartAttributes(partAttributes);
            if (!changed)
            {
                return false;
            }

            bool attributesSet = SetPropertyValue(drawingObject, "Attributes", partAttributes, "Tekla.Structures.Drawing.Part")
                || SetPropertyValue(drawingObject, "Attributes", partAttributes);
            return attributesSet;
        }

        private static bool TryApplyGhostStyleToPartAttributes(object partAttributes)
        {
            if (partAttributes == null)
            {
                return false;
            }

            bool changed = false;

            object visibleLines = GetPropertyValue(partAttributes, "VisibleLines");
            if (visibleLines != null)
            {
                changed |= TrySetLineTypeDashed(visibleLines);
                SetPropertyValue(partAttributes, "VisibleLines", visibleLines);
            }

            object hiddenLines = GetPropertyValue(partAttributes, "HiddenLines");
            if (hiddenLines != null)
            {
                changed |= TrySetLineTypeDashed(hiddenLines);
                SetPropertyValue(partAttributes, "HiddenLines", hiddenLines);
            }

            object referenceLine = GetPropertyValue(partAttributes, "ReferenceLine");
            if (referenceLine != null)
            {
                changed |= TrySetLineTypeDashed(referenceLine);
                SetPropertyValue(partAttributes, "ReferenceLine", referenceLine);
            }

            object faceHatch = GetPropertyValue(partAttributes, "FaceHatch");
            if (faceHatch != null)
            {
                changed |= TrySetHatchInvisible(faceHatch);
                SetPropertyValue(partAttributes, "FaceHatch", faceHatch);
            }

            object sectionFaceHatch = GetPropertyValue(partAttributes, "SectionFaceHatch");
            if (sectionFaceHatch != null)
            {
                changed |= TrySetHatchInvisible(sectionFaceHatch);
                SetPropertyValue(partAttributes, "SectionFaceHatch", sectionFaceHatch);
            }

            return changed;
        }

        private static bool TrySetLineTypeDashed(object lineTypeAttributes)
        {
            if (lineTypeAttributes == null || !IsLineTypeAttributesObject(lineTypeAttributes))
            {
                return false;
            }

            Type lineTypeType = lineTypeAttributes.GetType();
            Assembly drawingAssembly = lineTypeType.Assembly;
            if (drawingAssembly == null)
            {
                return false;
            }

            Type lineTypesClass = drawingAssembly.GetType("Tekla.Structures.Drawing.LineTypes", false);
            if (lineTypesClass == null)
            {
                return false;
            }

            FieldInfo dashedField = lineTypesClass.GetField("DashedLine", BindingFlags.Public | BindingFlags.Static);
            if (dashedField == null)
            {
                return false;
            }

            object dashedValue = dashedField.GetValue(null);
            if (dashedValue == null)
            {
                return false;
            }

            return SetPropertyValue(lineTypeAttributes, "Type", dashedValue);
        }

        private static bool TrySetHatchInvisible(object hatchAttributes)
        {
            if (hatchAttributes == null)
            {
                return false;
            }

            bool changed = false;
            changed |= TrySetColorProperty(hatchAttributes, "Color", 152);
            changed |= TrySetColorProperty(hatchAttributes, "BackgroundColor", 152);
            changed |= SetPropertyValue(hatchAttributes, "DrawBackgroundColor", false);
            return changed;
        }

        private static int GetContourColorForPlan(ExplodedPartPlan plan)
        {
            if (plan == null)
            {
                return 0;
            }

            if (plan.IsMainPart)
            {
                return MainPartContourColor;
            }

            if (plan.IsOtherPlane)
            {
                return OtherPlaneContourColor;
            }

            int sideForColor = plan.SideOfXYPlane;
            if (sideForColor == 0)
            {
                sideForColor = plan.SideOfXZPlane;
            }

            if (sideForColor == 0)
            {
                sideForColor = plan.SideOfZYPlane;
            }

            if (sideForColor > 0)
            {
                return SidePositiveContourColor;
            }

            if (sideForColor < 0)
            {
                return SideNegativeContourColor;
            }

            return 0;
        }

        private static bool TryApplyContourColorToView(object view, int colorIndex)
        {
            if (view == null || colorIndex <= 0)
            {
                return false;
            }

            bool anyChange = false;

            object modelObjects = InvokeParameterlessMethod(view, "GetModelObjects");
            if (modelObjects != null)
            {
                anyChange |= TryApplyContourColorToEnumerator(modelObjects, colorIndex);
            }

            if (!anyChange)
            {
                object allObjects = InvokeParameterlessMethod(view, "GetObjects")
                    ?? InvokeParameterlessMethod(view, "GetAllObjects");
                if (allObjects != null)
                {
                    anyChange |= TryApplyContourColorToEnumerator(allObjects, colorIndex);
                }
            }

            if (anyChange)
            {
                bool? modifiedView = InvokeBoolMethod(view, "Modify");
                if (modifiedView.HasValue && !modifiedView.Value)
                {
                    return false;
                }
            }

            return anyChange;
        }

        private static bool TryApplyContourColorToEnumerator(object enumerator, int colorIndex)
        {
            if (enumerator == null || colorIndex <= 0)
            {
                return false;
            }

            MethodInfo moveNext = enumerator.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = enumerator.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return false;
            }

            bool anyChange = false;
            while (true)
            {
                object moved = moveNext.Invoke(enumerator, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(enumerator, null);
                if (drawingObject == null)
                {
                    continue;
                }

                bool objectChanged = false;
                if (TryApplyContourColorToPartObject(drawingObject, colorIndex))
                {
                    objectChanged = true;
                }
                else
                {
                    if (TryApplyContourColorToObject(drawingObject, colorIndex))
                    {
                        objectChanged = true;
                    }

                    object genericAttributes = GetPropertyValue(drawingObject, "Attributes");
                    if (genericAttributes != null && TryApplyContourColorToObject(genericAttributes, colorIndex))
                    {
                        SetPropertyValue(drawingObject, "Attributes", genericAttributes);
                        objectChanged = true;
                    }
                }

                if (objectChanged)
                {
                    anyChange = true;
                    InvokeBoolMethod(drawingObject, "Modify");
                }
            }

            return anyChange;
        }

        private static bool TryApplyContourColorToPartObject(object drawingObject, int colorIndex)
        {
            if (drawingObject == null || colorIndex <= 0)
            {
                return false;
            }

            Type type = drawingObject.GetType();
            string fullName = type != null ? type.FullName : null;
            if (!string.Equals(fullName, "Tekla.Structures.Drawing.Part", StringComparison.Ordinal))
            {
                return false;
            }

            object partAttributes = GetPropertyValue(drawingObject, "Attributes", "Tekla.Structures.Drawing.Part")
                ?? GetPropertyValue(drawingObject, "Attributes");
            if (partAttributes == null)
            {
                return false;
            }

            bool changed = TryApplyContourColorToPartAttributes(partAttributes, colorIndex);
            if (!changed)
            {
                return false;
            }

            bool attributesSet = SetPropertyValue(drawingObject, "Attributes", partAttributes, "Tekla.Structures.Drawing.Part")
                || SetPropertyValue(drawingObject, "Attributes", partAttributes);
            return attributesSet;
        }

        private static bool TryApplyContourColorToPartAttributes(object partAttributes, int colorIndex)
        {
            if (partAttributes == null || colorIndex <= 0)
            {
                return false;
            }

            bool changed = false;

            object visibleLines = GetPropertyValue(partAttributes, "VisibleLines");
            if (visibleLines != null)
            {
                changed |= TrySetLineTypeColor(visibleLines, colorIndex);
                SetPropertyValue(partAttributes, "VisibleLines", visibleLines);
            }

            object hiddenLines = GetPropertyValue(partAttributes, "HiddenLines");
            if (hiddenLines != null)
            {
                changed |= TrySetLineTypeColor(hiddenLines, colorIndex);
                SetPropertyValue(partAttributes, "HiddenLines", hiddenLines);
            }

            object referenceLine = GetPropertyValue(partAttributes, "ReferenceLine");
            if (referenceLine != null)
            {
                changed |= TrySetLineTypeColor(referenceLine, colorIndex);
                SetPropertyValue(partAttributes, "ReferenceLine", referenceLine);
            }

            return changed;
        }

        private static bool TrySetLineTypeColor(object lineTypeAttributes, int colorIndex)
        {
            if (lineTypeAttributes == null || colorIndex <= 0)
            {
                return false;
            }

            if (!IsLineTypeAttributesObject(lineTypeAttributes))
            {
                return false;
            }

            bool changed = TrySetColorProperty(lineTypeAttributes, "Color", colorIndex);

            Type lineType = lineTypeAttributes.GetType();
            Assembly drawingAssembly = lineType != null ? lineType.Assembly : null;
            if (drawingAssembly != null)
            {
                Type drawingColorsType = drawingAssembly.GetType("Tekla.Structures.Drawing.DrawingColors", false);
                Type teklaDrawingColorType = drawingAssembly.GetType("Tekla.Structures.Drawing.TeklaDrawingColor", false);
                if (drawingColorsType != null && drawingColorsType.IsEnum && teklaDrawingColorType != null)
                {
                    try
                    {
                        object drawingColor = Enum.ToObject(drawingColorsType, colorIndex);
                        object trueColor = Activator.CreateInstance(teklaDrawingColorType, new[] { drawingColor });
                        if (SetPropertyValue(lineTypeAttributes, "TrueColor", trueColor))
                        {
                            changed = true;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return changed;
        }

        private static bool TryApplyContourColorToObject(object target, int colorIndex)
        {
            if (target == null)
            {
                return false;
            }

            bool changed = false;

            if (IsLineTypeAttributesObject(target))
            {
                changed |= TrySetColorProperty(target, "Color", colorIndex);
                changed |= TrySetColorProperty(target, "TrueColor", colorIndex);
                return changed;
            }

            changed |= TrySetColorProperty(target, "LineColor", colorIndex);
            changed |= TrySetColorProperty(target, "VisibleLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "VisibleLinesColor", colorIndex);
            changed |= TrySetColorProperty(target, "HiddenLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "HiddenLinesColor", colorIndex);
            changed |= TrySetColorProperty(target, "CenterLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "ContourLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "CutLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "SectionLineColor", colorIndex);
            changed |= TrySetColorProperty(target, "OutlineColor", colorIndex);

            PropertyInfo[] properties = target.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                if (!property.CanRead || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                Type propertyType = property.PropertyType;
                if (propertyType.IsPrimitive || propertyType.IsEnum || propertyType == typeof(string))
                {
                    continue;
                }

                object value;
                try
                {
                    value = property.GetValue(target, null);
                }
                catch
                {
                    continue;
                }

                if (value == null || ReferenceEquals(value, target))
                {
                    continue;
                }

                if (IsLineTypeAttributesObject(value))
                {
                    changed |= TrySetColorProperty(value, "Color", colorIndex);
                    changed |= TrySetColorProperty(value, "TrueColor", colorIndex);
                    continue;
                }

                changed |= TrySetColorProperty(value, "LineColor", colorIndex);
                changed |= TrySetColorProperty(value, "VisibleLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "VisibleLinesColor", colorIndex);
                changed |= TrySetColorProperty(value, "HiddenLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "HiddenLinesColor", colorIndex);
                changed |= TrySetColorProperty(value, "CenterLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "ContourLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "CutLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "SectionLineColor", colorIndex);
                changed |= TrySetColorProperty(value, "OutlineColor", colorIndex);
            }

            return changed;
        }

        private static bool IsLineTypeAttributesObject(object target)
        {
            if (target == null)
            {
                return false;
            }

            Type type = target.GetType();
            string fullName = type != null ? type.FullName : null;
            return string.Equals(fullName, "Tekla.Structures.Drawing.LineTypeAttributes", StringComparison.Ordinal);
        }

        private static bool TrySetColorProperty(object target, string propertyName, int colorIndex)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            PropertyInfo[] properties = target.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            bool changed = false;
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!property.CanWrite || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                try
                {
                    Type propertyType = property.PropertyType;
                    object colorValue = null;

                    if (propertyType.IsEnum)
                    {
                        colorValue = Enum.ToObject(propertyType, colorIndex);
                    }
                    else if (propertyType == typeof(int))
                    {
                        colorValue = colorIndex;
                    }
                    else if (propertyType == typeof(short))
                    {
                        colorValue = (short)colorIndex;
                    }
                    else if (propertyType == typeof(byte))
                    {
                        colorValue = (byte)colorIndex;
                    }
                    else if (propertyType == typeof(long))
                    {
                        colorValue = (long)colorIndex;
                    }
                    else if (propertyType == typeof(double))
                    {
                        colorValue = (double)colorIndex;
                    }
                    else
                    {
                        continue;
                    }

                    property.SetValue(target, colorValue, null);
                    changed = true;
                }
                catch
                {
                }
            }

            return changed;
        }

        private void UpdateStatus(string message, Color color)
        {
            lblStatusTekla.Text = "Status: " + message;
            lblStatusTekla.ForeColor = color;
        }

        private static void CommitActiveDrawingChanges(object activeDrawing)
        {
            if (activeDrawing != null)
            {
                InvokeParameterlessMethod(activeDrawing, "CommitChanges");
            }
        }

        private void SetOutput(string text)
        {
            txtSaida.Text = text ?? string.Empty;
        }

        private static object CreateDrawingHandler()
        {
            Type drawingHandlerType = ResolveTeklaType(
                "Tekla.Structures.Drawing.DrawingHandler",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);

            return drawingHandlerType != null ? Activator.CreateInstance(drawingHandlerType) : null;
        }

        private static object CreateModelInstance()
        {
            Type modelType = ResolveTeklaModelType();
            return modelType != null ? Activator.CreateInstance(modelType) : null;
        }

        private static Type ResolveTeklaModelType()
        {
            return ResolveTeklaType(
                "Tekla.Structures.Model.Model",
                "Tekla.Structures.Model",
                TeklaModelDllCandidates);
        }

        private static Type ResolveTeklaPointType()
        {
            return ResolveTeklaType(
                "Tekla.Structures.Geometry3d.Point",
                "Tekla.Structures",
                TeklaStructuresDllCandidates);
        }

        private static Type ResolveTeklaType(string fullTypeName, string assemblyName, string[] candidatePaths)
        {
            Type resolved = Type.GetType(fullTypeName + ", " + assemblyName);
            if (resolved != null)
            {
                return resolved;
            }

            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (!string.Equals(assembly.GetName().Name, assemblyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                resolved = assembly.GetType(fullTypeName, false);
                if (resolved != null)
                {
                    return resolved;
                }
            }

            foreach (string dllPath in candidatePaths)
            {
                if (!File.Exists(dllPath))
                {
                    continue;
                }

                Assembly loadedAssembly = Assembly.LoadFrom(dllPath);
                resolved = loadedAssembly.GetType(fullTypeName, false);
                if (resolved != null)
                {
                    return resolved;
                }
            }

            return null;
        }
        private static bool? InvokeBoolMethod(object target, string methodName)
        {
            object value = InvokeParameterlessMethod(target, methodName);
            if (value is bool)
            {
                return (bool)value;
            }

            bool parsedValue;
            if (value != null && bool.TryParse(value.ToString(), out parsedValue))
            {
                return parsedValue;
            }

            return null;
        }

        private static object InvokeParameterlessMethod(object target, string methodName)
        {
            return InvokeMethod(target, methodName);
        }

        private static object InvokeMethod(object target, string methodName, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            object[] invokeArgs = args ?? new object[0];
            MethodInfo method = FindCompatibleMethod(target.GetType(), methodName, invokeArgs);
            if (method == null)
            {
                return null;
            }

            return method.Invoke(target, invokeArgs);
        }

        private static MethodInfo FindCompatibleMethod(Type type, string methodName, object[] args)
        {
            MethodInfo[] methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < methods.Length; i++)
            {
                MethodInfo method = methods[i];
                if (!string.Equals(method.Name, methodName, StringComparison.Ordinal))
                {
                    continue;
                }

                ParameterInfo[] parameters = method.GetParameters();
                if (parameters.Length != args.Length)
                {
                    continue;
                }

                bool compatible = true;
                for (int p = 0; p < parameters.Length; p++)
                {
                    object arg = args[p];
                    Type parameterType = parameters[p].ParameterType;

                    if (arg == null)
                    {
                        if (parameterType.IsValueType && Nullable.GetUnderlyingType(parameterType) == null)
                        {
                            compatible = false;
                            break;
                        }

                        continue;
                    }

                    Type argType = arg.GetType();
                    if (!parameterType.IsAssignableFrom(argType))
                    {
                        compatible = false;
                        break;
                    }
                }

                if (compatible)
                {
                    return method;
                }
            }

            return null;
        }

        private static object TryGetSelectedView(object drawingHandler, out int selectedObjectCount)
        {
            selectedObjectCount = 0;

            object selector = InvokeParameterlessMethod(drawingHandler, "GetDrawingObjectSelector");
            if (selector == null)
            {
                return null;
            }

            object selectedEnumerator = InvokeParameterlessMethod(selector, "GetSelected");
            if (selectedEnumerator == null)
            {
                return null;
            }

            Type enumeratorType = selectedEnumerator.GetType();
            MethodInfo moveNextMethod = enumeratorType.GetMethod(
                "MoveNext",
                BindingFlags.Public | BindingFlags.Instance,
                null,
                Type.EmptyTypes,
                null);
            PropertyInfo currentProperty = enumeratorType.GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNextMethod == null || currentProperty == null)
            {
                return null;
            }

            Type viewBaseType = ResolveTeklaType(
                "Tekla.Structures.Drawing.ViewBase",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);

            while (true)
            {
                object moveNextResult = moveNextMethod.Invoke(selectedEnumerator, null);
                if (!(moveNextResult is bool) || !(bool)moveNextResult)
                {
                    break;
                }

                object current = currentProperty.GetValue(selectedEnumerator, null);
                if (current == null)
                {
                    continue;
                }

                selectedObjectCount++;
                object view = ExtractViewFromSelectedObject(current, viewBaseType);
                if (view != null)
                {
                    return view;
                }
            }

            return null;
        }

        private static object TryDetectIsometricReferenceView(object activeDrawing, out int evaluatedViews, out int isometricCandidates)
        {
            evaluatedViews = 0;
            isometricCandidates = 0;

            if (activeDrawing == null)
            {
                return null;
            }

            object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
            if (sheet == null)
            {
                return null;
            }

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects");
            if (allObjects == null)
            {
                return null;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return null;
            }

            object bestIsoView = null;
            double bestIsoScore = double.MinValue;
            double bestIsoArea = double.MinValue;
            object bestFallbackView = null;
            double bestFallbackArea = double.MinValue;

            while (true)
            {
                object moved = moveNext.Invoke(allObjects, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(allObjects, null);
                if (!IsViewObject(drawingObject, null))
                {
                    continue;
                }

                object rawName = GetPropertyValue(drawingObject, "Name");
                string viewName = rawName != null ? rawName.ToString() : string.Empty;
                if (!string.IsNullOrWhiteSpace(viewName)
                    && viewName.StartsWith(ExplodedViewPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                evaluatedViews++;

                double viewWidth = ReadDoubleProperty(drawingObject, "Width", 1.0);
                double viewHeight = ReadDoubleProperty(drawingObject, "Height", 1.0);
                double viewArea = Math.Max(1.0, Math.Abs(viewWidth * viewHeight));
                double isoScore = ComputeIsometricScore(drawingObject);

                const double isoThreshold = 0.45;
                bool isIso = isoScore >= isoThreshold;
                if (isIso)
                {
                    isometricCandidates++;
                    if (isoScore > bestIsoScore
                        || (Math.Abs(isoScore - bestIsoScore) < 1e-6 && viewArea > bestIsoArea))
                    {
                        bestIsoScore = isoScore;
                        bestIsoArea = viewArea;
                        bestIsoView = drawingObject;
                    }
                }

                if (viewArea > bestFallbackArea)
                {
                    bestFallbackArea = viewArea;
                    bestFallbackView = drawingObject;
                }
            }

            return bestIsoView ?? bestFallbackView;
        }

        private static double ComputeIsometricScore(object view)
        {
            double axisXx;
            double axisXy;
            double axisXz;
            double axisYx;
            double axisYy;
            double axisYz;
            if (!TryGetViewDisplayAxes(view, out axisXx, out axisXy, out axisXz, out axisYx, out axisYy, out axisYz))
            {
                return 0.0;
            }

            double axisZx = (axisXy * axisYz) - (axisXz * axisYy);
            double axisZy = (axisXz * axisYx) - (axisXx * axisYz);
            double axisZz = (axisXx * axisYy) - (axisXy * axisYx);
            if (!NormalizeVector3D(ref axisZx, ref axisZy, ref axisZz))
            {
                return 0.0;
            }

            double scoreX = ComputeAxisSpreadScore(axisXx, axisXy, axisXz);
            double scoreY = ComputeAxisSpreadScore(axisYx, axisYy, axisYz);
            double scoreZ = ComputeAxisSpreadScore(axisZx, axisZy, axisZz);
            return scoreX + scoreY + scoreZ;
        }

        private static double ComputeAxisSpreadScore(double x, double y, double z)
        {
            double ax = Math.Abs(x);
            double ay = Math.Abs(y);
            double az = Math.Abs(z);
            double dominant = Math.Max(ax, Math.Max(ay, az));
            return Math.Max(0.0, 1.0 - dominant);
        }

        private static object ExtractViewFromSelectedObject(object selectedObject, Type viewBaseType)
        {
            if (IsViewObject(selectedObject, viewBaseType))
            {
                return selectedObject;
            }

            foreach (string propertyName in new[] { "View", "ParentView" })
            {
                PropertyInfo property = selectedObject.GetType().GetProperty(
                    propertyName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (property == null || !property.CanRead || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                object candidate = property.GetValue(selectedObject, null);
                if (IsViewObject(candidate, viewBaseType))
                {
                    return candidate;
                }
            }

            foreach (string methodName in new[] { "GetView", "GetParentView" })
            {
                MethodInfo method = selectedObject.GetType().GetMethod(
                    methodName,
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    Type.EmptyTypes,
                    null);
                if (method == null)
                {
                    continue;
                }

                object candidate = method.Invoke(selectedObject, null);
                if (IsViewObject(candidate, viewBaseType))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static bool IsViewObject(object candidate, Type viewBaseType)
        {
            if (candidate == null)
            {
                return false;
            }

            if (viewBaseType != null && viewBaseType.IsAssignableFrom(candidate.GetType()))
            {
                return true;
            }

            return candidate.GetType().Name.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static List<ExplodedPartPlan> CollectPartPlansFromView(object sourceView)
        {
            List<ExplodedPartPlan> plans = new List<ExplodedPartPlan>();
            HashSet<string> partIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<ExplodedPartPlan> fallbackObjects = new List<ExplodedPartPlan>();
            HashSet<string> fallbackIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            object modelObjects = InvokeParameterlessMethod(sourceView, "GetModelObjects");
            if (modelObjects == null)
            {
                return plans;
            }

            MethodInfo moveNext = modelObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = modelObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return plans;
            }

            while (true)
            {
                object moved = moveNext.Invoke(modelObjects, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingModelObject = current.GetValue(modelObjects, null);
                if (drawingModelObject == null)
                {
                    continue;
                }

                object identifier = GetPropertyValue(drawingModelObject, "ModelIdentifier");
                if (identifier == null)
                {
                    continue;
                }

                string key = GetIdentifierKey(identifier);
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                ExplodedPartPlan plan = new ExplodedPartPlan();
                plan.Identifier = identifier;
                plan.IdentifierKey = key;
                plan.DrawingObject = drawingModelObject;

                double center2DX;
                double center2DY;
                double size2D;
                if (TryGetDrawingObjectCenter2D(drawingModelObject, out center2DX, out center2DY, out size2D))
                {
                    plan.HasCenter2D = true;
                    plan.Center2DX = center2DX;
                    plan.Center2DY = center2DY;
                    plan.ApproxSize2D = size2D;
                }

                if (!fallbackIds.Contains(key))
                {
                    fallbackObjects.Add(plan);
                    fallbackIds.Add(key);
                }

                if (!IsPartLikeTypeName(drawingModelObject.GetType().Name))
                {
                    continue;
                }

                if (!partIds.Contains(key))
                {
                    plans.Add(plan);
                    partIds.Add(key);
                }
            }

            if (plans.Count > 0)
            {
                return plans;
            }

            return fallbackObjects;
        }

        private static bool IsPartLikeTypeName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
            {
                return false;
            }

            if (typeName.IndexOf("Part", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (typeName.IndexOf("Beam", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (typeName.IndexOf("Plate", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static string GetIdentifierKey(object identifier)
        {
            if (identifier == null)
            {
                return string.Empty;
            }

            object idNumber = GetPropertyValue(identifier, "ID") ?? GetPropertyValue(identifier, "Id");
            if (idNumber != null)
            {
                return idNumber.ToString();
            }

            return identifier.ToString();
        }

        private static bool TryGetMainPartCenter2D(List<ExplodedPartPlan> plans, out double centerX, out double centerY)
        {
            centerX = 0.0;
            centerY = 0.0;

            if (plans == null || plans.Count == 0)
            {
                return false;
            }

            ExplodedPartPlan mainPlan = null;
            double bestScore = double.MinValue;
            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                if (!plan.HasCenter2D)
                {
                    continue;
                }

                if (plan.ApproxSize2D > bestScore)
                {
                    bestScore = plan.ApproxSize2D;
                    mainPlan = plan;
                }
            }

            if (mainPlan == null)
            {
                return false;
            }

            centerX = mainPlan.Center2DX;
            centerY = mainPlan.Center2DY;
            return true;
        }

        private static void MarkMainPartByLargest2D(List<ExplodedPartPlan> plans)
        {
            if (plans == null || plans.Count == 0)
            {
                return;
            }

            ExplodedPartPlan mainPlan = null;
            double bestScore = double.MinValue;

            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                plan.IsMainPart = false;
                plan.IsOtherPlane = false;
                plan.SideOfXYPlane = 0;
                plan.SideOfXZPlane = 0;
                plan.SideOfZYPlane = 0;
                plan.ColorPlane = 0;

                if (!plan.HasCenter2D)
                {
                    continue;
                }

                if (plan.ApproxSize2D > bestScore)
                {
                    bestScore = plan.ApproxSize2D;
                    mainPlan = plan;
                }
            }

            if (mainPlan == null)
            {
                mainPlan = plans[0];
            }

            mainPlan.IsMainPart = true;
        }

        private static bool ForceViewCenterToTarget(object view, double targetCenterX, double targetCenterY, double targetCenterZ)
        {
            if (view == null)
            {
                return false;
            }

            const double tolerance = 0.05;
            const int maxCorrectionPasses = 3;
            double currentOriginZ = targetCenterZ;

            for (int pass = 0; pass < maxCorrectionPasses; pass++)
            {
                object currentOrigin = GetPropertyValue(view, "Origin") ?? GetPropertyValue(view, "InsertionPoint");
                double currentOriginX;
                double currentOriginY;
                if (!TryGetXYZ(currentOrigin, out currentOriginX, out currentOriginY, out currentOriginZ))
                {
                    currentOriginX = targetCenterX;
                    currentOriginY = targetCenterY;
                    currentOriginZ = targetCenterZ;
                }

                double deltaX = 0.0;
                double deltaY = 0.0;

                if (pass == 0)
                {
                    deltaX = targetCenterX - currentOriginX;
                    deltaY = targetCenterY - currentOriginY;
                }
                else
                {
                    double measuredCenterX;
                    double measuredCenterY;
                    double measuredSize;
                    if (!TryGetDrawingObjectCenter2D(view, out measuredCenterX, out measuredCenterY, out measuredSize))
                    {
                        return pass > 0;
                    }

                    deltaX = targetCenterX - measuredCenterX;
                    deltaY = targetCenterY - measuredCenterY;

                    if (Math.Abs(deltaX) <= tolerance && Math.Abs(deltaY) <= tolerance)
                    {
                        return true;
                    }
                }

                object movedOrigin = CreatePoint(currentOriginX + deltaX, currentOriginY + deltaY, currentOriginZ);
                if (movedOrigin == null)
                {
                    return false;
                }

                bool originSet = SetPropertyValue(view, "Origin", movedOrigin)
                    || SetPropertyValue(view, "InsertionPoint", movedOrigin);
                if (!originSet)
                {
                    return false;
                }

                bool? modified = InvokeBoolMethod(view, "Modify");
                if (modified.HasValue && !modified.Value)
                {
                    return false;
                }
            }

            double finalCenterX;
            double finalCenterY;
            double finalSize;
            if (!TryGetDrawingObjectCenter2D(view, out finalCenterX, out finalCenterY, out finalSize))
            {
                return true;
            }

            return Math.Abs(targetCenterX - finalCenterX) <= tolerance
                && Math.Abs(targetCenterY - finalCenterY) <= tolerance;
        }

        private static bool TryGetDrawingObjectCenter2D(object drawingObject, out double centerX, out double centerY, out double sizeScore)
        {
            centerX = 0.0;
            centerY = 0.0;
            sizeScore = 0.0;

            if (drawingObject == null)
            {
                return false;
            }

            double minX;
            double minY;
            double maxX;
            double maxY;

            if (TryGetDrawingObjectBounds2D(drawingObject, out minX, out minY, out maxX, out maxY))
            {
                centerX = (minX + maxX) * 0.5;
                centerY = (minY + maxY) * 0.5;
                sizeScore = Math.Max(1.0, Math.Max(Math.Abs(maxX - minX), Math.Abs(maxY - minY)));
                return true;
            }

            object origin = GetPropertyValue(drawingObject, "Origin") ?? GetPropertyValue(drawingObject, "InsertionPoint");
            double ox;
            double oy;
            double oz;
            if (TryGetXYZ(origin, out ox, out oy, out oz))
            {
                centerX = ox;
                centerY = oy;
                sizeScore = 1.0;
                return true;
            }

            return false;
        }

        private static bool TryGetDrawingObjectBounds2D(object drawingObject, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (drawingObject == null)
            {
                return false;
            }

            object bounding = InvokeParameterlessMethod(drawingObject, "GetAxisAlignedBoundingBox")
                ?? GetPropertyValue(drawingObject, "BoundingBox")
                ?? GetPropertyValue(drawingObject, "AxisAlignedBoundingBox");

            if (bounding == null)
            {
                bounding = drawingObject;
            }

            object minimum = GetPropertyValue(bounding, "MinimumPoint")
                ?? GetPropertyValue(bounding, "MinPoint")
                ?? GetPropertyValue(bounding, "Minimum");
            object maximum = GetPropertyValue(bounding, "MaximumPoint")
                ?? GetPropertyValue(bounding, "MaxPoint")
                ?? GetPropertyValue(bounding, "Maximum");

            double minZ;
            double maxZ;
            if (TryGetXYZ(minimum, out minX, out minY, out minZ)
                && TryGetXYZ(maximum, out maxX, out maxY, out maxZ))
            {
                return true;
            }

            return false;
        }

        private static int DeleteExistingExplodedViews(object sheet, string prefix)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(prefix))
            {
                return 0;
            }

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects");
            if (allObjects == null)
            {
                return 0;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return 0;
            }

            List<object> viewsToDelete = new List<object>();
            while (true)
            {
                object moved = moveNext.Invoke(allObjects, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(allObjects, null);
                if (!IsViewObject(drawingObject, null))
                {
                    continue;
                }

                object rawName = GetPropertyValue(drawingObject, "Name");
                string viewName = rawName != null ? rawName.ToString() : string.Empty;
                if (string.IsNullOrWhiteSpace(viewName))
                {
                    continue;
                }

                if (!viewName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                viewsToDelete.Add(drawingObject);
            }

            int removed = 0;
            for (int i = 0; i < viewsToDelete.Count; i++)
            {
                bool? deleted = InvokeBoolMethod(viewsToDelete[i], "Delete");
                if (deleted.HasValue && deleted.Value)
                {
                    removed++;
                }
            }

            return removed;
        }

        private static bool ComputeExplodedOffsets(
            object sourceView,
            List<ExplodedPartPlan> plans,
            double sourceScale,
            object model,
            out int centersFromModel)
        {
            centersFromModel = 0;
            if (plans == null || plans.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < plans.Count; i++)
            {
                plans[i].OriginalOffsetX = 0.0;
                plans[i].OriginalOffsetY = 0.0;
                plans[i].OffsetX = 0.0;
                plans[i].OffsetY = 0.0;
                plans[i].HasComputedOffset = false;
            }

            if (TryComputeModelRelativeOffsets(sourceView, plans, sourceScale, model, out centersFromModel))
            {
                return true;
            }

            if (TryComputeOriginalLayoutOffsets(plans))
            {
                return false;
            }

            ApplyRadialOffsets(sourceView, plans, false);
            return false;
        }

        private static bool ComputeExplodedOffsetsByPlanes(
            object sourceView,
            List<ExplodedPartPlan> plans,
            double sourceScale,
            object model,
            bool usePlaneXY,
            bool usePlaneXZ,
            bool usePlaneZY,
            out int centersFromModel,
            out int movedByXY,
            out int movedByXZ,
            out int movedByZY)
        {
            centersFromModel = 0;
            movedByXY = 0;
            movedByXZ = 0;
            movedByZY = 0;

            if (plans == null || plans.Count == 0)
            {
                return false;
            }

            if (!usePlaneXY && !usePlaneXZ && !usePlaneZY)
            {
                return false;
            }

            for (int i = 0; i < plans.Count; i++)
            {
                plans[i].OriginalOffsetX = 0.0;
                plans[i].OriginalOffsetY = 0.0;
                plans[i].OffsetX = 0.0;
                plans[i].OffsetY = 0.0;
                plans[i].HasComputedOffset = false;
            }

            return TryComputeModelRelativeOffsetsByPlanes(
                sourceView,
                plans,
                sourceScale,
                model,
                usePlaneXY,
                usePlaneXZ,
                usePlaneZY,
                out centersFromModel,
                out movedByXY,
                out movedByXZ,
                out movedByZY);
        }

        private static bool TryComputeModelRelativeOffsetsByPlanes(
            object sourceView,
            List<ExplodedPartPlan> plans,
            double sourceScale,
            object model,
            bool usePlaneXY,
            bool usePlaneXZ,
            bool usePlaneZY,
            out int centersFromModel,
            out int movedByXY,
            out int movedByXZ,
            out int movedByZY)
        {
            centersFromModel = 0;
            movedByXY = 0;
            movedByXZ = 0;
            movedByZY = 0;

            if (model == null)
            {
                return false;
            }

            bool? modelConnected = InvokeBoolMethod(model, "GetConnectionStatus");
            if (!modelConnected.HasValue || !modelConnected.Value)
            {
                return false;
            }

            double axisXx;
            double axisXy;
            double axisXz;
            double axisYx;
            double axisYy;
            double axisYz;
            if (!TryGetViewDisplayAxes(sourceView, out axisXx, out axisXy, out axisXz, out axisYx, out axisYy, out axisYz))
            {
                return false;
            }

            List<ExplodedPartPlan> modelPlans = new List<ExplodedPartPlan>();
            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                double centerX;
                double centerY;
                double centerZ;
                double sizeScore;
                bool hasBounds;
                double minX;
                double minY;
                double minZ;
                double maxX;
                double maxY;
                double maxZ;
                if (!TryGetModelObjectCenterAndSize(
                        model,
                        plan.Identifier,
                        out centerX,
                        out centerY,
                        out centerZ,
                        out sizeScore,
                        out hasBounds,
                        out minX,
                        out minY,
                        out minZ,
                        out maxX,
                        out maxY,
                        out maxZ))
                {
                    continue;
                }

                plan.HasModelCenter = true;
                plan.ModelCenterX = centerX;
                plan.ModelCenterY = centerY;
                plan.ModelCenterZ = centerZ;
                plan.ModelSizeScore = sizeScore;
                plan.HasModelBounds = hasBounds;
                plan.ModelMinX = minX;
                plan.ModelMinY = minY;
                plan.ModelMinZ = minZ;
                plan.ModelMaxX = maxX;
                plan.ModelMaxY = maxY;
                plan.ModelMaxZ = maxZ;
                modelPlans.Add(plan);
            }

            centersFromModel = modelPlans.Count;
            if (modelPlans.Count < 2)
            {
                return false;
            }

            ExplodedPartPlan mainPlan = null;
            double bestMainScore = double.MinValue;
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan candidate = modelPlans[i];
                if (candidate.ModelSizeScore > bestMainScore)
                {
                    bestMainScore = candidate.ModelSizeScore;
                    mainPlan = candidate;
                }
            }

            if (mainPlan == null)
            {
                return false;
            }

            for (int i = 0; i < modelPlans.Count; i++)
            {
                modelPlans[i].IsMainPart = ReferenceEquals(modelPlans[i], mainPlan);
            }

            double localAxisXx;
            double localAxisXy;
            double localAxisXz;
            double localAxisYx;
            double localAxisYy;
            double localAxisYz;
            double localAxisZx;
            double localAxisZy;
            double localAxisZz;
            if (!TryGetMainLocalAxes(
                    model,
                    mainPlan.Identifier,
                    out localAxisXx,
                    out localAxisXy,
                    out localAxisXz,
                    out localAxisYx,
                    out localAxisYy,
                    out localAxisYz,
                    out localAxisZx,
                    out localAxisZy,
                    out localAxisZz))
            {
                return false;
            }

            double axisXSheetX = Dot3D(localAxisXx, localAxisXy, localAxisXz, axisXx, axisXy, axisXz);
            double axisXSheetY = Dot3D(localAxisXx, localAxisXy, localAxisXz, axisYx, axisYy, axisYz);
            double axisYSheetX = Dot3D(localAxisYx, localAxisYy, localAxisYz, axisXx, axisXy, axisXz);
            double axisYSheetY = Dot3D(localAxisYx, localAxisYy, localAxisYz, axisYx, axisYy, axisYz);
            double axisZSheetX = Dot3D(localAxisZx, localAxisZy, localAxisZz, axisXx, axisXy, axisXz);
            double axisZSheetY = Dot3D(localAxisZx, localAxisZy, localAxisZz, axisYx, axisYy, axisYz);

            if (!NormalizeVector2D(ref axisXSheetX, ref axisXSheetY)
                || !NormalizeVector2D(ref axisYSheetX, ref axisYSheetY)
                || !NormalizeVector2D(ref axisZSheetX, ref axisZSheetY))
            {
                return false;
            }

            double maxAbsLocalX = 1.0;
            double maxAbsLocalY = 1.0;
            double maxAbsLocalZ = 1.0;
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                double dx = plan.ModelCenterX - mainPlan.ModelCenterX;
                double dy = plan.ModelCenterY - mainPlan.ModelCenterY;
                double dz = plan.ModelCenterZ - mainPlan.ModelCenterZ;

                plan.LocalXFromMain = Dot3D(dx, dy, dz, localAxisXx, localAxisXy, localAxisXz);
                plan.LocalYFromMain = Dot3D(dx, dy, dz, localAxisYx, localAxisYy, localAxisYz);
                plan.LocalZFromMain = Dot3D(dx, dy, dz, localAxisZx, localAxisZy, localAxisZz);

                maxAbsLocalX = Math.Max(maxAbsLocalX, Math.Abs(plan.LocalXFromMain));
                maxAbsLocalY = Math.Max(maxAbsLocalY, Math.Abs(plan.LocalYFromMain));
                maxAbsLocalZ = Math.Max(maxAbsLocalZ, Math.Abs(plan.LocalZFromMain));
            }

            double scaleDivisor = sourceScale > 0.0 ? sourceScale : 1.0;
            double viewWidth = ReadDoubleProperty(sourceView, "Width", 200.0);
            double viewHeight = ReadDoubleProperty(sourceView, "Height", 150.0);
            double minViewDimension = Math.Max(30.0, Math.Min(viewWidth, viewHeight));
            double baseExplodeSheet = Math.Max(16.0, minViewDimension * 0.18);
            double gainExplodeSheet = Math.Max(8.0, minViewDimension * 0.14);
            double axisTolerance = ComputeAxisSideTolerance(mainPlan.ModelSizeScore);

            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                double baseOffsetX;
                double baseOffsetY;
                if (plan.HasCenter2D && mainPlan.HasCenter2D)
                {
                    baseOffsetX = plan.Center2DX - mainPlan.Center2DX;
                    baseOffsetY = plan.Center2DY - mainPlan.Center2DY;
                }
                else
                {
                    double worldOffsetX = plan.ModelCenterX - mainPlan.ModelCenterX;
                    double worldOffsetY = plan.ModelCenterY - mainPlan.ModelCenterY;
                    double worldOffsetZ = plan.ModelCenterZ - mainPlan.ModelCenterZ;
                    baseOffsetX = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisXx, axisXy, axisXz) / scaleDivisor;
                    baseOffsetY = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisYx, axisYy, axisYz) / scaleDivisor;
                }

                double explodeOffsetX = 0.0;
                double explodeOffsetY = 0.0;

                if (!ReferenceEquals(plan, mainPlan))
                {
                    int sideOnXY = ClassifyPartSideOnAxis(
                        plan,
                        mainPlan,
                        localAxisZx,
                        localAxisZy,
                        localAxisZz,
                        axisTolerance);
                    plan.SideOfXYPlane = sideOnXY;

                    int sideOnXZ = ClassifyPartSideOnAxis(
                        plan,
                        mainPlan,
                        localAxisYx,
                        localAxisYy,
                        localAxisYz,
                        axisTolerance);
                    plan.SideOfXZPlane = sideOnXZ;

                    int sideOnZY = ClassifyPartSideOnAxis(
                        plan,
                        mainPlan,
                        localAxisXx,
                        localAxisXy,
                        localAxisXz,
                        axisTolerance);
                    plan.SideOfZYPlane = sideOnZY;

                    plan.ColorPlane = ResolveDominantColorPlane(
                        sideOnXY,
                        sideOnXZ,
                        sideOnZY,
                        Math.Abs(plan.LocalZFromMain) / maxAbsLocalZ,
                        Math.Abs(plan.LocalYFromMain) / maxAbsLocalY,
                        Math.Abs(plan.LocalXFromMain) / maxAbsLocalX);

                    double normalizedXFromMain = maxAbsLocalX > 1e-9
                        ? Math.Abs(plan.LocalXFromMain) / maxAbsLocalX
                        : 0.0;
                    plan.IsOtherPlane =
                        sideOnZY != 0
                        && sideOnXY == 0
                        && sideOnXZ == 0
                        && normalizedXFromMain >= 0.85;

                    if (plan.IsOtherPlane)
                    {
                        double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalXFromMain) / maxAbsLocalX);
                        double amount = (baseExplodeSheet * 1.35) + ((gainExplodeSheet * 1.15) * normalizedDepth);
                        explodeOffsetX += amount * sideOnZY * axisXSheetX;
                        explodeOffsetY += amount * sideOnZY * axisXSheetY;
                        movedByZY++;
                    }

                    if (usePlaneXY)
                    {
                        if (sideOnXY != 0)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalZFromMain) / maxAbsLocalZ);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            explodeOffsetX += amount * sideOnXY * axisZSheetX;
                            explodeOffsetY += amount * sideOnXY * axisZSheetY;
                            movedByXY++;
                        }
                    }

                    if (usePlaneXZ)
                    {
                        if (sideOnXZ != 0)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalYFromMain) / maxAbsLocalY);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            explodeOffsetX += amount * sideOnXZ * axisYSheetX;
                            explodeOffsetY += amount * sideOnXZ * axisYSheetY;
                            movedByXZ++;
                        }
                    }

                    if (usePlaneZY)
                    {
                        if (sideOnZY != 0 && !plan.IsOtherPlane)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalXFromMain) / maxAbsLocalX);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            explodeOffsetX += amount * sideOnZY * axisXSheetX;
                            explodeOffsetY += amount * sideOnZY * axisXSheetY;
                            movedByZY++;
                        }
                    }
                }
                else
                {
                    plan.SideOfXYPlane = 0;
                    plan.SideOfXZPlane = 0;
                    plan.SideOfZYPlane = 0;
                    plan.ColorPlane = 0;
                    plan.IsOtherPlane = false;
                }

                plan.OriginalOffsetX = baseOffsetX;
                plan.OriginalOffsetY = baseOffsetY;
                plan.OffsetX = baseOffsetX + explodeOffsetX;
                plan.OffsetY = baseOffsetY + explodeOffsetY;
                plan.HasComputedOffset = true;
            }

            return true;
        }

        private static int ClassifyPartSideOnAxis(
            ExplodedPartPlan plan,
            ExplodedPartPlan mainPlan,
            double axisX,
            double axisY,
            double axisZ,
            double tolerance)
        {
            double minAxis;
            double maxAxis;
            if (!TryGetLocalAxisRange(plan, mainPlan, axisX, axisY, axisZ, out minAxis, out maxAxis))
            {
                return 0;
            }

            if (minAxis > tolerance)
            {
                return 1;
            }

            if (maxAxis < -tolerance)
            {
                return -1;
            }

            double centerAxis = (minAxis + maxAxis) * 0.5;
            double centerTolerance = Math.Max(0.25, tolerance * 0.35);
            if (centerAxis > centerTolerance)
            {
                return 1;
            }

            if (centerAxis < -centerTolerance)
            {
                return -1;
            }

            return 0;
        }

        private static double ComputeAxisSideTolerance(double mainSizeScore)
        {
            double scaledTolerance = Math.Abs(mainSizeScore) * 0.001;
            if (double.IsNaN(scaledTolerance) || double.IsInfinity(scaledTolerance))
            {
                scaledTolerance = 1.0;
            }

            return Math.Max(0.5, Math.Min(5.0, scaledTolerance));
        }

        private static int ResolveDominantColorPlane(
            int sideOnXY,
            int sideOnXZ,
            int sideOnZY,
            double normalizedZ,
            double normalizedY,
            double normalizedX)
        {
            double bestScore = double.MinValue;
            int bestPlane = 0;

            if (sideOnXY != 0)
            {
                bestScore = normalizedZ;
                bestPlane = 1;
            }

            if (sideOnXZ != 0 && normalizedY > bestScore)
            {
                bestScore = normalizedY;
                bestPlane = 2;
            }

            if (sideOnZY != 0 && normalizedX > bestScore)
            {
                bestPlane = 3;
            }

            if (bestPlane == 0)
            {
                double fallbackScore = normalizedZ;
                bestPlane = 1;
                if (normalizedY > fallbackScore)
                {
                    fallbackScore = normalizedY;
                    bestPlane = 2;
                }

                if (normalizedX > fallbackScore)
                {
                    bestPlane = 3;
                }
            }

            return bestPlane;
        }

        private static bool TryComputeModelRelativeOffsets(
            object sourceView,
            List<ExplodedPartPlan> plans,
            double sourceScale,
            object model,
            out int centersFromModel)
        {
            centersFromModel = 0;
            if (model == null)
            {
                return false;
            }

            bool? modelConnected = InvokeBoolMethod(model, "GetConnectionStatus");
            if (!modelConnected.HasValue || !modelConnected.Value)
            {
                return false;
            }

            double axisXx;
            double axisXy;
            double axisXz;
            double axisYx;
            double axisYy;
            double axisYz;
            if (!TryGetViewDisplayAxes(sourceView, out axisXx, out axisXy, out axisXz, out axisYx, out axisYy, out axisYz))
            {
                return false;
            }

            List<ExplodedPartPlan> modelPlans = new List<ExplodedPartPlan>();
            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                double centerX;
                double centerY;
                double centerZ;
                double sizeScore;
                bool hasBounds;
                double minX;
                double minY;
                double minZ;
                double maxX;
                double maxY;
                double maxZ;
                if (!TryGetModelObjectCenterAndSize(
                        model,
                        plan.Identifier,
                        out centerX,
                        out centerY,
                        out centerZ,
                        out sizeScore,
                        out hasBounds,
                        out minX,
                        out minY,
                        out minZ,
                        out maxX,
                        out maxY,
                        out maxZ))
                {
                    continue;
                }

                plan.HasModelCenter = true;
                plan.ModelCenterX = centerX;
                plan.ModelCenterY = centerY;
                plan.ModelCenterZ = centerZ;
                plan.ModelSizeScore = sizeScore;
                plan.HasModelBounds = hasBounds;
                plan.ModelMinX = minX;
                plan.ModelMinY = minY;
                plan.ModelMinZ = minZ;
                plan.ModelMaxX = maxX;
                plan.ModelMaxY = maxY;
                plan.ModelMaxZ = maxZ;
                modelPlans.Add(plan);
            }

            centersFromModel = modelPlans.Count;
            if (modelPlans.Count < 2)
            {
                return false;
            }

            ExplodedPartPlan mainPlan = null;
            double bestMainScore = double.MinValue;
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan candidate = modelPlans[i];
                if (candidate.ModelSizeScore > bestMainScore)
                {
                    bestMainScore = candidate.ModelSizeScore;
                    mainPlan = candidate;
                }
            }

            if (mainPlan == null)
            {
                return false;
            }

            for (int i = 0; i < modelPlans.Count; i++)
            {
                modelPlans[i].IsMainPart = ReferenceEquals(modelPlans[i], mainPlan);
                modelPlans[i].IsOtherPlane = false;
                if (ReferenceEquals(modelPlans[i], mainPlan))
                {
                    modelPlans[i].ColorPlane = 0;
                    modelPlans[i].SideOfXYPlane = 0;
                    modelPlans[i].SideOfXZPlane = 0;
                    modelPlans[i].SideOfZYPlane = 0;
                }
            }

            double localAxisXx;
            double localAxisXy;
            double localAxisXz;
            double localAxisYx;
            double localAxisYy;
            double localAxisYz;
            double localAxisZx;
            double localAxisZy;
            double localAxisZz;
            if (!TryGetMainLocalAxes(
                    model,
                    mainPlan.Identifier,
                    out localAxisXx,
                    out localAxisXy,
                    out localAxisXz,
                    out localAxisYx,
                    out localAxisYy,
                    out localAxisYz,
                    out localAxisZx,
                    out localAxisZy,
                    out localAxisZz))
            {
                return false;
            }

            double zTolerance = ComputeAxisSideTolerance(mainPlan.ModelSizeScore);
            double maxAbsLocalZ = 1.0;
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                double dx = plan.ModelCenterX - mainPlan.ModelCenterX;
                double dy = plan.ModelCenterY - mainPlan.ModelCenterY;
                double dz = plan.ModelCenterZ - mainPlan.ModelCenterZ;

                plan.LocalXFromMain = Dot3D(dx, dy, dz, localAxisXx, localAxisXy, localAxisXz);
                plan.LocalYFromMain = Dot3D(dx, dy, dz, localAxisYx, localAxisYy, localAxisYz);
                plan.LocalZFromMain = Dot3D(dx, dy, dz, localAxisZx, localAxisZy, localAxisZz);
                maxAbsLocalZ = Math.Max(maxAbsLocalZ, Math.Abs(plan.LocalZFromMain));

                double localZMin;
                double localZMax;
                bool hasLocalZRange = TryGetLocalZRange(
                    plan,
                    mainPlan,
                    localAxisZx,
                    localAxisZy,
                    localAxisZz,
                    out localZMin,
                    out localZMax);

                if (hasLocalZRange && localZMin > zTolerance)
                {
                    plan.SideOfXYPlane = 1;
                }
                else if (hasLocalZRange && localZMax < -zTolerance)
                {
                    plan.SideOfXYPlane = -1;
                }
                else
                {
                    plan.SideOfXYPlane = 0;
                }

                if (!ReferenceEquals(plan, mainPlan))
                {
                    plan.ColorPlane = plan.SideOfXYPlane != 0 ? 1 : 0;
                }
            }

            double scaleDivisor = sourceScale > 0.0 ? sourceScale : 1.0;
            double viewWidth = ReadDoubleProperty(sourceView, "Width", 200.0);
            double viewHeight = ReadDoubleProperty(sourceView, "Height", 150.0);
            double minViewDimension = Math.Max(30.0, Math.Min(viewWidth, viewHeight));
            double baseExplodeSheet = Math.Max(18.0, minViewDimension * 0.22);
            double gainExplodeSheet = Math.Max(10.0, minViewDimension * 0.18);

            double beamSheetX = Dot3D(localAxisXx, localAxisXy, localAxisXz, axisXx, axisXy, axisXz);
            double beamSheetY = Dot3D(localAxisXx, localAxisXy, localAxisXz, axisYx, axisYy, axisYz);
            if (!NormalizeVector2D(ref beamSheetX, ref beamSheetY))
            {
                return false;
            }

            double perpSheetX = -beamSheetY;
            double perpSheetY = beamSheetX;
            if (!NormalizeVector2D(ref perpSheetX, ref perpSheetY))
            {
                return false;
            }

            double localZSheetX = Dot3D(localAxisZx, localAxisZy, localAxisZz, axisXx, axisXy, axisXz);
            double localZSheetY = Dot3D(localAxisZx, localAxisZy, localAxisZz, axisYx, axisYy, axisYz);
            if (((perpSheetX * localZSheetX) + (perpSheetY * localZSheetY)) < 0.0)
            {
                perpSheetX = -perpSheetX;
                perpSheetY = -perpSheetY;
            }

            double maxAbsBasePerp = 1.0;
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                double baseOffsetX;
                double baseOffsetY;
                if (plan.HasCenter2D && mainPlan.HasCenter2D)
                {
                    baseOffsetX = plan.Center2DX - mainPlan.Center2DX;
                    baseOffsetY = plan.Center2DY - mainPlan.Center2DY;
                }
                else
                {
                    double worldOffsetX = plan.ModelCenterX - mainPlan.ModelCenterX;
                    double worldOffsetY = plan.ModelCenterY - mainPlan.ModelCenterY;
                    double worldOffsetZ = plan.ModelCenterZ - mainPlan.ModelCenterZ;
                    baseOffsetX = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisXx, axisXy, axisXz) / scaleDivisor;
                    baseOffsetY = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisYx, axisYy, axisYz) / scaleDivisor;
                }

                double basePerp = (baseOffsetX * perpSheetX) + (baseOffsetY * perpSheetY);
                maxAbsBasePerp = Math.Max(maxAbsBasePerp, Math.Abs(basePerp));
            }

            double perpTolerance = Math.Max(2.0, maxAbsBasePerp * 0.03);
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                double baseOffsetX;
                double baseOffsetY;
                if (plan.HasCenter2D && mainPlan.HasCenter2D)
                {
                    baseOffsetX = plan.Center2DX - mainPlan.Center2DX;
                    baseOffsetY = plan.Center2DY - mainPlan.Center2DY;
                }
                else
                {
                    double worldOffsetX = plan.ModelCenterX - mainPlan.ModelCenterX;
                    double worldOffsetY = plan.ModelCenterY - mainPlan.ModelCenterY;
                    double worldOffsetZ = plan.ModelCenterZ - mainPlan.ModelCenterZ;
                    baseOffsetX = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisXx, axisXy, axisXz) / scaleDivisor;
                    baseOffsetY = Dot3D(worldOffsetX, worldOffsetY, worldOffsetZ, axisYx, axisYy, axisYz) / scaleDivisor;
                }

                double explodeOffsetX = 0.0;
                double explodeOffsetY = 0.0;

                if (!ReferenceEquals(plan, mainPlan))
                {
                    double basePerp = (baseOffsetX * perpSheetX) + (baseOffsetY * perpSheetY);
                    int sideFromDrawing = 0;
                    if (basePerp > perpTolerance)
                    {
                        sideFromDrawing = 1;
                    }
                    else if (basePerp < -perpTolerance)
                    {
                        sideFromDrawing = -1;
                    }

                    int sideToUse = plan.SideOfXYPlane != 0 ? plan.SideOfXYPlane : sideFromDrawing;
                    if (sideToUse == 0 && sideFromDrawing != 0)
                    {
                        sideToUse = sideFromDrawing;
                    }

                    if (sideToUse != 0 && sideFromDrawing != 0 && (sideToUse * sideFromDrawing) < 0)
                    {
                        sideToUse = sideFromDrawing;
                    }

                    if (sideToUse != 0)
                    {
                        double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalZFromMain) / maxAbsLocalZ);
                        if (normalizedDepth < 1e-6)
                        {
                            normalizedDepth = Math.Min(1.0, Math.Abs(basePerp) / maxAbsBasePerp);
                        }

                        double explodeAmountSheet = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                        double signedExplodeSheet = explodeAmountSheet * sideToUse;
                        explodeOffsetX = signedExplodeSheet * perpSheetX;
                        explodeOffsetY = signedExplodeSheet * perpSheetY;
                    }
                }

                plan.OriginalOffsetX = baseOffsetX;
                plan.OriginalOffsetY = baseOffsetY;
                plan.OffsetX = baseOffsetX + explodeOffsetX;
                plan.OffsetY = baseOffsetY + explodeOffsetY;
                plan.HasComputedOffset = true;
            }

            return true;
        }

        private static bool TryGetModelObjectCenterAndSize(
            object model,
            object identifier,
            out double centerX,
            out double centerY,
            out double centerZ,
            out double sizeScore,
            out bool hasBounds,
            out double minX,
            out double minY,
            out double minZ,
            out double maxX,
            out double maxY,
            out double maxZ)
        {
            centerX = 0.0;
            centerY = 0.0;
            centerZ = 0.0;
            sizeScore = 0.0;
            hasBounds = false;
            minX = 0.0;
            minY = 0.0;
            minZ = 0.0;
            maxX = 0.0;
            maxY = 0.0;
            maxZ = 0.0;

            if (model == null || identifier == null)
            {
                return false;
            }

            object modelObject = InvokeMethod(model, "SelectModelObject", identifier);
            if (modelObject == null)
            {
                return false;
            }

            object solid = InvokeMethod(modelObject, "GetSolid");
            if (solid != null)
            {
                object minimum = GetPropertyValue(solid, "MinimumPoint") ?? GetPropertyValue(solid, "MinPoint");
                object maximum = GetPropertyValue(solid, "MaximumPoint") ?? GetPropertyValue(solid, "MaxPoint");

                double solidMinX;
                double solidMinY;
                double solidMinZ;
                double solidMaxX;
                double solidMaxY;
                double solidMaxZ;
                if (TryGetXYZ(minimum, out solidMinX, out solidMinY, out solidMinZ)
                    && TryGetXYZ(maximum, out solidMaxX, out solidMaxY, out solidMaxZ))
                {
                    minX = solidMinX;
                    minY = solidMinY;
                    minZ = solidMinZ;
                    maxX = solidMaxX;
                    maxY = solidMaxY;
                    maxZ = solidMaxZ;

                    centerX = (solidMinX + solidMaxX) * 0.5;
                    centerY = (solidMinY + solidMaxY) * 0.5;
                    centerZ = (solidMinZ + solidMaxZ) * 0.5;
                    sizeScore = Math.Max(1.0, Math.Max(Math.Abs(solidMaxX - solidMinX), Math.Max(Math.Abs(solidMaxY - solidMinY), Math.Abs(solidMaxZ - solidMinZ))));
                    hasBounds = true;
                    return true;
                }
            }

            object coordinateSystem = InvokeParameterlessMethod(modelObject, "GetCoordinateSystem");
            object origin = GetPropertyValue(coordinateSystem, "Origin");
            if (TryGetXYZ(origin, out centerX, out centerY, out centerZ))
            {
                sizeScore = 1.0;
                return true;
            }

            return false;
        }

        private static bool TryGetMainLocalAxes(
            object model,
            object mainIdentifier,
            out double axisXx,
            out double axisXy,
            out double axisXz,
            out double axisYx,
            out double axisYy,
            out double axisYz,
            out double axisZx,
            out double axisZy,
            out double axisZz)
        {
            axisXx = 1.0;
            axisXy = 0.0;
            axisXz = 0.0;
            axisYx = 0.0;
            axisYy = 1.0;
            axisYz = 0.0;
            axisZx = 0.0;
            axisZy = 0.0;
            axisZz = 1.0;

            if (model == null || mainIdentifier == null)
            {
                return false;
            }

            object mainModelObject = InvokeMethod(model, "SelectModelObject", mainIdentifier);
            if (mainModelObject == null)
            {
                return false;
            }

            object coordinateSystem = InvokeParameterlessMethod(mainModelObject, "GetCoordinateSystem");
            if (coordinateSystem == null)
            {
                return false;
            }

            object axisX = GetPropertyValue(coordinateSystem, "AxisX");
            object axisY = GetPropertyValue(coordinateSystem, "AxisY");
            if (!TryGetXYZ(axisX, out axisXx, out axisXy, out axisXz)
                || !TryGetXYZ(axisY, out axisYx, out axisYy, out axisYz))
            {
                return false;
            }

            if (!NormalizeVector3D(ref axisXx, ref axisXy, ref axisXz)
                || !NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz))
            {
                return false;
            }

            object axisZ = GetPropertyValue(coordinateSystem, "AxisZ");
            if (!TryGetXYZ(axisZ, out axisZx, out axisZy, out axisZz))
            {
                axisZx = (axisXy * axisYz) - (axisXz * axisYy);
                axisZy = (axisXz * axisYx) - (axisXx * axisYz);
                axisZz = (axisXx * axisYy) - (axisXy * axisYx);
            }

            if (!NormalizeVector3D(ref axisZx, ref axisZy, ref axisZz))
            {
                return false;
            }

            return true;
        }

        private static bool TryGetLocalZRange(
            ExplodedPartPlan plan,
            ExplodedPartPlan mainPlan,
            double localAxisZx,
            double localAxisZy,
            double localAxisZz,
            out double minLocalZ,
            out double maxLocalZ)
        {
            return TryGetLocalAxisRange(
                plan,
                mainPlan,
                localAxisZx,
                localAxisZy,
                localAxisZz,
                out minLocalZ,
                out maxLocalZ);
        }

        private static bool TryGetLocalAxisRange(
            ExplodedPartPlan plan,
            ExplodedPartPlan mainPlan,
            double axisX,
            double axisY,
            double axisZ,
            out double minLocal,
            out double maxLocal)
        {
            minLocal = 0.0;
            maxLocal = 0.0;

            if (plan == null || mainPlan == null || !plan.HasModelCenter || !mainPlan.HasModelCenter)
            {
                return false;
            }

            if (!plan.HasModelBounds)
            {
                double dxCenter = plan.ModelCenterX - mainPlan.ModelCenterX;
                double dyCenter = plan.ModelCenterY - mainPlan.ModelCenterY;
                double dzCenter = plan.ModelCenterZ - mainPlan.ModelCenterZ;
                double localCenter = Dot3D(dxCenter, dyCenter, dzCenter, axisX, axisY, axisZ);
                minLocal = localCenter;
                maxLocal = localCenter;
                return true;
            }

            double[] xs = { plan.ModelMinX, plan.ModelMaxX };
            double[] ys = { plan.ModelMinY, plan.ModelMaxY };
            double[] zs = { plan.ModelMinZ, plan.ModelMaxZ };

            bool initialized = false;
            for (int ix = 0; ix < xs.Length; ix++)
            {
                for (int iy = 0; iy < ys.Length; iy++)
                {
                    for (int iz = 0; iz < zs.Length; iz++)
                    {
                        double dx = xs[ix] - mainPlan.ModelCenterX;
                        double dy = ys[iy] - mainPlan.ModelCenterY;
                        double dz = zs[iz] - mainPlan.ModelCenterZ;
                        double localValue = Dot3D(dx, dy, dz, axisX, axisY, axisZ);

                        if (!initialized)
                        {
                            minLocal = localValue;
                            maxLocal = localValue;
                            initialized = true;
                        }
                        else
                        {
                            if (localValue < minLocal)
                            {
                                minLocal = localValue;
                            }

                            if (localValue > maxLocal)
                            {
                                maxLocal = localValue;
                            }
                        }
                    }
                }
            }

            return initialized;
        }

        private static bool TryGetViewDisplayAxes(
            object sourceView,
            out double axisXx,
            out double axisXy,
            out double axisXz,
            out double axisYx,
            out double axisYy,
            out double axisYz)
        {
            axisXx = 1.0;
            axisXy = 0.0;
            axisXz = 0.0;
            axisYx = 0.0;
            axisYy = 1.0;
            axisYz = 0.0;

            object displayCoordinateSystem = GetPropertyValue(sourceView, "DisplayCoordinateSystem");
            if (displayCoordinateSystem == null)
            {
                return false;
            }

            object axisX = GetPropertyValue(displayCoordinateSystem, "AxisX");
            object axisY = GetPropertyValue(displayCoordinateSystem, "AxisY");
            if (!TryGetXYZ(axisX, out axisXx, out axisXy, out axisXz)
                || !TryGetXYZ(axisY, out axisYx, out axisYy, out axisYz))
            {
                return false;
            }

            if (!NormalizeVector3D(ref axisXx, ref axisXy, ref axisXz))
            {
                return false;
            }

            if (!NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz))
            {
                return false;
            }

            return true;
        }

        private static bool NormalizeVector3D(ref double x, ref double y, ref double z)
        {
            double length = Math.Sqrt((x * x) + (y * y) + (z * z));
            if (length < 1e-12)
            {
                return false;
            }

            x /= length;
            y /= length;
            z /= length;
            return true;
        }

        private static bool NormalizeVector2D(ref double x, ref double y)
        {
            double length = Math.Sqrt((x * x) + (y * y));
            if (length < 1e-12)
            {
                return false;
            }

            x /= length;
            y /= length;
            return true;
        }

        private static double Dot3D(double ax, double ay, double az, double bx, double by, double bz)
        {
            return (ax * bx) + (ay * by) + (az * bz);
        }

        private static bool TryComputeOriginalLayoutOffsets(List<ExplodedPartPlan> plans)
        {
            List<ExplodedPartPlan> centeredPlans = new List<ExplodedPartPlan>();
            for (int i = 0; i < plans.Count; i++)
            {
                if (plans[i].HasCenter2D)
                {
                    centeredPlans.Add(plans[i]);
                }
            }

            if (centeredPlans.Count < 2)
            {
                return false;
            }

            ExplodedPartPlan mainPlan = null;
            double bestMainScore = double.MinValue;
            for (int i = 0; i < centeredPlans.Count; i++)
            {
                ExplodedPartPlan candidate = centeredPlans[i];
                if (candidate.ApproxSize2D > bestMainScore)
                {
                    bestMainScore = candidate.ApproxSize2D;
                    mainPlan = candidate;
                }
            }

            if (mainPlan == null)
            {
                return false;
            }

            for (int i = 0; i < centeredPlans.Count; i++)
            {
                ExplodedPartPlan plan = centeredPlans[i];
                plan.OriginalOffsetX = plan.Center2DX - mainPlan.Center2DX;
                plan.OriginalOffsetY = plan.Center2DY - mainPlan.Center2DY;
                plan.OffsetX = plan.Center2DX - mainPlan.Center2DX;
                plan.OffsetY = plan.Center2DY - mainPlan.Center2DY;
                plan.HasComputedOffset = true;
            }

            return true;
        }

        private static void ApplyRadialOffsets(object sourceView, List<ExplodedPartPlan> plans, bool onlyMissing)
        {
            List<ExplodedPartPlan> targetPlans = new List<ExplodedPartPlan>();
            for (int i = 0; i < plans.Count; i++)
            {
                if (!onlyMissing || !plans[i].HasComputedOffset)
                {
                    targetPlans.Add(plans[i]);
                }
            }

            if (targetPlans.Count == 0)
            {
                return;
            }

            double viewWidth = ReadDoubleProperty(sourceView, "Width", 200.0);
            double viewHeight = ReadDoubleProperty(sourceView, "Height", 150.0);
            double maxDimension = Math.Max(viewWidth, viewHeight);
            double minRadius = Math.Max(45.0, maxDimension * 0.55);
            double ringSpacing = Math.Max(18.0, maxDimension * 0.28);
            int firstRingCapacity = Math.Max(8, (int)Math.Ceiling(Math.Sqrt(targetPlans.Count) * 3.0));

            int ringIndex = 0;
            int currentIndex = 0;
            while (currentIndex < targetPlans.Count)
            {
                int ringCapacity = firstRingCapacity + (ringIndex * 4);
                int countInRing = Math.Min(ringCapacity, targetPlans.Count - currentIndex);
                double radius = minRadius + (ringIndex * ringSpacing);

                for (int i = 0; i < countInRing; i++)
                {
                    double angle = (2.0 * Math.PI * i) / Math.Max(1, countInRing);
                    ExplodedPartPlan plan = targetPlans[currentIndex + i];
                    plan.OriginalOffsetX = 0.0;
                    plan.OriginalOffsetY = 0.0;
                    plan.OffsetX = Math.Cos(angle) * radius;
                    plan.OffsetY = Math.Sin(angle) * radius;
                    plan.HasComputedOffset = true;
                }

                currentIndex += countInRing;
                ringIndex++;
            }
        }

        private object CreateSinglePartView(
            object sheet,
            object viewCoordinateSystem,
            object displayCoordinateSystem,
            object identifier,
            double sourceScale,
            string viewName,
            out bool autoIsoApplied)
        {
            autoIsoApplied = false;

            Type viewType = ResolveTeklaType(
                "Tekla.Structures.Drawing.View",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (viewType == null)
            {
                return null;
            }

            ArrayList partList = new ArrayList();
            partList.Add(identifier);

            object view;
            try
            {
                view = Activator.CreateInstance(
                    viewType,
                    new object[] { sheet, viewCoordinateSystem, displayCoordinateSystem, partList });
            }
            catch
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(viewName))
            {
                SetPropertyValue(view, "Name", viewName);
            }

            object viewAttributes = GetPropertyValue(view, "Attributes", "Tekla.Structures.Drawing.View");
            if (viewAttributes != null)
            {
                autoIsoApplied = TryLoadAttributesByName(viewAttributes, ExplodedViewAttributeFile);
                if (!autoIsoApplied)
                {
                    autoIsoApplied = TryLoadAttributesByName(view, ExplodedViewAttributeFile);
                }

                if (!autoIsoApplied)
                {
                    SetPropertyValue(viewAttributes, "Scale", sourceScale);
                }

                SetPropertyValue(viewAttributes, "FixedViewPlacing", true);
                SetPropertyValue(view, "Attributes", viewAttributes, "Tekla.Structures.Drawing.View");
            }

            bool? inserted = InvokeBoolMethod(view, "Insert");
            if (!inserted.HasValue || !inserted.Value)
            {
                return null;
            }

            return view;
        }

        private static bool TryLoadAttributesByName(object target, string attributesName)
        {
            if (target == null || string.IsNullOrWhiteSpace(attributesName))
            {
                return false;
            }

            object result = InvokeMethod(target, "LoadAttributes", attributesName);
            if (result is bool)
            {
                return (bool)result;
            }

            bool parsedResult;
            if (result != null && bool.TryParse(result.ToString(), out parsedResult))
            {
                return parsedResult;
            }

            return result != null;
        }

        private static void RotateViewToIsometric(object view)
        {
            InvokeMethod(view, "RotateViewOnAxisX", -35.264);
            InvokeMethod(view, "RotateViewOnAxisY", 45.0);
        }

        private static double GetSourceViewScale(object sourceView)
        {
            object viewAttributes = GetPropertyValue(sourceView, "Attributes", "Tekla.Structures.Drawing.View");
            if (viewAttributes == null)
            {
                return 1.0;
            }

            object scale = GetPropertyValue(viewAttributes, "Scale");
            double parsedScale;
            return TryConvertToDouble(scale, out parsedScale) && parsedScale > 0.0 ? parsedScale : 1.0;
        }

        private static object CreatePoint(double x, double y, double z)
        {
            Type pointType = ResolveTeklaPointType();
            if (pointType == null)
            {
                return null;
            }

            try
            {
                return Activator.CreateInstance(pointType, new object[] { x, y, z });
            }
            catch
            {
                return null;
            }
        }
        private static object GetPropertyValue(object target, string propertyName, string declaringTypeFullName = null)
        {
            if (target == null)
            {
                return null;
            }

            PropertyInfo[] properties = target.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (declaringTypeFullName != null
                    && (property.DeclaringType == null
                        || !string.Equals(property.DeclaringType.FullName, declaringTypeFullName, StringComparison.Ordinal)))
                {
                    continue;
                }

                if (!property.CanRead || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                try
                {
                    return property.GetValue(target, null);
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        private static bool SetPropertyValue(object target, string propertyName, object value, string declaringTypeFullName = null)
        {
            if (target == null)
            {
                return false;
            }

            PropertyInfo[] properties = target.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (declaringTypeFullName != null
                    && (property.DeclaringType == null
                        || !string.Equals(property.DeclaringType.FullName, declaringTypeFullName, StringComparison.Ordinal)))
                {
                    continue;
                }

                if (!property.CanWrite || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                try
                {
                    object convertedValue = value;
                    if (value != null)
                    {
                        Type propertyType = property.PropertyType;
                        if (!propertyType.IsInstanceOfType(value))
                        {
                            if (propertyType == typeof(double))
                            {
                                double doubleValue;
                                if (TryConvertToDouble(value, out doubleValue))
                                {
                                    convertedValue = doubleValue;
                                }
                            }
                            else if (propertyType == typeof(int))
                            {
                                int intValue;
                                if (int.TryParse(value.ToString(), out intValue))
                                {
                                    convertedValue = intValue;
                                }
                            }
                        }
                    }

                    property.SetValue(target, convertedValue, null);
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        private static double ReadDoubleProperty(object target, string propertyName, double defaultValue)
        {
            object rawValue = GetPropertyValue(target, propertyName);
            double parsedValue;
            if (TryConvertToDouble(rawValue, out parsedValue))
            {
                return parsedValue;
            }

            return defaultValue;
        }

        private static bool TryGetXYZ(object pointOrVector, out double x, out double y, out double z)
        {
            x = 0.0;
            y = 0.0;
            z = 0.0;

            if (pointOrVector == null)
            {
                return false;
            }

            Type type = pointOrVector.GetType();

            FieldInfo fx = type.GetField("X", BindingFlags.Public | BindingFlags.Instance);
            FieldInfo fy = type.GetField("Y", BindingFlags.Public | BindingFlags.Instance);
            FieldInfo fz = type.GetField("Z", BindingFlags.Public | BindingFlags.Instance);

            if (fx != null && fy != null && fz != null)
            {
                object vx = fx.GetValue(pointOrVector);
                object vy = fy.GetValue(pointOrVector);
                object vz = fz.GetValue(pointOrVector);

                return TryConvertToDouble(vx, out x)
                    && TryConvertToDouble(vy, out y)
                    && TryConvertToDouble(vz, out z);
            }

            object px = GetPropertyValue(pointOrVector, "X");
            object py = GetPropertyValue(pointOrVector, "Y");
            object pz = GetPropertyValue(pointOrVector, "Z");

            return TryConvertToDouble(px, out x)
                && TryConvertToDouble(py, out y)
                && TryConvertToDouble(pz, out z);
        }

        private static bool TryConvertToDouble(object value, out double parsed)
        {
            if (value == null)
            {
                parsed = 0.0;
                return false;
            }

            if (value is double)
            {
                parsed = (double)value;
                return true;
            }

            if (value is float)
            {
                parsed = (float)value;
                return true;
            }

            if (value is int)
            {
                parsed = (int)value;
                return true;
            }

            if (value is long)
            {
                parsed = (long)value;
                return true;
            }

            return double.TryParse(value.ToString(), out parsed);
        }

        private static string BuildObjectData(object source, string title, string[] preferredMembers)
        {
            StringBuilder report = new StringBuilder();
            report.AppendLine(title);
            report.AppendLine(new string('-', title.Length));
            report.AppendLine("Tipo CLR: " + source.GetType().FullName);

            HashSet<string> addedMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string memberName in preferredMembers)
            {
                string valueText = TryReadMemberAsText(source, memberName);
                if (valueText == null)
                {
                    continue;
                }

                report.AppendLine(memberName + ": " + valueText);
                addedMembers.Add(memberName);
            }

            PropertyInfo[] publicProperties = source.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            int fallbackCount = 0;
            for (int i = 0; i < publicProperties.Length; i++)
            {
                PropertyInfo property = publicProperties[i];
                if (!property.CanRead || property.GetIndexParameters().Length != 0 || addedMembers.Contains(property.Name))
                {
                    continue;
                }

                object value;
                try
                {
                    value = property.GetValue(source, null);
                }
                catch
                {
                    continue;
                }

                if (!IsSimpleValue(value))
                {
                    continue;
                }

                report.AppendLine(property.Name + ": " + ConvertValueToText(value));
                fallbackCount++;

                if (fallbackCount >= 8)
                {
                    break;
                }
            }

            return report.ToString().TrimEnd();
        }
        private static string TryReadMemberAsText(object source, string memberName)
        {
            try
            {
                BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase;
                PropertyInfo property = source.GetType().GetProperty(memberName, flags);
                if (property != null && property.CanRead && property.GetIndexParameters().Length != 0)
                {
                    return null;
                }

                if (property != null && property.CanRead)
                {
                    object value = property.GetValue(source, null);
                    return ConvertValueToText(value);
                }

                MethodInfo method = source.GetType().GetMethod(memberName, flags, null, Type.EmptyTypes, null);
                if (method != null)
                {
                    object value = method.Invoke(source, null);
                    return ConvertValueToText(value);
                }
            }
            catch (AmbiguousMatchException)
            {
                PropertyInfo[] properties = source.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
                for (int i = 0; i < properties.Length; i++)
                {
                    PropertyInfo property = properties[i];
                    if (!string.Equals(property.Name, memberName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!property.CanRead || property.GetIndexParameters().Length != 0)
                    {
                        continue;
                    }

                    object value = property.GetValue(source, null);
                    return ConvertValueToText(value);
                }
            }

            return null;
        }

        private static bool IsSimpleValue(object value)
        {
            return value == null || IsSimpleType(value.GetType());
        }

        private static bool IsSimpleType(Type type)
        {
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(Guid);
        }

        private static string ConvertValueToText(object value)
        {
            if (value == null)
            {
                return "(null)";
            }

            if (value is DateTime)
            {
                return ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
            }

            Type valueType = value.GetType();
            if (IsSimpleType(valueType))
            {
                return value.ToString();
            }

            PropertyInfo idProperty = valueType.GetProperty("ID", BindingFlags.Public | BindingFlags.Instance)
                ?? valueType.GetProperty("Id", BindingFlags.Public | BindingFlags.Instance);
            if (idProperty != null && idProperty.CanRead)
            {
                object idValue = idProperty.GetValue(value, null);
                if (idValue != null)
                {
                    return idValue.ToString();
                }
            }

            return value.ToString();
        }

        private static string GetInnermostExceptionMessage(Exception exception)
        {
            Exception current = exception;
            while (current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current.Message;
        }

        private static string BoolToText(bool? value)
        {
            if (!value.HasValue)
            {
                return "Nao disponivel";
            }

            return value.Value ? "True" : "False";
        }

        private sealed class ExplodedPartPlan
        {
            public object Identifier;
            public string IdentifierKey;
            public object DrawingObject;
            public bool HasCenter2D;
            public double Center2DX;
            public double Center2DY;
            public double ApproxSize2D;
            public bool IsMainPart;
            public bool IsOtherPlane;
            public bool HasModelCenter;
            public double ModelCenterX;
            public double ModelCenterY;
            public double ModelCenterZ;
            public double ModelSizeScore;
            public bool HasModelBounds;
            public double ModelMinX;
            public double ModelMinY;
            public double ModelMinZ;
            public double ModelMaxX;
            public double ModelMaxY;
            public double ModelMaxZ;
            public double LocalXFromMain;
            public double LocalYFromMain;
            public double LocalZFromMain;
            public int SideOfXYPlane;
            public int SideOfXZPlane;
            public int SideOfZYPlane;
            public int ColorPlane;
            public bool HasComputedOffset;
            public double OriginalOffsetX;
            public double OriginalOffsetY;
            public double OffsetX;
            public double OffsetY;
        }
    }
}
