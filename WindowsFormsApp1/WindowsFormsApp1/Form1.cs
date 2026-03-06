
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private static string[] TeklaModelDllCandidates
        {
            get { return BuildTeklaDllCandidates("Tekla.Structures.Model.dll", true); }
        }

        private static string[] TeklaDrawingDllCandidates
        {
            get { return BuildTeklaDllCandidates("Tekla.Structures.Drawing.dll", true); }
        }

        private static string[] TeklaStructuresDllCandidates
        {
            get { return BuildTeklaDllCandidates("Tekla.Structures.dll", false); }
        }

        private const string ExplodedViewPrefix = "EXP_AUTO_";
        private const int MaxExplodedViews = 80;
        private const bool RotateExplodedToIsometric = true;
        private const string ExplodedViewAttributeFile = "auto-iso";
        private const int MainPartContourColor = 165; // DrawingColors.Magenta
        private const int SidePositiveContourColor = 162; // DrawingColors.Blue
        private const int SideNegativeContourColor = 161; // DrawingColors.Green
        private const int OtherPlaneContourColor = 160; // DrawingColors.Red
        private const string GuideLineNamePrefix = "EXP_AUTO_GUIDE_";
        private const int GuideLineColor = 160; // DrawingColors.Red
        private const string FitRectangleNamePrefix = "EXP_AUTO_RECT_";
        private const string TestAxisNamePrefix = "EXP_AUTO_AXIS_";
        private const int FitRectangleColor = 164; // DrawingColors.Yellow
        private const double FitRectanglePadding = 8.0;
        private const double FitRectangleBorderPadding = 3.0;
        private const double TestAxisPaddingRatio = 0.12;
        private const double TestAxisMinimumLength = 18.0;
        private const double FitScaleFactorMin = 0.20;
        private const double FitScaleFactorMax = 4.00;
        private const int CollapsedClientHeight = 387;
        private const int ExpandedClientHeight = 571;

        private bool logExpanded;
        private Control orbitPreviewControl;
        private bool previewCacheReady;
        private List<ExplodedPartPlan> previewCachedPartPlans = new List<ExplodedPartPlan>();
        private object previewCachedViewCoordinateSystem;
        private object previewCachedDisplayCoordinateSystem;
        private double previewCachedSourceScale = 1.0;

        private static string[] BuildTeklaDllCandidates(string dllName, bool includePlugins)
        {
            List<string> candidates = new List<string>();
            string teklaHome = Environment.GetEnvironmentVariable("TEKLA_HOME");

            AddTeklaDllCandidatesFromRoot(candidates, teklaHome, dllName, includePlugins);

            string[] runningTeklaRoots = GetRunningTeklaInstallationRoots();
            for (int i = 0; i < runningTeklaRoots.Length; i++)
            {
                AddTeklaDllCandidatesFromRoot(candidates, runningTeklaRoots[i], dllName, includePlugins);
            }

            List<string> programFilesRoots = new List<string>();
            AddUniqueCandidate(programFilesRoots, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
            AddUniqueCandidate(programFilesRoots, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
            AddUniqueCandidate(programFilesRoots, Environment.GetEnvironmentVariable("ProgramW6432"));

            for (int rootIndex = 0; rootIndex < programFilesRoots.Count; rootIndex++)
            {
                string programFilesRoot = programFilesRoots[rootIndex];
                if (string.IsNullOrWhiteSpace(programFilesRoot))
                {
                    continue;
                }

                string teklaBasePath = Path.Combine(programFilesRoot, "Tekla Structures");
                if (Directory.Exists(teklaBasePath))
                {
                    string[] installationDirectories = Directory.GetDirectories(teklaBasePath);
                    Array.Sort(installationDirectories, StringComparer.OrdinalIgnoreCase);
                    for (int i = installationDirectories.Length - 1; i >= 0; i--)
                    {
                        AddTeklaDllCandidatesFromRoot(candidates, installationDirectories[i], dllName, includePlugins);
                    }
                }

                // Mantem compatibilidade com layouts antigos caso a varredura dinamica nao encontre nada.
                AddTeklaDllCandidatesFromRoot(
                    candidates,
                    Path.Combine(programFilesRoot, "Tekla Structures", "2024.0"),
                    dllName,
                    includePlugins);
            }

            return candidates.ToArray();
        }

        private static string[] GetRunningTeklaInstallationRoots()
        {
            List<string> roots = new List<string>();
            try
            {
                Process[] processes = Process.GetProcessesByName("TeklaStructures");
                for (int i = 0; i < processes.Length; i++)
                {
                    Process process = processes[i];
                    try
                    {
                        string processPath = null;
                        try
                        {
                            if (process.MainModule != null)
                            {
                                processPath = process.MainModule.FileName;
                            }
                        }
                        catch
                        {
                            processPath = null;
                        }

                        string installationRoot = TryGetTeklaInstallationRootFromProcessPath(processPath);
                        AddUniqueCandidate(roots, installationRoot);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch
            {
                // Mantem fallback para varredura de instalacoes quando leitura de processo nao for permitida.
            }

            return roots.ToArray();
        }

        private static string TryGetTeklaInstallationRootFromProcessPath(string processPath)
        {
            if (string.IsNullOrWhiteSpace(processPath))
            {
                return null;
            }

            try
            {
                DirectoryInfo current = new FileInfo(processPath).Directory;
                while (current != null)
                {
                    DirectoryInfo parent = current.Parent;
                    if (parent != null
                        && string.Equals(parent.Name, "Tekla Structures", StringComparison.OrdinalIgnoreCase))
                    {
                        return current.FullName;
                    }

                    current = parent;
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        private static void AddTeklaDllCandidatesFromRoot(
            List<string> candidates,
            string rootPath,
            string dllName,
            bool includePlugins)
        {
            if (candidates == null
                || string.IsNullOrWhiteSpace(rootPath)
                || string.IsNullOrWhiteSpace(dllName))
            {
                return;
            }

            if (includePlugins)
            {
                AddUniqueCandidate(candidates, Path.Combine(rootPath, "bin", "plugins", dllName));
                AddUniqueCandidate(candidates, Path.Combine(rootPath, "nt", "bin", "plugins", dllName));
            }

            AddUniqueCandidate(candidates, Path.Combine(rootPath, "bin", dllName));
            AddUniqueCandidate(candidates, Path.Combine(rootPath, "nt", "bin", dllName));
            AddUniqueCandidate(candidates, Path.Combine(rootPath, "bin", "Net48Runtime", dllName));
            AddUniqueCandidate(candidates, Path.Combine(rootPath, "nt", "bin", "Net48Runtime", dllName));
        }

        private static void AddUniqueCandidate(List<string> candidates, string candidatePath)
        {
            if (string.IsNullOrWhiteSpace(candidatePath))
            {
                return;
            }

            for (int i = 0; i < candidates.Count; i++)
            {
                if (string.Equals(candidates[i], candidatePath, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            candidates.Add(candidatePath);
        }

        public Form1()
        {
            InitializeComponent();
            ConfigureStatusIndicators();
            SetLogExpanded(false);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            RefreshTeklaStatus(false);
            RefreshDrawingStatus(false);
            RefreshEmbeddedOrbitPreview(false);
        }

        private void btnToggleLog_Click(object sender, EventArgs e)
        {
            SetLogExpanded(!logExpanded);
        }

        private void btnVerificarTekla_Click(object sender, EventArgs e)
        {
            RefreshTeklaStatus(true);
        }

        private void btnVerificarDesenho_Click(object sender, EventArgs e)
        {
            RefreshDrawingStatus(true);
        }

        private void btnFerramentaOrbita_Click(object sender, EventArgs e)
        {
            RefreshEmbeddedOrbitPreview(true);
        }

        private void btnCriarEixos_Click(object sender, EventArgs e)
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
                    SetOutput("Abra um desenho no Tekla antes de criar os eixos.");
                    return;
                }

                object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
                if (sheet == null)
                {
                    UpdateStatus("Falha: nao foi possivel ler a folha do desenho.", Color.DarkRed);
                    SetOutput("Drawing.GetSheet retornou nulo.");
                    return;
                }

                double yawRadians;
                double pitchRadians;
                if (!TryGetPreviewCameraAngles(out yawRadians, out pitchRadians))
                {
                    UpdateStatus("Falha: preview 3D indisponivel para extrair rotacao.", Color.DarkOrange);
                    SetOutput("Atualize o preview 3D e ajuste a camera antes de criar os eixos.");
                    return;
                }

                double rectMinX;
                double rectMinY;
                double rectMaxX;
                double rectMaxY;
                if (!TryGetSheetPlacementRectangle(sheet, out rectMinX, out rectMinY, out rectMaxX, out rectMaxY))
                {
                    UpdateStatus("Falha: nao foi possivel ler os limites da folha.", Color.DarkOrange);
                    SetOutput("TryGetSheetPlacementRectangle falhou para o desenho ativo.");
                    return;
                }

                object sheetOrigin = GetPropertyValue(sheet, "Origin");
                double sheetOriginX;
                double sheetOriginY;
                double sheetOriginZ;
                if (!TryGetXYZ(sheetOrigin, out sheetOriginX, out sheetOriginY, out sheetOriginZ))
                {
                    sheetOriginZ = 0.0;
                }

                double anchorX = (rectMinX + rectMaxX) * 0.5;
                double anchorY = (rectMinY + rectMaxY) * 0.5;
                double anchorZ = sheetOriginZ;

                object model = CreateModelInstance();
                if (model == null)
                {
                    UpdateStatus("Falha: API de model do Tekla nao encontrada.", Color.DarkRed);
                    SetOutput("Nao foi possivel carregar Tekla.Structures.Model.Model.");
                    return;
                }

                bool? modelConnected = InvokeBoolMethod(model, "GetConnectionStatus");
                if (!modelConnected.HasValue || !modelConnected.Value)
                {
                    UpdateStatus("Falha: Tekla aberto, mas sem conexao com o modelo.", Color.DarkOrange);
                    SetOutput("Conecte-se ao modelo no Tekla antes de usar Criar eixos.");
                    return;
                }

                int inspectedViews = 0;
                int candidateViews = 0;
                List<ExplodedPartPlan> partPlans = CollectPartPlansFromDrawing(activeDrawing, out inspectedViews, out candidateViews);
                if (partPlans.Count == 0)
                {
                    UpdateStatus("Falha: nao foi possivel ler as pecas do conjunto no desenho ativo.", Color.DarkOrange);
                    SetOutput("Nenhum ModelIdentifier de parte foi encontrado no contexto do desenho ativo.");
                    return;
                }

                ExplodedPartPlan mainPlan = null;
                double bestMainScore = double.MinValue;
                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];
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

                    plan.ModelSizeScore = sizeScore;
                    if (sizeScore > bestMainScore)
                    {
                        bestMainScore = sizeScore;
                        mainPlan = plan;
                    }
                }

                if (mainPlan == null || mainPlan.Identifier == null)
                {
                    UpdateStatus("Falha: nao foi possivel identificar a peca principal.", Color.DarkOrange);
                    SetOutput("Nao foi possivel determinar a peca principal para centralizar a vista.");
                    return;
                }

                for (int i = 0; i < partPlans.Count; i++)
                {
                    partPlans[i].IsMainPart = ReferenceEquals(partPlans[i], mainPlan);
                }

                object viewCoordinateSystem = previewCachedViewCoordinateSystem;
                object displayCoordinateSystem = previewCachedDisplayCoordinateSystem;
                bool coordinateSystemFromMainPart = false;
                bool coordinateSystemFromPreviewCamera = false;

                double previewAxisXx;
                double previewAxisXy;
                double previewAxisXz;
                double previewAxisYx;
                double previewAxisYy;
                double previewAxisYz;
                if (TryGetPreviewDisplayAxesForView(
                        null,
                        out previewAxisXx,
                        out previewAxisXy,
                        out previewAxisXz,
                        out previewAxisYx,
                        out previewAxisYy,
                        out previewAxisYz))
                {
                    object previewCameraCoordinateSystem = CreateCoordinateSystem(
                        0.0,
                        0.0,
                        0.0,
                        previewAxisXx,
                        previewAxisXy,
                        previewAxisXz,
                        previewAxisYx,
                        previewAxisYy,
                        previewAxisYz);
                    if (previewCameraCoordinateSystem != null)
                    {
                        viewCoordinateSystem = previewCameraCoordinateSystem;
                        displayCoordinateSystem = previewCameraCoordinateSystem;
                        coordinateSystemFromPreviewCamera = true;
                    }
                }

                if (viewCoordinateSystem == null || displayCoordinateSystem == null)
                {
                    double localAxisXx;
                    double localAxisXy;
                    double localAxisXz;
                    double localAxisYx;
                    double localAxisYy;
                    double localAxisYz;
                    double localAxisZx;
                    double localAxisZy;
                    double localAxisZz;
                    if (TryGetMainLocalAxes(
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
                        object mainPartCoordinateSystem = CreateCoordinateSystem(
                            0.0,
                            0.0,
                            0.0,
                            localAxisXx,
                            localAxisXy,
                            localAxisXz,
                            localAxisYx,
                            localAxisYy,
                            localAxisYz);
                        if (mainPartCoordinateSystem != null)
                        {
                            if (viewCoordinateSystem == null)
                            {
                                viewCoordinateSystem = mainPartCoordinateSystem;
                            }

                            if (displayCoordinateSystem == null)
                            {
                                displayCoordinateSystem = mainPartCoordinateSystem;
                            }

                            coordinateSystemFromMainPart = true;
                        }
                    }
                }

                if (viewCoordinateSystem == null)
                {
                    viewCoordinateSystem = CreateGlobalCoordinateSystem();
                }

                if (displayCoordinateSystem == null)
                {
                    displayCoordinateSystem = CreateGlobalCoordinateSystem();
                }

                if (viewCoordinateSystem == null || displayCoordinateSystem == null)
                {
                    UpdateStatus("Falha: coordenadas base indisponiveis para criar a vista.", Color.DarkRed);
                    SetOutput("Nao foi possivel obter ou criar ViewCoordinateSystem/DisplayCoordinateSystem.");
                    return;
                }

                int removedViews = DeleteExistingExplodedViews(sheet, ExplodedViewPrefix);
                int removedGuideLines = DeleteExistingGuideLines(sheet);
                int removedAxes = DeleteExistingTestAxes(sheet);

                ArrayList partIdentifiers = new ArrayList();
                HashSet<string> includedIds = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];
                    if (plan == null || plan.Identifier == null)
                    {
                        continue;
                    }

                    string key = !string.IsNullOrWhiteSpace(plan.IdentifierKey)
                        ? plan.IdentifierKey
                        : GetIdentifierKey(plan.Identifier);
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        key = "idx_" + i.ToString();
                    }

                    if (includedIds.Add(key))
                    {
                        partIdentifiers.Add(plan.Identifier);
                    }
                }

                int identifiersInView = partIdentifiers.Count;
                if (identifiersInView == 0)
                {
                    UpdateStatus("Falha: nao foi possivel montar a lista de pecas para criar a vista.", Color.DarkOrange);
                    SetOutput("As pecas foram identificadas, mas nenhum ModelIdentifier valido foi incluido na vista.");
                    return;
                }

                bool mainViewAutoIsoApplied;
                double mainViewScale = previewCachedSourceScale > 0.0 ? previewCachedSourceScale : 1.0;
                object mainView = CreatePartListViewForFit(
                    sheet,
                    viewCoordinateSystem,
                    displayCoordinateSystem,
                    partIdentifiers,
                    mainViewScale,
                    ExplodedViewPrefix + "SET",
                    out mainViewAutoIsoApplied);

                bool mainViewRotationApplied = false;
                bool mainViewPositioned = false;
                bool fitApplied = false;
                double fitFactorApplied;
                int fitAdjustedViews;

                if (mainView != null)
                {
                    if (!coordinateSystemFromPreviewCamera)
                    {
                        const double ToDegrees = 180.0 / Math.PI;
                        RotateViewToAngles(mainView, -pitchRadians * ToDegrees, -yawRadians * ToDegrees);
                        mainViewRotationApplied = true;
                    }
                    mainViewPositioned = ForceViewCenterToTarget(mainView, anchorX, anchorY, anchorZ);

                    if (mainViewPositioned)
                    {
                        List<object> fitViews = new List<object>();
                        fitViews.Add(mainView);

                        Dictionary<string, double> fitViewScales = new Dictionary<string, double>();
                        fitViewScales[GetDrawingObjectKey(mainView)] = mainViewScale;

                        fitApplied = TryFitCreatedViewsIntoRectangle(
                            fitViews,
                            fitViewScales,
                            rectMinX,
                            rectMinY,
                            rectMaxX,
                            rectMaxY,
                            anchorZ,
                            out fitFactorApplied,
                            out fitAdjustedViews);
                    }
                }

                double rectWidth = Math.Max(1.0, Math.Abs(rectMaxX - rectMinX));
                double rectHeight = Math.Max(1.0, Math.Abs(rectMaxY - rectMinY));
                double axisLength = Math.Max(TestAxisMinimumLength, Math.Min(rectWidth, rectHeight) * 0.38);
                double cosYaw = Math.Cos(yawRadians);
                double sinYaw = Math.Sin(yawRadians);
                double cosPitch = Math.Cos(pitchRadians);
                double sinPitch = Math.Sin(pitchRadians);
                double previewAxisXDirX = cosYaw;
                double previewAxisXDirY = sinPitch * sinYaw;
                double previewAxisYDirX = 0.0;
                double previewAxisYDirY = cosPitch;
                double previewAxisZDirX = sinYaw;
                double previewAxisZDirY = -sinPitch * cosYaw;

                int axisCreated = 0;
                if (mainView != null)
                {
                    double viewScaleForAxes = GetSourceViewScale(mainView);
                    if (!(viewScaleForAxes > 0.0))
                    {
                        viewScaleForAxes = mainViewScale > 0.0 ? mainViewScale : 1.0;
                    }

                    axisCreated = TryCreateMainPartTestAxesOverlay(
                        sheet,
                        mainView,
                        model,
                        partPlans,
                        viewScaleForAxes,
                        anchorX,
                        anchorY,
                        anchorZ,
                        true,
                        true,
                        true,
                        true,
                        previewAxisXDirX,
                        previewAxisXDirY,
                        previewAxisYDirX,
                        previewAxisYDirY,
                        previewAxisZDirX,
                        previewAxisZDirY);
                }

                if (axisCreated <= 0)
                {
                    axisCreated = TryCreatePreviewAlignedTestAxesOverlay(
                        sheet,
                        anchorX,
                        anchorY,
                        anchorZ,
                        axisLength,
                        previewAxisXDirX,
                        previewAxisXDirY,
                        previewAxisYDirX,
                        previewAxisYDirY,
                        previewAxisZDirX,
                        previewAxisZDirY,
                        true,
                        true,
                        true);
                }

                if (removedViews > 0
                    || removedGuideLines > 0
                    || removedAxes > 0
                    || mainView != null
                    || axisCreated > 0)
                {
                    CommitActiveDrawingChanges(activeDrawing);
                    InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");
                }

                StringBuilder report = new StringBuilder();
                report.AppendLine("Criar eixos (sem explosao)");
                report.AppendLine("--------------------------");
                report.AppendLine("Conjunto lido do desenho ativo.");
                report.AppendLine("Vistas inspecionadas: " + inspectedViews);
                report.AppendLine("Vistas com pecas: " + candidateViews);
                report.AppendLine("Pecas candidatas: " + partPlans.Count);
                report.AppendLine("Pecas incluidas na vista: " + identifiersInView);
                report.AppendLine("Peca principal: " + (string.IsNullOrWhiteSpace(mainPlan.IdentifierKey) ? GetIdentifierKey(mainPlan.Identifier) : mainPlan.IdentifierKey));
                report.AppendLine("Centro usado: centro da folha");
                report.AppendLine("Centro (X,Y): " + anchorX.ToString("0.###") + ", " + anchorY.ToString("0.###"));
                report.AppendLine("Area da folha: W=" + rectWidth.ToString("0.###") + " H=" + rectHeight.ToString("0.###"));
                report.AppendLine("Sistema de coordenadas base: "
                    + (previewCachedDisplayCoordinateSystem != null ? "cache do preview" : (coordinateSystemFromMainPart ? "eixo local da peca principal" : "modelo global")));
                report.AppendLine("Sistema de coordenadas da vista principal: " + (coordinateSystemFromPreviewCamera ? "camera do preview" : "base + rotacao por angulos"));
                report.AppendLine("Vistas auto antigas removidas: " + removedViews);
                report.AppendLine("Linhas guia antigas removidas: " + removedGuideLines);
                report.AppendLine("Eixos antigos removidos: " + removedAxes);
                report.AppendLine("Vista do conjunto criada: " + (mainView != null ? "sim" : "nao"));
                report.AppendLine("Rotacao preview aplicada na vista principal: " + (mainViewRotationApplied ? "sim" : "nao"));
                report.AppendLine("Vista do conjunto centralizada: " + (mainViewPositioned ? "sim" : "nao"));
                report.AppendLine("Comprimento dos eixos: " + axisLength.ToString("0.###"));
                report.AppendLine("Rotacao camera preview: yaw=" + (yawRadians * (180.0 / Math.PI)).ToString("0.###")
                    + "deg, pitch=" + (pitchRadians * (180.0 / Math.PI)).ToString("0.###") + "deg");
                report.AppendLine("Eixos criados: " + axisCreated);
                report.AppendLine();
                report.AppendLine("Obs.: eixos e vista do conjunto alinhados com a camera atual do preview 3D.");
                SetOutput(report.ToString());

                if (axisCreated > 0 && mainViewPositioned)
                {
                    UpdateStatus("Sucesso: eixos criados e conjunto centralizado com rotacao do preview.", Color.DarkGreen);
                }
                else if (axisCreated > 0 && mainView != null)
                {
                    UpdateStatus("Parcial: eixos criados, mas falha ao centralizar a vista do conjunto.", Color.DarkOrange);
                }
                else
                {
                    UpdateStatus("Falha: nao foi possivel criar eixos e vista do conjunto no centro.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao criar eixos: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void RefreshEmbeddedOrbitPreview(bool updateStatus)
        {
            try
            {
                string previewSummary;
                List<OrbitToolForm.PreviewPart> previewParts = BuildOrbitPreviewParts(out previewSummary);
                Control previewSurface = OrbitToolForm.CreatePreviewSurface(previewParts, previewSummary);
                ReplaceOrbitPreviewSurface(previewSurface);

                if (updateStatus)
                {
                    if (previewParts.Count > 0)
                    {
                        UpdateStatus("Preview 3D atualizado na tela principal.", Color.DarkGreen);
                    }
                    else
                    {
                        UpdateStatus(previewSummary, Color.DarkOrange);
                    }
                }
            }
            catch (Exception ex)
            {
                if (updateStatus)
                {
                    UpdateStatus("Erro ao atualizar preview 3D: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
                }
            }
        }

        private void ReplaceOrbitPreviewSurface(Control previewSurface)
        {
            if (pnlPreviewHost == null || previewSurface == null)
            {
                return;
            }

            pnlPreviewHost.SuspendLayout();
            try
            {
                if (orbitPreviewControl != null)
                {
                    pnlPreviewHost.Controls.Remove(orbitPreviewControl);
                    orbitPreviewControl.Dispose();
                    orbitPreviewControl = null;
                }

                previewSurface.Dock = DockStyle.Fill;
                orbitPreviewControl = previewSurface;
                pnlPreviewHost.Controls.Add(orbitPreviewControl);
                orbitPreviewControl.BringToFront();
            }
            finally
            {
                pnlPreviewHost.ResumeLayout();
            }
        }

        private bool TryGetPreviewCameraAngles(out double yawRadians, out double pitchRadians)
        {
            yawRadians = 0.0;
            pitchRadians = 0.0;

            OrbitToolForm.PreviewSurfaceControl previewSurface =
                orbitPreviewControl as OrbitToolForm.PreviewSurfaceControl;
            if (previewSurface == null)
            {
                return false;
            }

            return previewSurface.TryGetCameraAngles(out yawRadians, out pitchRadians);
        }

        private bool TryGetPreviewRotationDegrees(out double rotationXDegrees, out double rotationYDegrees)
        {
            rotationXDegrees = 0.0;
            rotationYDegrees = 0.0;

            double yawRadians;
            double pitchRadians;
            if (!TryGetPreviewCameraAngles(out yawRadians, out pitchRadians))
            {
                return false;
            }

            const double ToDegrees = 180.0 / Math.PI;
            rotationXDegrees = -pitchRadians * ToDegrees;
            rotationYDegrees = -yawRadians * ToDegrees;
            return true;
        }

        private bool TryGetPreviewDisplayAxesForView(
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

            if (sourceView == null)
            {
                // Sem vista: usa eixos globais como base e, se houver cache valido de preview, tenta usar
                // o DisplayCoordinateSystem salvo.
            }

            double yawRadians;
            double pitchRadians;
            if (!TryGetPreviewCameraAngles(out yawRadians, out pitchRadians))
            {
                return false;
            }

            double viewAxisXx = 1.0;
            double viewAxisXy = 0.0;
            double viewAxisXz = 0.0;
            double viewAxisYx = 0.0;
            double viewAxisYy = 1.0;
            double viewAxisYz = 0.0;
            bool hasViewAxes = false;

            if (sourceView != null)
            {
                hasViewAxes = TryGetViewDisplayAxes(
                    sourceView,
                    out viewAxisXx,
                    out viewAxisXy,
                    out viewAxisXz,
                    out viewAxisYx,
                    out viewAxisYy,
                    out viewAxisYz);
            }
            else if (previewCachedDisplayCoordinateSystem != null)
            {
                object cachedAxisX = GetPropertyValue(previewCachedDisplayCoordinateSystem, "AxisX");
                object cachedAxisY = GetPropertyValue(previewCachedDisplayCoordinateSystem, "AxisY");
                hasViewAxes = TryGetXYZ(cachedAxisX, out viewAxisXx, out viewAxisXy, out viewAxisXz)
                    && TryGetXYZ(cachedAxisY, out viewAxisYx, out viewAxisYy, out viewAxisYz)
                    && NormalizeVector3D(ref viewAxisXx, ref viewAxisXy, ref viewAxisXz)
                    && NormalizeVector3D(ref viewAxisYx, ref viewAxisYy, ref viewAxisYz);
            }

            if (!hasViewAxes)
            {
                viewAxisXx = 1.0;
                viewAxisXy = 0.0;
                viewAxisXz = 0.0;
                viewAxisYx = 0.0;
                viewAxisYy = 1.0;
                viewAxisYz = 0.0;
            }

            double viewAxisZx = (viewAxisXy * viewAxisYz) - (viewAxisXz * viewAxisYy);
            double viewAxisZy = (viewAxisXz * viewAxisYx) - (viewAxisXx * viewAxisYz);
            double viewAxisZz = (viewAxisXx * viewAxisYy) - (viewAxisXy * viewAxisYx);
            if (!NormalizeVector3D(ref viewAxisZx, ref viewAxisZy, ref viewAxisZz))
            {
                return false;
            }

            double cosYaw = Math.Cos(yawRadians);
            double sinYaw = Math.Sin(yawRadians);
            double cosPitch = Math.Cos(pitchRadians);
            double sinPitch = Math.Sin(pitchRadians);

            double localRightX = cosYaw;
            double localRightY = 0.0;
            double localRightZ = sinYaw;
            double localUpX = sinPitch * sinYaw;
            double localUpY = cosPitch;
            double localUpZ = -sinPitch * cosYaw;

            axisXx = (localRightX * viewAxisXx) + (localRightY * viewAxisYx) + (localRightZ * viewAxisZx);
            axisXy = (localRightX * viewAxisXy) + (localRightY * viewAxisYy) + (localRightZ * viewAxisZy);
            axisXz = (localRightX * viewAxisXz) + (localRightY * viewAxisYz) + (localRightZ * viewAxisZz);

            axisYx = (localUpX * viewAxisXx) + (localUpY * viewAxisYx) + (localUpZ * viewAxisZx);
            axisYy = (localUpX * viewAxisXy) + (localUpY * viewAxisYy) + (localUpZ * viewAxisZy);
            axisYz = (localUpX * viewAxisXz) + (localUpY * viewAxisYz) + (localUpZ * viewAxisZz);

            if (!NormalizeVector3D(ref axisXx, ref axisXy, ref axisXz)
                || !NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz))
            {
                return false;
            }
            return true;
        }

        private void ApplyFallbackRotation(object view, bool autoIsoApplied, ref int fallbackRotationCount)
        {
            if (view == null)
            {
                return;
            }

            if (chkRotacaoPreview != null && chkRotacaoPreview.Checked)
            {
                double rotationXDegrees;
                double rotationYDegrees;
                if (TryGetPreviewRotationDegrees(out rotationXDegrees, out rotationYDegrees))
                {
                    RotateViewToAngles(view, rotationXDegrees, rotationYDegrees);
                    fallbackRotationCount++;
                    return;
                }
            }

            if (!autoIsoApplied && RotateExplodedToIsometric)
            {
                RotateViewToIsometric(view);
                fallbackRotationCount++;
            }
        }

        // Fluxo paralelo de preview 3D: somente leitura dos dados atuais, sem alterar as logicas de explosao.
        private List<OrbitToolForm.PreviewPart> BuildOrbitPreviewParts(out string summaryText)
        {
            List<OrbitToolForm.PreviewPart> previewParts = new List<OrbitToolForm.PreviewPart>();
            summaryText = "Preview 3D indisponivel.";
            object previewWorkPlaneHandler = null;
            object previewPreviousTransformationPlane = null;
            bool previewGlobalPlaneApplied = false;
            string previewTransformationPlaneNote = "nao alterado";

            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    summaryText = "Falha: API de drawing do Tekla nao encontrada.";
                    return previewParts;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    summaryText = "Abra um desenho no Tekla para carregar o preview 3D.";
                    return previewParts;
                }

                string sourceDescription = "conjunto do desenho ativo";
                int inspectedViews = 0;
                int candidateViews = 0;
                List<ExplodedPartPlan> partPlans = CollectPartPlansFromDrawing(activeDrawing, out inspectedViews, out candidateViews);
                if (partPlans.Count == 0)
                {
                    summaryText = "Nenhuma peca legivel foi encontrada no contexto do desenho ativo.";
                    return previewParts;
                }

                object model = CreateModelInstance();
                if (model == null)
                {
                    summaryText = "Falha: conexao com o modelo indisponivel para montar o preview.";
                    return previewParts;
                }

                bool? modelConnected = InvokeBoolMethod(model, "GetConnectionStatus");
                if (!modelConnected.HasValue || !modelConnected.Value)
                {
                    summaryText = "Falha: Tekla aberto, mas sem conexao com o modelo.";
                    return previewParts;
                }

                previewGlobalPlaneApplied = TrySetGlobalTransformationPlane(
                    model,
                    out previewWorkPlaneHandler,
                    out previewPreviousTransformationPlane);
                previewTransformationPlaneNote = previewGlobalPlaneApplied
                    ? "global"
                    : "mantido (nao foi possivel forcar global)";

                List<ExplodedPartPlan> modelPlans = new List<ExplodedPartPlan>();
                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];
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

                    if (!hasBounds)
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
                    plan.IsMainPart = false;
                    modelPlans.Add(plan);
                }

                if (modelPlans.Count == 0)
                {
                    summaryText = "Nao foi possivel montar bounds 3D das pecas do conjunto.";
                    return previewParts;
                }

                ExplodedPartPlan mainPlan = TryResolveMainPlanFromAssembly(model, modelPlans);
                if (mainPlan == null)
                {
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
                }

                if (mainPlan != null)
                {
                    mainPlan.IsMainPart = true;
                }

                double originX = 0.0;
                double originY = 0.0;
                double originZ = 0.0;
                double axisXx = 1.0;
                double axisXy = 0.0;
                double axisXz = 0.0;
                double axisYx = 0.0;
                double axisYy = 1.0;
                double axisYz = 0.0;
                double axisZx = 0.0;
                double axisZy = 0.0;
                double axisZz = 1.0;
                bool useMainPartLocalCoordinates = false;
                string previewAxisNormalizationNote = "nao aplicada";

                if (mainPlan != null && mainPlan.Identifier != null)
                {
                    originX = mainPlan.ModelCenterX;
                    originY = mainPlan.ModelCenterY;
                    originZ = mainPlan.ModelCenterZ;
                    useMainPartLocalCoordinates = TryGetMainLocalAxes(
                        model,
                        mainPlan.Identifier,
                        out axisXx,
                        out axisXy,
                        out axisXz,
                        out axisYx,
                        out axisYy,
                        out axisYz,
                        out axisZx,
                        out axisZy,
                        out axisZz);

                    if (useMainPartLocalCoordinates)
                    {
                        NormalizePreviewBasisToWorldOrientation(
                            ref axisXx,
                            ref axisXy,
                            ref axisXz,
                            ref axisYx,
                            ref axisYy,
                            ref axisYz,
                            ref axisZx,
                            ref axisZy,
                            ref axisZz,
                            out previewAxisNormalizationNote);

                        ApplyPreviewAxisConvention(
                            ref axisXx,
                            ref axisXy,
                            ref axisXz,
                            ref axisYx,
                            ref axisYy,
                            ref axisYz,
                            ref axisZx,
                            ref axisZy,
                            ref axisZz,
                            ref previewAxisNormalizationNote);
                    }
                }

                object previewCoordinateSystem = useMainPartLocalCoordinates
                    ? CreateCoordinateSystem(0.0, 0.0, 0.0, axisXx, axisXy, axisXz, axisYx, axisYy, axisYz)
                    : CreateGlobalCoordinateSystem();
                previewCachedSourceScale = 1.0;
                previewCachedViewCoordinateSystem = previewCoordinateSystem;
                previewCachedDisplayCoordinateSystem = previewCoordinateSystem;
                previewCachedPartPlans = new List<ExplodedPartPlan>(modelPlans);
                previewCacheReady = previewCachedViewCoordinateSystem != null
                    && previewCachedDisplayCoordinateSystem != null
                    && modelPlans.Count > 0;
                int meshPartsLoaded = 0;

                for (int i = 0; i < modelPlans.Count; i++)
                {
                    ExplodedPartPlan plan = modelPlans[i];
                    double previewMinX = plan.ModelMinX;
                    double previewMinY = plan.ModelMinY;
                    double previewMinZ = plan.ModelMinZ;
                    double previewMaxX = plan.ModelMaxX;
                    double previewMaxY = plan.ModelMaxY;
                    double previewMaxZ = plan.ModelMaxZ;

                    if (useMainPartLocalCoordinates)
                    {
                        double localMinX;
                        double localMinY;
                        double localMinZ;
                        double localMaxX;
                        double localMaxY;
                        double localMaxZ;
                        if (TryTransformModelBoundsToViewLocal(
                                plan.ModelMinX,
                                plan.ModelMinY,
                                plan.ModelMinZ,
                                plan.ModelMaxX,
                                plan.ModelMaxY,
                                plan.ModelMaxZ,
                                originX,
                                originY,
                                originZ,
                                axisXx,
                                axisXy,
                                axisXz,
                                axisYx,
                                axisYy,
                                axisYz,
                                axisZx,
                                axisZy,
                                axisZz,
                                out localMinX,
                                out localMinY,
                                out localMinZ,
                                out localMaxX,
                                out localMaxY,
                                out localMaxZ))
                        {
                            previewMinX = localMinX;
                            previewMinY = localMinY;
                            previewMinZ = localMinZ;
                            previewMaxX = localMaxX;
                            previewMaxY = localMaxY;
                            previewMaxZ = localMaxZ;
                        }
                    }

                    string label = string.IsNullOrWhiteSpace(plan.IdentifierKey)
                        ? "Parte " + (i + 1).ToString()
                        : plan.IdentifierKey;

                    List<OrbitToolForm.PreviewTriangle> previewTriangles =
                        TryBuildPreviewTrianglesFromModel(
                            model,
                            plan.Identifier,
                            useMainPartLocalCoordinates,
                            originX,
                            originY,
                            originZ,
                            axisXx,
                            axisXy,
                            axisXz,
                            axisYx,
                            axisYy,
                            axisYz,
                            axisZx,
                            axisZy,
                            axisZz);
                    if (previewTriangles.Count > 0)
                    {
                        meshPartsLoaded++;
                    }

                    previewParts.Add(
                        new OrbitToolForm.PreviewPart(
                            label,
                            previewMinX,
                            previewMinY,
                            previewMinZ,
                            previewMaxX,
                            previewMaxY,
                            previewMaxZ,
                            plan.IsMainPart,
                            previewTriangles));
                }

                summaryText = previewParts.Count > 0
                    ? "Preview 3D do "
                        + sourceDescription
                        + ". Vistas inspecionadas: "
                        + inspectedViews
                        + ". Vistas com pecas: "
                        + candidateViews
                        + ". Pecas carregadas: "
                        + previewParts.Count
                        + "/"
                        + modelPlans.Count
                        + ". Coordenadas usadas: "
                        + (useMainPartLocalCoordinates ? "eixo local da peca principal" : "modelo global")
                        + ". Geometria real: "
                        + meshPartsLoaded
                        + "/"
                        + modelPlans.Count
                        + " pecas. Plano de trabalho: "
                        + previewTransformationPlaneNote
                        + ". Normalizacao de eixo: "
                        + previewAxisNormalizationNote
                        + "."
                    : "Nao foi possivel montar bounds 3D das pecas encontradas no desenho ativo.";
            }
            catch (Exception ex)
            {
                summaryText = "Erro ao montar preview 3D: " + GetInnermostExceptionMessage(ex);
            }
            finally
            {
                if (previewGlobalPlaneApplied)
                {
                    TryRestoreTransformationPlane(previewWorkPlaneHandler, previewPreviousTransformationPlane);
                }
            }

            return previewParts;
        }

        private static void NormalizePreviewBasisToWorldOrientation(
            ref double axisXx,
            ref double axisXy,
            ref double axisXz,
            ref double axisYx,
            ref double axisYy,
            ref double axisYz,
            ref double axisZx,
            ref double axisZy,
            ref double axisZz,
            out string normalizationNote)
        {
            normalizationNote = "preservada";
            if (!NormalizeVector3D(ref axisXx, ref axisXy, ref axisXz)
                || !NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz))
            {
                normalizationNote = "falha normalizar XY";
                return;
            }

            // Mantem o eixo X da peca principal e ortonormaliza Y sem trocar sinais para "alinhar"
            // com o global; isso evita espelhamento visual do conjunto no preview.
            double dotXY = Dot3D(axisYx, axisYy, axisYz, axisXx, axisXy, axisXz);
            axisYx -= dotXY * axisXx;
            axisYy -= dotXY * axisXy;
            axisYz -= dotXY * axisXz;
            if (!NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz))
            {
                normalizationNote = "falha ortonormalizar Y";
                return;
            }

            double targetZx = axisZx;
            double targetZy = axisZy;
            double targetZz = axisZz;
            bool hasTargetZ = NormalizeVector3D(ref targetZx, ref targetZy, ref targetZz);

            axisZx = (axisXy * axisYz) - (axisXz * axisYy);
            axisZy = (axisXz * axisYx) - (axisXx * axisYz);
            axisZz = (axisXx * axisYy) - (axisXy * axisYx);
            if (!NormalizeVector3D(ref axisZx, ref axisZy, ref axisZz))
            {
                normalizationNote = "falha calcular Z";
                return;
            }

            if (hasTargetZ && Dot3D(axisZx, axisZy, axisZz, targetZx, targetZy, targetZz) < 0.0)
            {
                axisYx = -axisYx;
                axisYy = -axisYy;
                axisYz = -axisYz;
                axisZx = -axisZx;
                axisZy = -axisZy;
                axisZz = -axisZz;
                normalizationNote = "ajuste sinal YZ";
            }
        }

        private static bool TrySetGlobalTransformationPlane(
            object model,
            out object workPlaneHandler,
            out object previousTransformationPlane)
        {
            workPlaneHandler = null;
            previousTransformationPlane = null;
            if (model == null)
            {
                return false;
            }

            workPlaneHandler = InvokeParameterlessMethod(model, "GetWorkPlaneHandler");
            if (workPlaneHandler == null)
            {
                return false;
            }

            previousTransformationPlane = InvokeParameterlessMethod(workPlaneHandler, "GetCurrentTransformationPlane");

            Type transformationPlaneType = ResolveTeklaType(
                "Tekla.Structures.Model.TransformationPlane",
                "Tekla.Structures.Model",
                TeklaModelDllCandidates);
            if (transformationPlaneType == null)
            {
                return false;
            }

            object globalTransformationPlane;
            try
            {
                globalTransformationPlane = Activator.CreateInstance(transformationPlaneType);
            }
            catch
            {
                return false;
            }

            if (globalTransformationPlane == null)
            {
                return false;
            }

            try
            {
                object setResult = InvokeMethod(workPlaneHandler, "SetCurrentTransformationPlane", globalTransformationPlane);
                if (setResult is bool && !((bool)setResult))
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static void ApplyPreviewAxisConvention(
            ref double axisXx,
            ref double axisXy,
            ref double axisXz,
            ref double axisYx,
            ref double axisYy,
            ref double axisYz,
            ref double axisZx,
            ref double axisZy,
            ref double axisZz,
            ref string note)
        {
            if (!NormalizeVector3D(ref axisXx, ref axisXy, ref axisXz)
                || !NormalizeVector3D(ref axisYx, ref axisYy, ref axisYz)
                || !NormalizeVector3D(ref axisZx, ref axisZy, ref axisZz))
            {
                note = string.IsNullOrWhiteSpace(note)
                    ? "falha convencao preview"
                    : (note + "; falha convencao preview");
                return;
            }

            // Convencao visual do preview (Tekla drawing): espelha apenas X.
            // Y e Z permanecem como vieram do eixo local da peca principal.
            axisXx = -axisXx;
            axisXy = -axisXy;
            axisXz = -axisXz;

            note = string.IsNullOrWhiteSpace(note)
                ? "convencao preview: flip X"
                : (note + "; convencao preview: flip X");
        }

        private static void TryRestoreTransformationPlane(
            object workPlaneHandler,
            object previousTransformationPlane)
        {
            if (workPlaneHandler == null || previousTransformationPlane == null)
            {
                return;
            }

            try
            {
                InvokeMethod(workPlaneHandler, "SetCurrentTransformationPlane", previousTransformationPlane);
            }
            catch
            {
                // Restauracao best effort.
            }
        }

        private static ExplodedPartPlan TryResolveMainPlanFromAssembly(
            object model,
            List<ExplodedPartPlan> modelPlans)
        {
            if (model == null || modelPlans == null || modelPlans.Count == 0)
            {
                return null;
            }

            Dictionary<string, ExplodedPartPlan> byId = new Dictionary<string, ExplodedPartPlan>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                if (plan == null || plan.Identifier == null)
                {
                    continue;
                }

                string id = !string.IsNullOrWhiteSpace(plan.IdentifierKey)
                    ? plan.IdentifierKey
                    : GetIdentifierKey(plan.Identifier);
                if (string.IsNullOrWhiteSpace(id) || byId.ContainsKey(id))
                {
                    continue;
                }

                byId.Add(id, plan);
            }

            for (int i = 0; i < modelPlans.Count; i++)
            {
                ExplodedPartPlan plan = modelPlans[i];
                if (plan == null || plan.Identifier == null)
                {
                    continue;
                }

                object modelObject = InvokeMethod(model, "SelectModelObject", plan.Identifier);
                if (modelObject == null)
                {
                    continue;
                }

                object assemblyObject = InvokeParameterlessMethod(modelObject, "GetAssembly");
                if (assemblyObject == null)
                {
                    continue;
                }

                object mainPartObject = InvokeParameterlessMethod(assemblyObject, "GetMainPart");
                if (mainPartObject == null)
                {
                    continue;
                }

                object mainIdentifier = GetPropertyValue(mainPartObject, "Identifier");
                string mainId = GetIdentifierKey(mainIdentifier);
                if (string.IsNullOrWhiteSpace(mainId))
                {
                    continue;
                }

                ExplodedPartPlan resolved;
                if (byId.TryGetValue(mainId, out resolved))
                {
                    return resolved;
                }
            }

            return null;
        }

        private static bool TryTransformModelBoundsToViewLocal(
            double minX,
            double minY,
            double minZ,
            double maxX,
            double maxY,
            double maxZ,
            double originX,
            double originY,
            double originZ,
            double axisXx,
            double axisXy,
            double axisXz,
            double axisYx,
            double axisYy,
            double axisYz,
            double axisZx,
            double axisZy,
            double axisZz,
            out double localMinX,
            out double localMinY,
            out double localMinZ,
            out double localMaxX,
            out double localMaxY,
            out double localMaxZ)
        {
            localMinX = 0.0;
            localMinY = 0.0;
            localMinZ = 0.0;
            localMaxX = 0.0;
            localMaxY = 0.0;
            localMaxZ = 0.0;

            double[] xs = { minX, maxX };
            double[] ys = { minY, maxY };
            double[] zs = { minZ, maxZ };

            bool initialized = false;
            for (int ix = 0; ix < xs.Length; ix++)
            {
                for (int iy = 0; iy < ys.Length; iy++)
                {
                    for (int iz = 0; iz < zs.Length; iz++)
                    {
                        double dx = xs[ix] - originX;
                        double dy = ys[iy] - originY;
                        double dz = zs[iz] - originZ;

                        double lx = Dot3D(dx, dy, dz, axisXx, axisXy, axisXz);
                        double ly = Dot3D(dx, dy, dz, axisYx, axisYy, axisYz);
                        double lz = Dot3D(dx, dy, dz, axisZx, axisZy, axisZz);

                        if (!initialized)
                        {
                            localMinX = lx;
                            localMinY = ly;
                            localMinZ = lz;
                            localMaxX = lx;
                            localMaxY = ly;
                            localMaxZ = lz;
                            initialized = true;
                        }
                        else
                        {
                            localMinX = Math.Min(localMinX, lx);
                            localMinY = Math.Min(localMinY, ly);
                            localMinZ = Math.Min(localMinZ, lz);
                            localMaxX = Math.Max(localMaxX, lx);
                            localMaxY = Math.Max(localMaxY, ly);
                            localMaxZ = Math.Max(localMaxZ, lz);
                        }
                    }
                }
            }

            return initialized;
        }

        private static List<OrbitToolForm.PreviewTriangle> TryBuildPreviewTrianglesFromModel(
            object model,
            object identifier,
            bool useLocalCoordinates,
            double originX,
            double originY,
            double originZ,
            double axisXx,
            double axisXy,
            double axisXz,
            double axisYx,
            double axisYy,
            double axisYz,
            double axisZx,
            double axisZy,
            double axisZz)
        {
            List<OrbitToolForm.PreviewTriangle> triangles = new List<OrbitToolForm.PreviewTriangle>();
            if (model == null || identifier == null)
            {
                return triangles;
            }

            object modelObject = InvokeMethod(model, "SelectModelObject", identifier);
            if (modelObject == null)
            {
                return triangles;
            }

            object solid = InvokeMethod(modelObject, "GetSolid");
            if (solid == null)
            {
                return triangles;
            }

            List<PreviewVertex> rawTriangles;
            if (!TryExtractSolidTrianglesFromSolid(solid, out rawTriangles) || rawTriangles.Count == 0)
            {
                return triangles;
            }

            for (int i = 0; i + 2 < rawTriangles.Count; i += 3)
            {
                PreviewVertex v1 = rawTriangles[i];
                PreviewVertex v2 = rawTriangles[i + 1];
                PreviewVertex v3 = rawTriangles[i + 2];

                if (useLocalCoordinates)
                {
                    v1 = TransformToLocalPreview(v1, originX, originY, originZ, axisXx, axisXy, axisXz, axisYx, axisYy, axisYz, axisZx, axisZy, axisZz);
                    v2 = TransformToLocalPreview(v2, originX, originY, originZ, axisXx, axisXy, axisXz, axisYx, axisYy, axisYz, axisZx, axisZy, axisZz);
                    v3 = TransformToLocalPreview(v3, originX, originY, originZ, axisXx, axisXy, axisXz, axisYx, axisYy, axisYz, axisZx, axisZy, axisZz);
                }

                if (ComputeTriangleAreaSquared(v1, v2, v3) < 1e-10)
                {
                    continue;
                }

                triangles.Add(
                    new OrbitToolForm.PreviewTriangle(
                        v1.X, v1.Y, v1.Z,
                        v2.X, v2.Y, v2.Z,
                        v3.X, v3.Y, v3.Z));
            }

            return triangles;
        }

        private static bool TryExtractSolidTrianglesFromSolid(
            object solid,
            out List<PreviewVertex> triangleVertices)
        {
            triangleVertices = new List<PreviewVertex>();
            if (solid == null)
            {
                return false;
            }

            object faceEnumerator = InvokeParameterlessMethod(solid, "GetFaceEnumerator")
                ?? InvokeParameterlessMethod(solid, "GetFaces");
            if (faceEnumerator == null)
            {
                return false;
            }

            List<object> faces = EnumerateEnumeratorItems(faceEnumerator);
            if (faces.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < faces.Count; i++)
            {
                object face = faces[i];
                if (face == null)
                {
                    continue;
                }

                List<PreviewVertex> vertices = new List<PreviewVertex>();
                object loopEnumerator = InvokeParameterlessMethod(face, "GetLoopEnumerator")
                    ?? InvokeParameterlessMethod(face, "GetLoops");

                if (loopEnumerator != null)
                {
                    List<object> loops = EnumerateEnumeratorItems(loopEnumerator);
                    for (int loopIndex = 0; loopIndex < loops.Count; loopIndex++)
                    {
                        object loop = loops[loopIndex];
                        if (TryExtractLoopVertices(loop, out vertices) && vertices.Count >= 3)
                        {
                            break;
                        }
                    }
                }

                if (vertices.Count < 3)
                {
                    object vertexEnumerator = InvokeParameterlessMethod(face, "GetVertexEnumerator")
                        ?? InvokeParameterlessMethod(face, "GetVertexes")
                        ?? InvokeParameterlessMethod(face, "GetVertices");
                    if (vertexEnumerator != null)
                    {
                        List<PreviewVertex> faceVertices;
                        if (TryExtractVertexEnumeratorVertices(vertexEnumerator, out faceVertices))
                        {
                            vertices = faceVertices;
                        }
                    }
                }

                if (vertices.Count < 3)
                {
                    continue;
                }

                PreviewVertex first = vertices[0];
                for (int v = 1; v < vertices.Count - 1; v++)
                {
                    triangleVertices.Add(first);
                    triangleVertices.Add(vertices[v]);
                    triangleVertices.Add(vertices[v + 1]);
                }
            }

            return triangleVertices.Count >= 3;
        }

        private static bool TryExtractLoopVertices(object loop, out List<PreviewVertex> vertices)
        {
            vertices = new List<PreviewVertex>();
            if (loop == null)
            {
                return false;
            }

            object vertexEnumerator = InvokeParameterlessMethod(loop, "GetVertexEnumerator")
                ?? InvokeParameterlessMethod(loop, "GetVertexes")
                ?? InvokeParameterlessMethod(loop, "GetVertices");
            if (vertexEnumerator == null)
            {
                return false;
            }

            return TryExtractVertexEnumeratorVertices(vertexEnumerator, out vertices);
        }

        private static bool TryExtractVertexEnumeratorVertices(
            object vertexEnumerator,
            out List<PreviewVertex> vertices)
        {
            vertices = new List<PreviewVertex>();
            if (vertexEnumerator == null)
            {
                return false;
            }

            List<object> points = EnumerateEnumeratorItems(vertexEnumerator);
            for (int i = 0; i < points.Count; i++)
            {
                double x;
                double y;
                double z;
                if (!TryGetXYZ(points[i], out x, out y, out z))
                {
                    continue;
                }

                vertices.Add(new PreviewVertex(x, y, z));
            }

            RemoveConsecutiveDuplicateVertices(vertices);
            if (vertices.Count >= 2)
            {
                PreviewVertex first = vertices[0];
                PreviewVertex last = vertices[vertices.Count - 1];
                if (AreVerticesNear(first, last))
                {
                    vertices.RemoveAt(vertices.Count - 1);
                }
            }

            return vertices.Count >= 3;
        }

        private static List<object> EnumerateEnumeratorItems(object enumerator)
        {
            List<object> items = new List<object>();
            if (enumerator == null)
            {
                return items;
            }

            MethodInfo moveNext = enumerator.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = enumerator.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return items;
            }

            while (true)
            {
                object moved = moveNext.Invoke(enumerator, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object item = current.GetValue(enumerator, null);
                if (item != null)
                {
                    items.Add(item);
                }
            }

            return items;
        }

        private static void RemoveConsecutiveDuplicateVertices(List<PreviewVertex> vertices)
        {
            if (vertices == null || vertices.Count < 2)
            {
                return;
            }

            for (int i = vertices.Count - 1; i > 0; i--)
            {
                if (AreVerticesNear(vertices[i], vertices[i - 1]))
                {
                    vertices.RemoveAt(i);
                }
            }
        }

        private static bool AreVerticesNear(PreviewVertex a, PreviewVertex b)
        {
            const double tolerance = 1e-7;
            return Math.Abs(a.X - b.X) <= tolerance
                && Math.Abs(a.Y - b.Y) <= tolerance
                && Math.Abs(a.Z - b.Z) <= tolerance;
        }

        private static PreviewVertex TransformToLocalPreview(
            PreviewVertex world,
            double originX,
            double originY,
            double originZ,
            double axisXx,
            double axisXy,
            double axisXz,
            double axisYx,
            double axisYy,
            double axisYz,
            double axisZx,
            double axisZy,
            double axisZz)
        {
            double dx = world.X - originX;
            double dy = world.Y - originY;
            double dz = world.Z - originZ;

            return new PreviewVertex(
                Dot3D(dx, dy, dz, axisXx, axisXy, axisXz),
                Dot3D(dx, dy, dz, axisYx, axisYy, axisYz),
                Dot3D(dx, dy, dz, axisZx, axisZy, axisZz));
        }

        private static double ComputeTriangleAreaSquared(PreviewVertex v1, PreviewVertex v2, PreviewVertex v3)
        {
            double ux = v2.X - v1.X;
            double uy = v2.Y - v1.Y;
            double uz = v2.Z - v1.Z;
            double vx = v3.X - v1.X;
            double vy = v3.Y - v1.Y;
            double vz = v3.Z - v1.Z;
            double cx = (uy * vz) - (uz * vy);
            double cy = (uz * vx) - (ux * vz);
            double cz = (ux * vy) - (uy * vx);
            return (cx * cx) + (cy * cy) + (cz * cz);
        }

        private struct PreviewVertex
        {
            public PreviewVertex(double x, double y, double z)
            {
                X = x;
                Y = y;
                Z = z;
            }

            public double X;
            public double Y;
            public double Z;
        }

        private static List<ExplodedPartPlan> CollectPartPlansFromDrawing(
            object activeDrawing,
            out int inspectedViews,
            out int candidateViews)
        {
            inspectedViews = 0;
            candidateViews = 0;

            List<ExplodedPartPlan> contextPlans = CollectPartPlansFromDrawingContext(activeDrawing);
            if (contextPlans.Count > 0)
            {
                return contextPlans;
            }

            List<ExplodedPartPlan> mergedPlans = new List<ExplodedPartPlan>();
            if (activeDrawing == null)
            {
                return mergedPlans;
            }

            object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
            if (sheet == null)
            {
                return mergedPlans;
            }

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects")
                ?? InvokeParameterlessMethod(sheet, "GetObjects");
            if (allObjects == null)
            {
                return mergedPlans;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return mergedPlans;
            }

            HashSet<string> seenPartIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

                inspectedViews++;

                string viewName = Convert.ToString(GetPropertyValue(drawingObject, "Name")) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(viewName)
                    && viewName.StartsWith(ExplodedViewPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                List<ExplodedPartPlan> viewPlans = CollectPartPlansFromView(drawingObject);
                if (viewPlans == null || viewPlans.Count == 0)
                {
                    continue;
                }

                candidateViews++;

                for (int i = 0; i < viewPlans.Count; i++)
                {
                    ExplodedPartPlan plan = viewPlans[i];
                    if (plan == null)
                    {
                        continue;
                    }

                    string key = !string.IsNullOrWhiteSpace(plan.IdentifierKey)
                        ? plan.IdentifierKey
                        : GetIdentifierKey(plan.Identifier);
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        continue;
                    }

                    if (seenPartIds.Add(key))
                    {
                        mergedPlans.Add(plan);
                    }
                }
            }

            return mergedPlans;
        }

        private static List<ExplodedPartPlan> CollectPartPlansFromDrawingContext(object activeDrawing)
        {
            List<ExplodedPartPlan> plans = new List<ExplodedPartPlan>();
            if (activeDrawing == null)
            {
                return plans;
            }

            object model = CreateModelInstance();
            if (model == null)
            {
                return plans;
            }

            bool? connected = InvokeBoolMethod(model, "GetConnectionStatus");
            if (!connected.HasValue || !connected.Value)
            {
                return plans;
            }

            string[] contextMembers =
            {
                "GetAssembly",
                "GetPart",
                "GetCastUnit",
                "GetModelObject",
                "GetAssemblyIdentifier",
                "GetPartIdentifier",
                "GetCastUnitIdentifier",
                "GetModelIdentifier",
                "Assembly",
                "Part",
                "CastUnit",
                "ModelObject",
                "AssemblyIdentifier",
                "PartIdentifier",
                "CastUnitIdentifier",
                "ModelIdentifier"
            };

            List<object> contextValues = new List<object>();
            for (int i = 0; i < contextMembers.Length; i++)
            {
                string member = contextMembers[i];
                AddContextCandidateValue(contextValues, InvokeParameterlessMethod(activeDrawing, member));
                AddContextCandidateValue(contextValues, GetPropertyValue(activeDrawing, member));
            }

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < contextValues.Count; i++)
            {
                AppendPartPlansFromDrawingContextValue(model, contextValues[i], seen, plans);
            }

            return plans;
        }

        private static void AppendPartPlansFromDrawingContextValue(
            object model,
            object contextValue,
            HashSet<string> seen,
            List<ExplodedPartPlan> plans)
        {
            if (model == null || contextValue == null || seen == null || plans == null)
            {
                return;
            }

            List<object> values = new List<object>();
            CollectItemsFromUnknownCollection(contextValue, values);
            if (values.Count == 0)
            {
                values.Add(contextValue);
            }

            for (int i = 0; i < values.Count; i++)
            {
                object modelObject = TryResolveModelObjectFromContextValue(model, values[i]);
                if (modelObject == null)
                {
                    continue;
                }

                bool addedFromAssembly = false;
                addedFromAssembly |= TryAddPartPlanFromModelObject(modelObject, seen, plans);
                addedFromAssembly |= TryAddPartPlanFromModelObject(
                    InvokeParameterlessMethod(modelObject, "GetMainPart"),
                    seen,
                    plans);
                addedFromAssembly |= TryAddPartPlansFromModelCollection(
                    InvokeParameterlessMethod(modelObject, "GetSecondaries"),
                    seen,
                    plans);
                addedFromAssembly |= TryAddPartPlansFromModelCollection(
                    InvokeParameterlessMethod(modelObject, "GetParts"),
                    seen,
                    plans);

                object assemblyObject = InvokeParameterlessMethod(modelObject, "GetAssembly");
                if (assemblyObject != null && !ReferenceEquals(assemblyObject, modelObject))
                {
                    addedFromAssembly |= TryAddPartPlanFromModelObject(
                        InvokeParameterlessMethod(assemblyObject, "GetMainPart"),
                        seen,
                        plans);
                    addedFromAssembly |= TryAddPartPlansFromModelCollection(
                        InvokeParameterlessMethod(assemblyObject, "GetSecondaries"),
                        seen,
                        plans);
                }
            }
        }

        private static bool TryAddPartPlansFromModelCollection(
            object collection,
            HashSet<string> seen,
            List<ExplodedPartPlan> plans)
        {
            bool added = false;
            List<object> values = new List<object>();
            CollectItemsFromUnknownCollection(collection, values);
            for (int i = 0; i < values.Count; i++)
            {
                if (TryAddPartPlanFromModelObject(values[i], seen, plans))
                {
                    added = true;
                }
            }

            return added;
        }

        private static bool TryAddPartPlanFromModelObject(
            object modelObject,
            HashSet<string> seen,
            List<ExplodedPartPlan> plans)
        {
            if (modelObject == null || seen == null || plans == null)
            {
                return false;
            }

            Type type = modelObject.GetType();
            string typeName = type != null ? type.Name : string.Empty;
            if (!IsModelPartLikeTypeName(typeName))
            {
                return false;
            }

            object identifier = GetPropertyValue(modelObject, "Identifier");
            return TryAddPartPlanFromIdentifier(identifier, seen, plans);
        }

        private static bool TryAddPartPlanFromIdentifier(
            object identifier,
            HashSet<string> seen,
            List<ExplodedPartPlan> plans)
        {
            if (identifier == null || seen == null || plans == null)
            {
                return false;
            }

            string key = GetIdentifierKey(identifier);
            if (string.IsNullOrWhiteSpace(key) || !seen.Add(key))
            {
                return false;
            }

            ExplodedPartPlan plan = new ExplodedPartPlan();
            plan.Identifier = identifier;
            plan.IdentifierKey = key;
            plans.Add(plan);
            return true;
        }

        private static object TryResolveModelObjectFromContextValue(object model, object contextValue)
        {
            if (model == null || contextValue == null)
            {
                return null;
            }

            Type valueType = contextValue.GetType();
            string fullName = valueType != null ? valueType.FullName : string.Empty;
            if (!string.IsNullOrWhiteSpace(fullName)
                && fullName.StartsWith("Tekla.Structures.Model.", StringComparison.OrdinalIgnoreCase))
            {
                return contextValue;
            }

            object idValue = GetPropertyValue(contextValue, "ID") ?? GetPropertyValue(contextValue, "Id");
            if (idValue == null)
            {
                return null;
            }

            return InvokeMethod(model, "SelectModelObject", contextValue);
        }

        private static void CollectItemsFromUnknownCollection(object collection, List<object> values)
        {
            if (collection == null || values == null || collection is string)
            {
                return;
            }

            IEnumerable enumerable = collection as IEnumerable;
            if (enumerable != null)
            {
                foreach (object item in enumerable)
                {
                    if (item != null)
                    {
                        values.Add(item);
                    }
                }

                return;
            }

            MethodInfo moveNext = collection.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = collection.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return;
            }

            while (true)
            {
                object moved = moveNext.Invoke(collection, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object item = current.GetValue(collection, null);
                if (item != null)
                {
                    values.Add(item);
                }
            }
        }

        private static void AddContextCandidateValue(List<object> values, object candidate)
        {
            if (values == null || candidate == null)
            {
                return;
            }

            List<object> expanded = new List<object>();
            CollectItemsFromUnknownCollection(candidate, expanded);
            if (expanded.Count > 0)
            {
                for (int i = 0; i < expanded.Count; i++)
                {
                    if (expanded[i] != null && !values.Contains(expanded[i]))
                    {
                        values.Add(expanded[i]);
                    }
                }

                return;
            }

            if (!values.Contains(candidate))
            {
                values.Add(candidate);
            }
        }

        private static bool IsModelPartLikeTypeName(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return false;
            }

            string normalized = typeName.Trim().ToLowerInvariant();
            return normalized.Contains("part")
                || normalized.Contains("beam")
                || normalized.Contains("plate")
                || normalized.Contains("polybeam")
                || normalized.Contains("bentplate")
                || normalized.Contains("item");
        }

        private void RefreshTeklaStatus(bool updateOutput)
        {
            try
            {
                Type modelType = ResolveTeklaModelType();
                if (modelType == null)
                {
                    SetIndicatorColor(pnlTeklaIndicator, Color.IndianRed);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: API do Tekla nao encontrada.", Color.DarkRed);
                        SetOutput("Nao foi possivel carregar Tekla.Structures.Model.Model.");
                    }
                    return;
                }

                object modelInstance = Activator.CreateInstance(modelType);
                bool? connected = InvokeBoolMethod(modelInstance, "GetConnectionStatus");
                if (!connected.HasValue)
                {
                    SetIndicatorColor(pnlTeklaIndicator, Color.IndianRed);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: metodo GetConnectionStatus indisponivel.", Color.DarkRed);
                        SetOutput("Nao foi possivel ler o status da conexao de modelo.");
                    }
                    return;
                }

                if (connected.Value)
                {
                    SetIndicatorColor(pnlTeklaIndicator, Color.ForestGreen);
                    if (updateOutput)
                    {
                        UpdateStatus("Sucesso: comunicacao com Tekla estabelecida.", Color.DarkGreen);
                    }
                }
                else
                {
                    SetIndicatorColor(pnlTeklaIndicator, Color.DarkOrange);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: Tekla aberto, mas sem conexao com o modelo.", Color.DarkOrange);
                    }
                }

                if (updateOutput)
                {
                    StringBuilder report = new StringBuilder();
                    report.AppendLine("Conexao com modelo Tekla");
                    report.AppendLine("------------------------");
                    report.AppendLine("Tipo carregado: " + modelType.FullName);
                    report.AppendLine("GetConnectionStatus: " + (connected.Value ? "True" : "False"));
                    SetOutput(report.ToString());
                }
            }
            catch (Exception ex)
            {
                SetIndicatorColor(pnlTeklaIndicator, Color.IndianRed);
                if (updateOutput)
                {
                    UpdateStatus("Erro ao comunicar com Tekla: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
                }
            }
        }

        private void RefreshDrawingStatus(bool updateOutput)
        {
            try
            {
                object drawingHandler = CreateDrawingHandler();
                if (drawingHandler == null)
                {
                    SetIndicatorColor(pnlDrawingIndicator, Color.IndianRed);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: API de drawing do Tekla nao encontrada.", Color.DarkRed);
                        SetOutput("Nao foi possivel carregar Tekla.Structures.Drawing.DrawingHandler.");
                    }
                    return;
                }

                bool? drawingConnection = InvokeBoolMethod(drawingHandler, "GetConnectionStatus");
                if (drawingConnection.HasValue && !drawingConnection.Value)
                {
                    SetIndicatorColor(pnlDrawingIndicator, Color.DarkOrange);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: sem conexao com Drawing API.", Color.DarkOrange);
                        SetOutput("DrawingHandler.GetConnectionStatus retornou False.");
                    }
                    return;
                }

                object activeDrawing = InvokeParameterlessMethod(drawingHandler, "GetActiveDrawing");
                if (activeDrawing == null)
                {
                    SetIndicatorColor(pnlDrawingIndicator, Color.DarkOrange);
                    if (updateOutput)
                    {
                        UpdateStatus("Falha: nenhum desenho aberto no Tekla.", Color.DarkOrange);
                        SetOutput("Nenhum desenho ativo encontrado no Tekla.");
                    }
                    return;
                }

                SetIndicatorColor(pnlDrawingIndicator, Color.ForestGreen);
                if (updateOutput)
                {
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
            }
            catch (Exception ex)
            {
                SetIndicatorColor(pnlDrawingIndicator, Color.IndianRed);
                if (updateOutput)
                {
                    UpdateStatus("Erro ao ler desenho: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
                }
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
            bool ghostLineEnabled = chkGhostLinhas.Checked;
            bool colorirEnabled = chkColorir.Checked;
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
                if (ghostLineEnabled)
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
                int removedGuideLines = DeleteExistingGuideLines(sheet);

                double sourceScale = GetSourceViewScale(sourceView);
                object model = CreateModelInstance();
                int centersFromModel;
                bool modelLayoutUsed = ComputeExplodedOffsets(sourceView, partPlans, sourceScale, model, out centersFromModel);

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

                    if (ghostLineEnabled && !plan.IsMainPart)
                    {
                        ghostPairsRequested++;

                        bool originalAutoIsoApplied;
                        object originalView = CreateSinglePartView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScale,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_ORG",
                            out originalAutoIsoApplied);
                        viewSequence++;

                        bool originalCreated = false;
                        if (originalView != null)
                        {
                            ApplyFallbackRotation(originalView, originalAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    originalView,
                                    anchorX + plan.OriginalOffsetX,
                                    anchorY + plan.OriginalOffsetY,
                                    anchorZ))
                            {
                                TryApplyGhostStyleToView(originalView);
                                if (colorirEnabled)
                                {
                                    int originalContourColor = GetContourColorForPlan(plan);
                                    if (originalContourColor > 0 && TryApplyContourColorToView(originalView, originalContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (originalAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostOriginalCreated++;
                                originalCreated = true;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        bool explodedAutoIsoApplied;
                        object explodedView = CreateSinglePartView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScale,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_EXP",
                            out explodedAutoIsoApplied);
                        viewSequence++;

                        bool explodedCreated = false;
                        if (explodedView != null)
                        {
                            ApplyFallbackRotation(explodedView, explodedAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    explodedView,
                                    anchorX + plan.OffsetX,
                                    anchorY + plan.OffsetY,
                                    anchorZ))
                            {
                                if (colorirEnabled)
                                {
                                    int explodedContourColor = GetContourColorForPlan(plan);
                                    if (explodedContourColor > 0 && TryApplyContourColorToView(explodedView, explodedContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (explodedAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostExplodedCreated++;
                                explodedCreated = true;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        guideLinesRequested++;
                        if (originalCreated
                            && explodedCreated
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

                        continue;
                    }

                    string viewName = ExplodedViewPrefix + viewSequence.ToString("D3");
                    viewSequence++;
                    bool autoIsoApplied;

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
                        failed++;
                        continue;
                    }

                    if (autoIsoApplied)
                    {
                        autoIsoAppliedCount++;
                    }

                    ApplyFallbackRotation(newView, autoIsoApplied, ref fallbackRotationCount);

                    bool positioned = ForceViewCenterToTarget(
                        newView,
                        anchorX + plan.OffsetX,
                        anchorY + plan.OffsetY,
                        anchorZ);
                    if (!positioned)
                    {
                        failed++;
                        continue;
                    }

                    if (colorirEnabled)
                    {
                        int contourColor = GetContourColorForPlan(plan);
                        if (contourColor > 0 && TryApplyContourColorToView(newView, contourColor))
                        {
                            contourColoredCount++;
                        }
                        else
                        {
                            contourNotChangedCount++;
                        }
                    }

                    created++;
                }

                CommitActiveDrawingChanges(activeDrawing);
                InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");

                StringBuilder report = new StringBuilder();
                report.AppendLine("Vista explodida (somente drawing)");
                report.AppendLine("---------------------------------");
                report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                report.AppendLine("Partes candidatas na vista: " + candidateCount + (truncated ? " (limitado para " + MaxExplodedViews + ")" : string.Empty));
                report.AppendLine("Vistas explodidas antigas removidas: " + removedExisting);
                report.AppendLine("Base de posicionamento: " + (modelLayoutUsed ? "modelo (posicao relativa)" : "drawing (fallback)"));
                report.AppendLine("Anchor de posicionamento: " + (anchorFromMainPart ? "centro da peca principal" : "origem da vista"));
                report.AppendLine("Centros lidos do modelo: " + centersFromModel);
                report.AppendLine("Perfil de vista solicitado: " + ExplodedViewAttributeFile);
                report.AppendLine("Vistas com perfil aplicado: " + autoIsoAppliedCount);
                report.AppendLine("Rotacao isometrica de fallback: " + fallbackRotationCount);
                report.AppendLine("Colorir: " + (colorirEnabled ? "ligado" : "desligado"));
                report.AppendLine("Cor de contorno aplicada: " + contourColoredCount);
                report.AppendLine("Sem mudanca de cor: " + contourNotChangedCount);
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
                int removedGuideLines = DeleteExistingGuideLines(sheet);
                CommitActiveDrawingChanges(activeDrawing);

                StringBuilder report = new StringBuilder();
                report.AppendLine("Limpeza de vista explodida");
                report.AppendLine("--------------------------");
                report.AppendLine("Prefixo: " + ExplodedViewPrefix);
                report.AppendLine("Vistas removidas: " + removed);
                report.AppendLine("Linhas guia removidas: " + removedGuideLines);
                report.AppendLine("CommitChanges executado: sim");
                report.AppendLine();
                report.AppendLine("Observacao: somente objetos de desenho foram alterados.");
                SetOutput(report.ToString());

                if ((removed + removedGuideLines) > 0)
                {
                    UpdateStatus("Sucesso: vistas explodidas e linhas guia removidas.", Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Concluido: nenhuma vista explodida/linha guia para remover.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao limpar vista explodida: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
        }

        private void btnExplodirPorPlanos_Click(object sender, EventArgs e)
        {
            bool ghostLineEnabled = chkGhostLinhas.Checked;
            bool colorirEnabled = chkColorir.Checked;
            bool usePlaneZY = chkXPositivo.Checked || chkXNegativo.Checked;
            bool usePlaneXZ = chkYPositivo.Checked || chkYNegativo.Checked;
            bool usePlaneXY = chkZPositivo.Checked || chkZNegativo.Checked;

            if (!usePlaneXY && !usePlaneXZ && !usePlaneZY)
            {
                UpdateStatus("Falha: selecione pelo menos um sentido em X+/X-/Y+/Y-/Z+/Z-.", Color.DarkOrange);
                SetOutput("Marque ao menos um checkbox em X+/X-/Y+/Y-/Z+/Z- antes de explodir.");
                return;
            }

            string selectedPlanes = string.Empty;
            if (usePlaneXY)
            {
                selectedPlanes += "xy ";
            }

            if (usePlaneXZ)
            {
                selectedPlanes += "xz ";
            }

            if (usePlaneZY)
            {
                selectedPlanes += "zy ";
            }

            selectedPlanes = selectedPlanes.Trim();

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
                if (ghostLineEnabled)
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
                int removedGuideLines = DeleteExistingGuideLines(sheet);

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
                    usePlaneXY,
                    usePlaneXZ,
                    usePlaneZY,
                    false,
                    0.0,
                    0.0,
                    0.0,
                    0.0,
                    0.0,
                    0.0,
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

                    if (ghostLineEnabled && !plan.IsMainPart)
                    {
                        ghostPairsRequested++;

                        bool originalAutoIsoApplied;
                        object originalView = CreateSinglePartView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScale,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_ORG",
                            out originalAutoIsoApplied);
                        viewSequence++;

                        bool originalCreated = false;
                        if (originalView != null)
                        {
                            ApplyFallbackRotation(originalView, originalAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    originalView,
                                    anchorX + plan.OriginalOffsetX,
                                    anchorY + plan.OriginalOffsetY,
                                    anchorZ))
                            {
                                TryApplyGhostStyleToView(originalView);
                                if (colorirEnabled)
                                {
                                    int originalContourColor = GetContourColorForPlan(plan);
                                    if (originalContourColor > 0 && TryApplyContourColorToView(originalView, originalContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (originalAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostOriginalCreated++;
                                originalCreated = true;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        bool explodedAutoIsoApplied;
                        object explodedView = CreateSinglePartView(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScale,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_EXP",
                            out explodedAutoIsoApplied);
                        viewSequence++;

                        bool explodedCreated = false;
                        if (explodedView != null)
                        {
                            ApplyFallbackRotation(explodedView, explodedAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    explodedView,
                                    anchorX + plan.OffsetX,
                                    anchorY + plan.OffsetY,
                                    anchorZ))
                            {
                                if (colorirEnabled)
                                {
                                    int explodedContourColor = GetContourColorForPlan(plan);
                                    if (explodedContourColor > 0 && TryApplyContourColorToView(explodedView, explodedContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (explodedAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostExplodedCreated++;
                                explodedCreated = true;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        guideLinesRequested++;
                        if (originalCreated
                            && explodedCreated
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

                        continue;
                    }

                    string viewName = ExplodedViewPrefix + viewSequence.ToString("D3");
                    viewSequence++;
                    bool autoIsoApplied;

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
                        failed++;
                        continue;
                    }

                    if (autoIsoApplied)
                    {
                        autoIsoAppliedCount++;
                    }

                    ApplyFallbackRotation(newView, autoIsoApplied, ref fallbackRotationCount);

                    bool positioned = ForceViewCenterToTarget(
                        newView,
                        anchorX + plan.OffsetX,
                        anchorY + plan.OffsetY,
                        anchorZ);
                    if (!positioned)
                    {
                        failed++;
                        continue;
                    }

                    if (colorirEnabled)
                    {
                        int contourColor = GetContourColorForPlan(plan);
                        if (contourColor > 0 && TryApplyContourColorToView(newView, contourColor))
                        {
                            contourColoredCount++;
                        }
                        else
                        {
                            contourNotChangedCount++;
                        }
                    }

                    created++;
                }

                CommitActiveDrawingChanges(activeDrawing);
                InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");

                StringBuilder report = new StringBuilder();
                report.AppendLine("Vista explodida por planos (somente drawing)");
                report.AppendLine("--------------------------------------------");
                report.AppendLine("Planos ativos: " + selectedPlanes);
                report.AppendLine("Ghost + linha: " + (ghostLineEnabled ? "ligado" : "desligado"));
                report.AppendLine("Colorir: " + (colorirEnabled ? "ligado" : "desligado"));
                report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                report.AppendLine("Partes candidatas na vista: " + candidateCount + (truncated ? " (limitado para " + maxPartsAllowed + ")" : string.Empty));
                report.AppendLine("Vistas explodidas antigas removidas: " + removedExisting);
                report.AppendLine("Linhas guia antigas removidas: " + removedGuideLines);
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
                if (ghostLineEnabled)
                {
                    report.AppendLine("Ghost pares secundarios: " + ghostPairsRequested);
                    report.AppendLine("Ghost vistas originais criadas: " + ghostOriginalCreated);
                    report.AppendLine("Ghost vistas deslocadas criadas: " + ghostExplodedCreated);
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

        private void btnCriarFolhaInteira_Click(object sender, EventArgs e)
        {
            CriarVistaExplodidaEmArea(false);
        }

        private void btnCriarAreaDefinida_Click(object sender, EventArgs e)
        {
            CriarVistaExplodidaEmArea(true);
        }

        private static string BuildSelectedAxesText(bool usePlaneXY, bool usePlaneXZ, bool usePlaneZY)
        {
            StringBuilder axes = new StringBuilder();
            if (usePlaneZY)
            {
                axes.Append("X ");
            }

            if (usePlaneXZ)
            {
                axes.Append("Y ");
            }

            if (usePlaneXY)
            {
                axes.Append("Z ");
            }

            return axes.ToString().Trim();
        }

        private static string BuildSelectedDirectionsText(
            bool allowXPositive,
            bool allowXNegative,
            bool allowYPositive,
            bool allowYNegative,
            bool allowZPositive,
            bool allowZNegative)
        {
            StringBuilder directions = new StringBuilder();
            if (allowXPositive)
            {
                directions.Append("X+ ");
            }

            if (allowXNegative)
            {
                directions.Append("X- ");
            }

            if (allowYPositive)
            {
                directions.Append("Y+ ");
            }

            if (allowYNegative)
            {
                directions.Append("Y- ");
            }

            if (allowZPositive)
            {
                directions.Append("Z+ ");
            }

            if (allowZNegative)
            {
                directions.Append("Z- ");
            }

            string result = directions.ToString().Trim();
            return string.IsNullOrWhiteSpace(result) ? "(nenhum)" : result;
        }

        private void CriarVistaExplodidaEmArea(bool useSelectedArea)
        {
            bool ghostEnabled = chkGhostLinhas.Checked;
            bool guideLinesEnabled = chkLinhas.Checked;
            bool colorirEnabled = chkColorir.Checked;
            bool testEnabled = chkTeste.Checked;
            bool allowXPositive = chkXPositivo.Checked;
            bool allowXNegative = chkXNegativo.Checked;
            bool allowYPositive = chkYPositivo.Checked;
            bool allowYNegative = chkYNegativo.Checked;
            bool allowZPositive = chkZPositivo.Checked;
            bool allowZNegative = chkZNegativo.Checked;
            bool usePlaneZY = allowXPositive || allowXNegative;
            bool usePlaneXZ = allowYPositive || allowYNegative;
            bool usePlaneXY = allowZPositive || allowZNegative;

            if (!usePlaneXY && !usePlaneXZ && !usePlaneZY)
            {
                UpdateStatus("Falha: selecione pelo menos um eixo ou sentido.", Color.DarkOrange);
                SetOutput("Marque ao menos um checkbox em X+/X-/Y+/Y-/Z+/Z- antes de criar a vista ajustada.");
                return;
            }

            string selectedAxes = BuildSelectedAxesText(usePlaneXY, usePlaneXZ, usePlaneZY);

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
                    SetOutput("Abra um desenho no Tekla antes de criar a vista ajustada.");
                    return;
                }

                object sheet = InvokeParameterlessMethod(activeDrawing, "GetSheet");
                if (sheet == null)
                {
                    UpdateStatus("Falha: nao foi possivel ler a folha do desenho.", Color.DarkRed);
                    SetOutput("Drawing.GetSheet retornou nulo.");
                    return;
                }

                double rectMinX;
                double rectMinY;
                double rectMaxX;
                double rectMaxY;
                int selectedObjectCount = 0;
                int boundedSelectedCount = 0;
                string targetAreaSource = useSelectedArea ? "2 pontos escolhidos" : "folha inteira";
                bool pickerCancelled = false;
                bool hasTargetRectangle = useSelectedArea
                    ? TryPickPlacementRectangleByTwoPoints(
                        activeDrawing,
                        sheet,
                        out rectMinX,
                        out rectMinY,
                        out rectMaxX,
                        out rectMaxY,
                        out pickerCancelled)
                    : TryGetSheetPlacementRectangle(
                        sheet,
                        out rectMinX,
                        out rectMinY,
                        out rectMaxX,
                        out rectMaxY);

                if (!hasTargetRectangle)
                {
                    if (useSelectedArea)
                    {
                        if (pickerCancelled)
                        {
                            UpdateStatus("Cancelado: selecao da area interrompida.", Color.DarkOrange);
                            SetOutput("A selecao por 2 pontos foi cancelada no Tekla.");
                        }
                        else
                        {
                            UpdateStatus("Falha: area definida nao encontrada.", Color.DarkOrange);
                            SetOutput("Nao foi possivel capturar os 2 pontos da area no Tekla.");
                        }
                    }
                    else
                    {
                        UpdateStatus("Falha: nao foi possivel ler a area da folha.", Color.DarkOrange);
                        SetOutput("Nao foi possivel obter os limites da folha para usar o desenho completo.");
                    }
                    return;
                }

                object sourceView = null;
                int inspectedViews;
                int candidateViews;
                string sourceViewName = "conjunto do desenho ativo";
                double sourceIsoScore = 0.0;

                List<ExplodedPartPlan> partPlans = CollectPartPlansFromDrawing(activeDrawing, out inspectedViews, out candidateViews);
                if (partPlans.Count == 0)
                {
                    UpdateStatus("Falha: nao foi possivel ler as pecas do conjunto no desenho ativo.", Color.DarkOrange);
                    SetOutput("Nenhum ModelIdentifier de parte foi encontrado no contexto do desenho ativo.");
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
                int removedGuideLines = DeleteExistingGuideLines(sheet);

                double sourceScale = previewCachedSourceScale > 0.0 ? previewCachedSourceScale : 1.0;
                object model = CreateModelInstance();
                int centersFromModel;
                int movedByXY;
                int movedByXZ;
                int movedByZY;
                int blockedByDirection = 0;
                bool usePreviewAxesForExplode = false;
                double previewAxisXx = 0.0;
                double previewAxisXy = 0.0;
                double previewAxisXz = 0.0;
                double previewAxisYx = 0.0;
                double previewAxisYy = 0.0;
                double previewAxisYz = 0.0;

                if (chkRotacaoPreview != null && chkRotacaoPreview.Checked)
                {
                    if (TryGetPreviewDisplayAxesForView(
                            sourceView,
                            out previewAxisXx,
                            out previewAxisXy,
                            out previewAxisXz,
                            out previewAxisYx,
                            out previewAxisYy,
                            out previewAxisYz))
                    {
                        usePreviewAxesForExplode = true;
                    }
                }

                if (!usePreviewAxesForExplode && sourceView == null)
                {
                    if (!TryGetPreviewDisplayAxesForView(
                            null,
                            out previewAxisXx,
                            out previewAxisXy,
                            out previewAxisXz,
                            out previewAxisYx,
                            out previewAxisYy,
                            out previewAxisYz))
                    {
                        previewAxisXx = 1.0;
                        previewAxisXy = 0.0;
                        previewAxisXz = 0.0;
                        previewAxisYx = 0.0;
                        previewAxisYy = 1.0;
                        previewAxisYz = 0.0;
                    }

                    usePreviewAxesForExplode = true;
                }

                bool modelLayoutUsed = ComputeExplodedOffsetsByPlanes(
                    sourceView,
                    partPlans,
                    sourceScale,
                    model,
                    usePlaneXY,
                    usePlaneXZ,
                    usePlaneZY,
                    usePreviewAxesForExplode,
                    previewAxisXx,
                    previewAxisXy,
                    previewAxisXz,
                    previewAxisYx,
                    previewAxisYy,
                    previewAxisYz,
                    out centersFromModel,
                    out movedByXY,
                    out movedByXZ,
                    out movedByZY);

                bool fallbackLayoutUsed = false;
                if (!modelLayoutUsed)
                {
                    int fallbackCenters;
                    if (sourceView != null)
                    {
                        ComputeExplodedOffsets(sourceView, partPlans, sourceScale, model, out fallbackCenters);
                    }
                    else
                    {
                        ApplyRadialOffsets(null, partPlans, false);
                        fallbackCenters = 0;
                    }

                    centersFromModel = Math.Max(centersFromModel, fallbackCenters);
                    fallbackLayoutUsed = true;
                }

                ApplyDirectionalMovementFilters(
                    partPlans,
                    allowXPositive,
                    allowXNegative,
                    allowYPositive,
                    allowYNegative,
                    allowZPositive,
                    allowZNegative,
                    out blockedByDirection);

                double anchorZ;
                if (sourceView != null)
                {
                    double sourceOriginX;
                    double sourceOriginY;
                    object anchorOrigin = GetPropertyValue(sourceView, "Origin");
                    if (!TryGetXYZ(anchorOrigin, out sourceOriginX, out sourceOriginY, out anchorZ))
                    {
                        UpdateStatus("Falha: nao foi possivel ler origem da vista isometrica.", Color.DarkRed);
                        SetOutput("View.Origin da vista detectada nao esta disponivel.");
                        return;
                    }
                }
                else
                {
                    double sheetOriginX;
                    double sheetOriginY;
                    object sheetOrigin = GetPropertyValue(sheet, "Origin");
                    if (!TryGetXYZ(sheetOrigin, out sheetOriginX, out sheetOriginY, out anchorZ))
                    {
                        anchorZ = 0.0;
                    }
                }

                double layoutMinX;
                double layoutMinY;
                double layoutMaxX;
                double layoutMaxY;
                if (!TryGetPlannedLayoutBounds(
                    partPlans,
                    ghostEnabled,
                    1.0,
                    out layoutMinX,
                        out layoutMinY,
                        out layoutMaxX,
                        out layoutMaxY))
                {
                    UpdateStatus("Falha: nao foi possivel calcular o layout base da explosao.", Color.DarkRed);
                    SetOutput("Nao foi possivel calcular os limites base da vista explodida.");
                    return;
                }

                double layoutWidth = Math.Max(0.0, layoutMaxX - layoutMinX);
                double layoutHeight = Math.Max(0.0, layoutMaxY - layoutMinY);
                double fitScaleFactor;
                double anchorX;
                double anchorY;
                double fittedLayoutWidth;
                double fittedLayoutHeight;
                bool fitComputed = TryComputeFitIntoRectangle(
                    partPlans,
                    ghostEnabled,
                    rectMinX,
                    rectMinY,
                    rectMaxX,
                    rectMaxY,
                    !useSelectedArea,
                    out fitScaleFactor,
                    out anchorX,
                    out anchorY,
                    out fittedLayoutWidth,
                    out fittedLayoutHeight);
                if (!fitComputed)
                {
                    fitScaleFactor = 1.0;
                    fittedLayoutWidth = layoutWidth;
                    fittedLayoutHeight = layoutHeight;

                    double rectCenterX = (rectMinX + rectMaxX) * 0.5;
                    double rectCenterY = (rectMinY + rectMaxY) * 0.5;
                    double layoutCenterX = (layoutMinX + layoutMaxX) * 0.5;
                    double layoutCenterY = (layoutMinY + layoutMaxY) * 0.5;
                    anchorX = rectCenterX - layoutCenterX;
                    anchorY = rectCenterY - layoutCenterY;
                }
                else if (Math.Abs(fitScaleFactor - 1.0) > 1e-6)
                {
                    ScalePlannedOffsets(partPlans, fitScaleFactor);
                }

                double sourceScaleForFit = sourceScale > 0.0 ? sourceScale : 1.0;
                if (fitScaleFactor > 1e-6)
                {
                    sourceScaleForFit /= fitScaleFactor;
                }

                List<object> fitViews = new List<object>();

                object viewCoordinateSystem = sourceView != null
                    ? GetPropertyValue(sourceView, "ViewCoordinateSystem")
                    : null;
                object displayCoordinateSystem = sourceView != null
                    ? GetPropertyValue(sourceView, "DisplayCoordinateSystem")
                    : null;

                if (viewCoordinateSystem == null && previewCachedViewCoordinateSystem != null)
                {
                    viewCoordinateSystem = previewCachedViewCoordinateSystem;
                }

                if (displayCoordinateSystem == null && previewCachedDisplayCoordinateSystem != null)
                {
                    displayCoordinateSystem = previewCachedDisplayCoordinateSystem;
                }

                if (viewCoordinateSystem == null)
                {
                    viewCoordinateSystem = CreateGlobalCoordinateSystem();
                }

                if (displayCoordinateSystem == null)
                {
                    displayCoordinateSystem = CreateGlobalCoordinateSystem();
                }

                if (viewCoordinateSystem == null || displayCoordinateSystem == null)
                {
                    UpdateStatus("Falha: coordenadas base indisponiveis para criar as vistas.", Color.DarkRed);
                    SetOutput("Nao foi possivel obter ou criar ViewCoordinateSystem/DisplayCoordinateSystem.");
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
                int testAxisRequested = 0;
                int testAxisCreated = 0;
                int viewSequence = 1;

                for (int i = 0; i < partPlans.Count; i++)
                {
                    ExplodedPartPlan plan = partPlans[i];

                    if (ghostEnabled && !plan.IsMainPart && IsPlanActuallyDisplaced(plan))
                    {
                        ghostPairsRequested++;

                        bool originalAutoIsoApplied;
                        object originalView = CreateSinglePartViewForFit(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScaleForFit,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_ORG",
                            out originalAutoIsoApplied);
                        viewSequence++;

                        if (originalView != null)
                        {
                            ApplyFallbackRotation(originalView, originalAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    originalView,
                                    anchorX + plan.OriginalOffsetX,
                                    anchorY + plan.OriginalOffsetY,
                                    anchorZ))
                            {
                                TryApplyGhostStyleToView(originalView);
                                if (colorirEnabled)
                                {
                                    int originalContourColor = GetContourColorForPlan(plan);
                                    if (originalContourColor > 0 && TryApplyContourColorToView(originalView, originalContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (originalAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostOriginalCreated++;
                                fitViews.Add(originalView);
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                        bool explodedAutoIsoApplied;
                        object explodedView = CreateSinglePartViewForFit(
                            sheet,
                            viewCoordinateSystem,
                            displayCoordinateSystem,
                            plan.Identifier,
                            sourceScaleForFit,
                            ExplodedViewPrefix + viewSequence.ToString("D3") + "_EXP",
                            out explodedAutoIsoApplied);
                        viewSequence++;

                        bool explodedCreated = false;
                        if (explodedView != null)
                        {
                            ApplyFallbackRotation(explodedView, explodedAutoIsoApplied, ref fallbackRotationCount);

                            if (ForceViewCenterToTarget(
                                    explodedView,
                                    anchorX + plan.OffsetX,
                                    anchorY + plan.OffsetY,
                                    anchorZ))
                            {
                                if (colorirEnabled)
                                {
                                    int explodedContourColor = GetContourColorForPlan(plan);
                                    if (explodedContourColor > 0 && TryApplyContourColorToView(explodedView, explodedContourColor))
                                    {
                                        contourColoredCount++;
                                    }
                                    else
                                    {
                                        contourNotChangedCount++;
                                    }
                                }

                                if (explodedAutoIsoApplied)
                                {
                                    autoIsoAppliedCount++;
                                }

                                created++;
                                ghostExplodedCreated++;
                                fitViews.Add(explodedView);
                                explodedCreated = true;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        else
                        {
                            failed++;
                        }

                if (guideLinesEnabled)
                {
                    guideLinesRequested++;
                    if (explodedCreated
                        && TryCreateGuideLineForPlan(
                            sheet,
                            anchorX + plan.OriginalOffsetX,
                            anchorY + plan.OriginalOffsetY,
                            anchorZ,
                            anchorX + plan.OffsetX,
                            anchorY + plan.OffsetY,
                            anchorZ,
                            plan,
                            colorirEnabled))
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

                    string viewName = ExplodedViewPrefix + viewSequence.ToString("D3");
                    viewSequence++;
                    bool autoIsoApplied;

                    object newView = CreateSinglePartViewForFit(
                        sheet,
                        viewCoordinateSystem,
                        displayCoordinateSystem,
                        plan.Identifier,
                        sourceScaleForFit,
                        viewName,
                        out autoIsoApplied);

                    if (newView == null)
                    {
                        failed++;
                        continue;
                    }

                    if (autoIsoApplied)
                    {
                        autoIsoAppliedCount++;
                    }

                    ApplyFallbackRotation(newView, autoIsoApplied, ref fallbackRotationCount);

                    bool positioned = ForceViewCenterToTarget(
                        newView,
                        anchorX + plan.OffsetX,
                        anchorY + plan.OffsetY,
                        anchorZ);
                    if (!positioned)
                    {
                        failed++;
                        continue;
                    }

                    if (colorirEnabled)
                    {
                        int contourColor = GetContourColorForPlan(plan);
                        if (contourColor > 0 && TryApplyContourColorToView(newView, contourColor))
                        {
                            contourColoredCount++;
                        }
                        else
                        {
                            contourNotChangedCount++;
                        }
                    }

                    created++;
                    fitViews.Add(newView);

                    if (guideLinesEnabled && !plan.IsMainPart && IsPlanActuallyDisplaced(plan))
                    {
                        guideLinesRequested++;
                        if (TryCreateGuideLineForPlan(
                                sheet,
                                anchorX + plan.OriginalOffsetX,
                                anchorY + plan.OriginalOffsetY,
                                anchorZ,
                                anchorX + plan.OffsetX,
                                anchorY + plan.OffsetY,
                                anchorZ,
                                plan,
                                colorirEnabled))
                        {
                            guideLinesCreated++;
                        }
                        else
                        {
                            guideLinesFailed++;
                        }
                    }
                }

                bool fitRectangleCreated = false;
                double explodedBoundsMinX;
                double explodedBoundsMinY;
                double explodedBoundsMaxX;
                double explodedBoundsMaxY;
                double yellowRectWidth = 0.0;
                double yellowRectHeight = 0.0;
                if (testEnabled
                    && TryGetViewsUnionBounds(fitViews, out explodedBoundsMinX, out explodedBoundsMinY, out explodedBoundsMaxX, out explodedBoundsMaxY))
                {
                    yellowRectWidth = Math.Max(0.0, (explodedBoundsMaxX - explodedBoundsMinX) + (2.0 * FitRectangleBorderPadding));
                    yellowRectHeight = Math.Max(0.0, (explodedBoundsMaxY - explodedBoundsMinY) + (2.0 * FitRectangleBorderPadding));
                    fitRectangleCreated = TryCreateFitRectangle(
                        sheet,
                        explodedBoundsMinX - FitRectangleBorderPadding,
                        explodedBoundsMinY - FitRectangleBorderPadding,
                        explodedBoundsMaxX + FitRectangleBorderPadding,
                        explodedBoundsMaxY + FitRectangleBorderPadding,
                        anchorZ);
                }

                if (testEnabled)
                {
                    testAxisRequested = 3;
                    if (sourceView != null)
                    {
                        testAxisCreated = TryCreateMainPartTestAxesOverlay(
                            sheet,
                            sourceView,
                            model,
                            partPlans,
                            sourceScaleForFit,
                            anchorX,
                            anchorY,
                            anchorZ,
                            true,
                            true,
                            true);
                    }
                    else
                    {
                        double yawRadians;
                        double pitchRadians;
                        if (!TryGetPreviewCameraAngles(out yawRadians, out pitchRadians))
                        {
                            yawRadians = 0.0;
                            pitchRadians = 0.0;
                        }

                        double cosYaw = Math.Cos(yawRadians);
                        double sinYaw = Math.Sin(yawRadians);
                        double cosPitch = Math.Cos(pitchRadians);
                        double sinPitch = Math.Sin(pitchRadians);
                        double axisLength = Math.Max(TestAxisMinimumLength, Math.Min(Math.Abs(rectMaxX - rectMinX), Math.Abs(rectMaxY - rectMinY)) * 0.32);

                        testAxisCreated = TryCreatePreviewAlignedTestAxesOverlay(
                            sheet,
                            anchorX,
                            anchorY,
                            anchorZ,
                            axisLength,
                            cosYaw,
                            sinPitch * sinYaw,
                            0.0,
                            cosPitch,
                            sinYaw,
                            -sinPitch * cosYaw,
                            true,
                            true,
                            true);
                    }
                }

                CommitActiveDrawingChanges(activeDrawing);
                InvokeParameterlessMethod(drawingHandler, "SaveActiveDrawing");

                double rectWidth = Math.Abs(rectMaxX - rectMinX);
                double rectHeight = Math.Abs(rectMaxY - rectMinY);

                StringBuilder report = new StringBuilder();
                report.AppendLine(useSelectedArea
                    ? "Vista explodida em area definida"
                    : "Vista explodida em folha inteira");
                report.AppendLine("--------------------------------");
                report.AppendLine("Eixos ativos: " + selectedAxes);
                report.AppendLine("Ghost: " + (ghostEnabled ? "ligado" : "desligado"));
                report.AppendLine("Linhas: " + (guideLinesEnabled ? "ligado" : "desligado"));
                report.AppendLine("Colorir: " + (colorirEnabled ? "ligado" : "desligado"));
                report.AppendLine("Teste: " + (testEnabled ? "ligado" : "desligado"));
                report.AppendLine("Sentidos ativos: " + BuildSelectedDirectionsText(
                    allowXPositive,
                    allowXNegative,
                    allowYPositive,
                    allowYNegative,
                    allowZPositive,
                    allowZNegative));
                report.AppendLine("Modo: " + targetAreaSource);
                report.AppendLine("Area alvo usada: " + targetAreaSource);
                report.AppendLine("Objetos selecionados no desenho: " + selectedObjectCount);
                report.AppendLine("Objetos usados para retangulo: " + boundedSelectedCount);
                report.AppendLine("Area alvo: W=" + rectWidth.ToString("0.###") + " H=" + rectHeight.ToString("0.###"));
                report.AppendLine("Vista isometrica detectada: " + (string.IsNullOrWhiteSpace(sourceViewName) ? "(sem nome)" : sourceViewName));
                report.AppendLine("Score isometrico da vista: " + sourceIsoScore.ToString("0.###"));
                report.AppendLine("Vistas inspecionadas na folha: " + inspectedViews);
                report.AppendLine("Vistas candidatas com partes: " + candidateViews);
                report.AppendLine("Partes candidatas na vista: " + candidateCount + (truncated ? " (limitado para " + maxPartsAllowed + ")" : string.Empty));
                report.AppendLine("Vistas explodidas antigas removidas: " + removedExisting);
                report.AppendLine("Linhas guia antigas removidas: " + removedGuideLines);
                report.AppendLine("Base de posicionamento: " + (modelLayoutUsed ? "modelo (explosao por planos)" : "fallback"));
                report.AppendLine("Fallback usado: " + (fallbackLayoutUsed ? "sim" : "nao"));
                report.AppendLine("Pecas bloqueadas por sentido: " + blockedByDirection);
                report.AppendLine("Centros lidos do modelo: " + centersFromModel);
                report.AppendLine("Movimentos aplicados por plano:");
                report.AppendLine("  xy: " + movedByXY);
                report.AppendLine("  xz: " + movedByXZ);
                report.AppendLine("  zy: " + movedByZY);
                report.AppendLine("Eixos usados para explosao: " + (usePreviewAxesForExplode ? "camera do preview" : "vista do desenho"));
                report.AppendLine("Ajuste de escala automatico: " + (fitComputed ? "ligado" : "fallback sem ajuste"));
                report.AppendLine("Fator de fit aplicado: " + fitScaleFactor.ToString("0.###"));
                report.AppendLine("Layout base calculado: W=" + layoutWidth.ToString("0.###") + " H=" + layoutHeight.ToString("0.###"));
                report.AppendLine("Layout ajustado para fit: W=" + fittedLayoutWidth.ToString("0.###") + " H=" + fittedLayoutHeight.ToString("0.###"));
                report.AppendLine("Area alvo informada: W=" + rectWidth.ToString("0.###") + " H=" + rectHeight.ToString("0.###"));
                report.AppendLine("Retangulo de fit criado: W=" + yellowRectWidth.ToString("0.###") + " H=" + yellowRectHeight.ToString("0.###"));
                report.AppendLine("Retangulo de fit criado: " + (fitRectangleCreated ? "sim" : "nao"));
                report.AppendLine("Escala da vista fonte: " + sourceScale.ToString("0.###"));
                report.AppendLine("Escala usada na nova vista: " + sourceScaleForFit.ToString("0.###"));
                report.AppendLine("Perfil de vista solicitado: " + ExplodedViewAttributeFile);
                report.AppendLine("Vistas com perfil aplicado: " + autoIsoAppliedCount);
                report.AppendLine("Rotacao isometrica de fallback: " + fallbackRotationCount);
                report.AppendLine("Cor de contorno aplicada: " + contourColoredCount);
                report.AppendLine("Sem mudanca de cor: " + contourNotChangedCount);
                if (ghostEnabled)
                {
                    report.AppendLine("Ghost pares secundarios: " + ghostPairsRequested);
                    report.AppendLine("Ghost vistas originais criadas: " + ghostOriginalCreated);
                    report.AppendLine("Ghost vistas deslocadas criadas: " + ghostExplodedCreated);
                }

                if (guideLinesEnabled)
                {
                    report.AppendLine("Linhas guia solicitadas: " + guideLinesRequested);
                    report.AppendLine("Linhas guia criadas: " + guideLinesCreated);
                    report.AppendLine("Linhas guia falhas: " + guideLinesFailed);
                }
                if (testEnabled)
                {
                    report.AppendLine("Linhas de teste solicitadas: " + testAxisRequested);
                    report.AppendLine("Linhas de teste criadas: " + testAxisCreated);
                }
                report.AppendLine("Vistas criadas: " + created);
                report.AppendLine("Falhas: " + failed);
                report.AppendLine();
                report.AppendLine("Anchor final no retangulo (sheet): X=" + anchorX.ToString("0.###") + " Y=" + anchorY.ToString("0.###"));
                report.AppendLine("Origem Z da vista fonte: " + anchorZ.ToString("0.###"));
                report.AppendLine();
                report.AppendLine("Observacao: fluxo 100% drawing-only (sem leitura/escrita no modelo).");

                SetOutput(report.ToString());

                if (created > 0)
                {
                    UpdateStatus(
                        useSelectedArea
                            ? "Sucesso: vista explodida criada na area definida."
                            : "Sucesso: vista explodida criada na folha inteira.",
                        Color.DarkGreen);
                }
                else
                {
                    UpdateStatus("Falha: nenhuma vista foi criada na area alvo.", Color.DarkOrange);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Erro ao criar vista ajustada: " + GetInnermostExceptionMessage(ex), Color.DarkRed);
            }
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

            return GetAxisColorForPlan(plan);
        }

        private static int GetAxisColorForPlan(ExplodedPartPlan plan)
        {
            if (plan == null)
            {
                return 0;
            }

            if (plan.ColorPlane == 3 || plan.IsOtherPlane)
            {
                return OtherPlaneContourColor;
            }

            if (plan.ColorPlane == 2)
            {
                return SideNegativeContourColor;
            }

            if (plan.ColorPlane == 1)
            {
                return SidePositiveContourColor;
            }

            if (plan.SideOfZYPlane != 0)
            {
                return OtherPlaneContourColor;
            }

            if (plan.SideOfXZPlane != 0)
            {
                return SideNegativeContourColor;
            }

            if (plan.SideOfXYPlane != 0)
            {
                return SidePositiveContourColor;
            }

            return 0;
        }

        private static void ApplyDirectionalMovementFilters(
            List<ExplodedPartPlan> plans,
            bool allowXPositive,
            bool allowXNegative,
            bool allowYPositive,
            bool allowYNegative,
            bool allowZPositive,
            bool allowZNegative,
            out int blockedCount)
        {
            blockedCount = 0;
            if (plans == null || plans.Count == 0)
            {
                return;
            }

            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                if (plan == null || plan.IsMainPart)
                {
                    continue;
                }

                double filteredOffsetX = plan.OriginalOffsetX;
                double filteredOffsetY = plan.OriginalOffsetY;
                int dominantAxisCode;
                int dominantAxisSide;
                if (TryGetPlanDominantDirection(plan, out dominantAxisCode, out dominantAxisSide))
                {
                    if (dominantAxisCode == 1 && IsDirectionEnabled(dominantAxisSide, allowXPositive, allowXNegative))
                    {
                        filteredOffsetX += plan.OffsetFromXAxisX;
                        filteredOffsetY += plan.OffsetFromXAxisY;
                    }
                    else if (dominantAxisCode == 2 && IsDirectionEnabled(dominantAxisSide, allowYPositive, allowYNegative))
                    {
                        filteredOffsetX += plan.OffsetFromYAxisX;
                        filteredOffsetY += plan.OffsetFromYAxisY;
                    }
                    else if (dominantAxisCode == 3 && IsDirectionEnabled(dominantAxisSide, allowZPositive, allowZNegative))
                    {
                        filteredOffsetX += plan.OffsetFromZAxisX;
                        filteredOffsetY += plan.OffsetFromZAxisY;
                    }
                }

                bool movementBlocked = Math.Abs(filteredOffsetX - plan.OffsetX) > 1e-6
                    || Math.Abs(filteredOffsetY - plan.OffsetY) > 1e-6;

                plan.OffsetX = filteredOffsetX;
                plan.OffsetY = filteredOffsetY;

                if (movementBlocked)
                {
                    blockedCount++;
                }
            }
        }

        private static bool IsDirectionEnabled(int axisSide, bool allowPositive, bool allowNegative)
        {
            if (axisSide > 0)
            {
                return allowPositive;
            }

            if (axisSide < 0)
            {
                return allowNegative;
            }

            return false;
        }

        private static bool TryGetPlanDominantDirection(ExplodedPartPlan plan, out int axisCode, out int axisSide)
        {
            axisCode = 0;
            axisSide = 0;
            if (plan == null)
            {
                return false;
            }

            if ((plan.ColorPlane == 3 || plan.IsOtherPlane) && plan.SideOfZYPlane != 0)
            {
                axisCode = 1;
                axisSide = plan.SideOfZYPlane;
                return true;
            }

            if (plan.ColorPlane == 2 && plan.SideOfXZPlane != 0)
            {
                axisCode = 2;
                axisSide = plan.SideOfXZPlane;
                return true;
            }

            if (plan.ColorPlane == 1 && plan.SideOfXYPlane != 0)
            {
                axisCode = 3;
                axisSide = plan.SideOfXYPlane;
                return true;
            }

            if (plan.SideOfZYPlane != 0)
            {
                axisCode = 1;
                axisSide = plan.SideOfZYPlane;
                return true;
            }

            if (plan.SideOfXZPlane != 0)
            {
                axisCode = 2;
                axisSide = plan.SideOfXZPlane;
                return true;
            }

            if (plan.SideOfXYPlane != 0)
            {
                axisCode = 3;
                axisSide = plan.SideOfXYPlane;
                return true;
            }

            return false;
        }

        private static bool IsPlanActuallyDisplaced(ExplodedPartPlan plan)
        {
            if (plan == null)
            {
                return false;
            }

            return Math.Abs(plan.OffsetX - plan.OriginalOffsetX) > 1e-6
                || Math.Abs(plan.OffsetY - plan.OriginalOffsetY) > 1e-6;
        }

        private static bool TryCreateGuideLineForPlan(
            object viewBase,
            double startX,
            double startY,
            double startZ,
            double endX,
            double endY,
            double endZ,
            ExplodedPartPlan plan,
            bool colorirEnabled)
        {
            if (!colorirEnabled)
            {
                return TryCreateGuideLine(viewBase, startX, startY, startZ, endX, endY, endZ);
            }

            int colorIndex = GetAxisColorForPlan(plan);
            if (colorIndex <= 0)
            {
                colorIndex = GuideLineColor;
            }

            return TryCreateGuideLineColored(viewBase, startX, startY, startZ, endX, endY, endZ, colorIndex);
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

            string guideLineName = GuideLineNamePrefix + Guid.NewGuid().ToString("N");
            if (lineAttributes != null)
            {
                SetPropertyValue(lineAttributes, "Name", guideLineName);
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

            SetPropertyValue(lineObject, "Name", guideLineName);
            bool? inserted = InvokeBoolMethod(lineObject, "Insert");
            return inserted.HasValue && inserted.Value;
        }

        private static bool TryCreateGuideLineColored(
            object viewBase,
            double startX,
            double startY,
            double startZ,
            double endX,
            double endY,
            double endZ,
            int colorIndex)
        {
            string guideLineName = GuideLineNamePrefix + Guid.NewGuid().ToString("N");
            return TryCreateNamedLine(
                viewBase,
                startX,
                startY,
                startZ,
                endX,
                endY,
                endZ,
                guideLineName,
                colorIndex,
                true);
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
                changed |= TrySetColorProperty(lineTypeAttributes, "Color", GuideLineColor);
                changed |= TrySetColorProperty(lineTypeAttributes, "TrueColor", GuideLineColor);
                SetPropertyValue(lineAttributes, "Line", lineTypeAttributes);
            }

            return changed;
        }

        private static bool TryCreateFitRectangle(
            object viewBase,
            double minX,
            double minY,
            double maxX,
            double maxY,
            double z)
        {
            if (viewBase == null || maxX <= minX || maxY <= minY)
            {
                return false;
            }

            string rectangleName = FitRectangleNamePrefix + Guid.NewGuid().ToString("N");
            int created = 0;

            if (TryCreateNamedLine(viewBase, minX, minY, z, maxX, minY, z, rectangleName + "_01", GuideLineColor, true))
            {
                created++;
            }

            if (TryCreateNamedLine(viewBase, maxX, minY, z, maxX, maxY, z, rectangleName + "_02", GuideLineColor, true))
            {
                created++;
            }

            if (TryCreateNamedLine(viewBase, maxX, maxY, z, minX, maxY, z, rectangleName + "_03", GuideLineColor, true))
            {
                created++;
            }

            if (TryCreateNamedLine(viewBase, minX, maxY, z, minX, minY, z, rectangleName + "_04", GuideLineColor, true))
            {
                created++;
            }

            return created == 4;
        }

        private static int TryCreateMainPartTestAxesOverlay(
            object viewBase,
            object sourceView,
            object model,
            List<ExplodedPartPlan> plans,
            double viewScale,
            double anchorX,
            double anchorY,
            double anchorZ,
            bool drawAxisX,
            bool drawAxisY,
            bool drawAxisZ)
        {
            return TryCreateMainPartTestAxesOverlay(
                viewBase,
                sourceView,
                model,
                plans,
                viewScale,
                anchorX,
                anchorY,
                anchorZ,
                drawAxisX,
                drawAxisY,
                drawAxisZ,
                false,
                0.0,
                0.0,
                0.0,
                0.0,
                0.0,
                0.0);
        }

        private static int TryCreateMainPartTestAxesOverlay(
            object viewBase,
            object sourceView,
            object model,
            List<ExplodedPartPlan> plans,
            double viewScale,
            double anchorX,
            double anchorY,
            double anchorZ,
            bool drawAxisX,
            bool drawAxisY,
            bool drawAxisZ,
            bool alignSignsToReference,
            double referenceAxisXDirX,
            double referenceAxisXDirY,
            double referenceAxisYDirX,
            double referenceAxisYDirY,
            double referenceAxisZDirX,
            double referenceAxisZDirY)
        {
            if (viewBase == null || sourceView == null || model == null || plans == null || plans.Count == 0)
            {
                return 0;
            }

            bool? modelConnected = InvokeBoolMethod(model, "GetConnectionStatus");
            if (!modelConnected.HasValue || !modelConnected.Value)
            {
                return 0;
            }

            ExplodedPartPlan mainPlan;
            if (!TryGetMainPartPlanForTestAxes(plans, out mainPlan) || mainPlan == null || mainPlan.Identifier == null)
            {
                return 0;
            }

            double modelCenterX;
            double modelCenterY;
            double modelCenterZ;
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
                    mainPlan.Identifier,
                    out modelCenterX,
                    out modelCenterY,
                    out modelCenterZ,
                    out sizeScore,
                    out hasBounds,
                    out minX,
                    out minY,
                    out minZ,
                    out maxX,
                    out maxY,
                    out maxZ))
            {
                return 0;
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
                return 0;
            }

            double viewAxisXx;
            double viewAxisXy;
            double viewAxisXz;
            double viewAxisYx;
            double viewAxisYy;
            double viewAxisYz;
            if (!TryGetViewDisplayAxes(
                    sourceView,
                    out viewAxisXx,
                    out viewAxisXy,
                    out viewAxisXz,
                    out viewAxisYx,
                    out viewAxisYy,
                    out viewAxisYz))
            {
                return 0;
            }

            double axisXSheetX = Dot3D(localAxisXx, localAxisXy, localAxisXz, viewAxisXx, viewAxisXy, viewAxisXz);
            double axisXSheetY = Dot3D(localAxisXx, localAxisXy, localAxisXz, viewAxisYx, viewAxisYy, viewAxisYz);
            double axisYSheetX = Dot3D(localAxisYx, localAxisYy, localAxisYz, viewAxisXx, viewAxisXy, viewAxisXz);
            double axisYSheetY = Dot3D(localAxisYx, localAxisYy, localAxisYz, viewAxisYx, viewAxisYy, viewAxisYz);
            double axisZSheetX = Dot3D(localAxisZx, localAxisZy, localAxisZz, viewAxisXx, viewAxisXy, viewAxisXz);
            double axisZSheetY = Dot3D(localAxisZx, localAxisZy, localAxisZz, viewAxisYx, viewAxisYy, viewAxisYz);

            if (!NormalizeVector2D(ref axisXSheetX, ref axisXSheetY)
                || !NormalizeVector2D(ref axisYSheetX, ref axisYSheetY)
                || !NormalizeVector2D(ref axisZSheetX, ref axisZSheetY))
            {
                return 0;
            }

            if (alignSignsToReference)
            {
                AlignAxisDirectionByReference(ref axisXSheetX, ref axisXSheetY, referenceAxisXDirX, referenceAxisXDirY);
                AlignAxisDirectionByReference(ref axisYSheetX, ref axisYSheetY, referenceAxisYDirX, referenceAxisYDirY);
                AlignAxisDirectionByReference(ref axisZSheetX, ref axisZSheetY, referenceAxisZDirX, referenceAxisZDirY);
            }

            double scaleDivisor = viewScale > 0.0 ? viewScale : 1.0;
            double centerSheetX = anchorX + mainPlan.OffsetX;
            double centerSheetY = anchorY + mainPlan.OffsetY;

            int created = 0;
            if (drawAxisX)
            {
                double spanX = ComputeMainPartAxisSpanForTest(
                    hasBounds,
                    modelCenterX,
                    modelCenterY,
                    modelCenterZ,
                    minX,
                    minY,
                    minZ,
                    maxX,
                    maxY,
                    maxZ,
                    localAxisXx,
                    localAxisXy,
                    localAxisXz,
                    sizeScore);
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerSheetX,
                        centerSheetY,
                        anchorZ,
                        axisXSheetX,
                        axisXSheetY,
                        spanX / scaleDivisor,
                        OtherPlaneContourColor,
                        "X"))
                {
                    created++;
                }
            }

            if (drawAxisY)
            {
                double spanY = ComputeMainPartAxisSpanForTest(
                    hasBounds,
                    modelCenterX,
                    modelCenterY,
                    modelCenterZ,
                    minX,
                    minY,
                    minZ,
                    maxX,
                    maxY,
                    maxZ,
                    localAxisYx,
                    localAxisYy,
                    localAxisYz,
                    sizeScore);
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerSheetX,
                        centerSheetY,
                        anchorZ,
                        axisYSheetX,
                        axisYSheetY,
                        spanY / scaleDivisor,
                        SideNegativeContourColor,
                        "Y"))
                {
                    created++;
                }
            }

            if (drawAxisZ)
            {
                double spanZ = ComputeMainPartAxisSpanForTest(
                    hasBounds,
                    modelCenterX,
                    modelCenterY,
                    modelCenterZ,
                    minX,
                    minY,
                    minZ,
                    maxX,
                    maxY,
                    maxZ,
                    localAxisZx,
                    localAxisZy,
                    localAxisZz,
                    sizeScore);
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerSheetX,
                        centerSheetY,
                        anchorZ,
                        axisZSheetX,
                        axisZSheetY,
                        spanZ / scaleDivisor,
                        SidePositiveContourColor,
                        "Z"))
                {
                    created++;
                }
            }

            return created;
        }

        private static void AlignAxisDirectionByReference(
            ref double axisDirX,
            ref double axisDirY,
            double referenceDirX,
            double referenceDirY)
        {
            double normalizedAxisX = axisDirX;
            double normalizedAxisY = axisDirY;
            if (!NormalizeVector2D(ref normalizedAxisX, ref normalizedAxisY))
            {
                return;
            }

            double normalizedReferenceX = referenceDirX;
            double normalizedReferenceY = referenceDirY;
            if (!NormalizeVector2D(ref normalizedReferenceX, ref normalizedReferenceY))
            {
                return;
            }

            double sameDirectionScore =
                (normalizedAxisX * normalizedReferenceX) + (normalizedAxisY * normalizedReferenceY);
            if (sameDirectionScore < 0.0)
            {
                axisDirX = -axisDirX;
                axisDirY = -axisDirY;
            }
        }

        private static int TryCreatePreviewAlignedTestAxesOverlay(
            object viewBase,
            double centerX,
            double centerY,
            double centerZ,
            double axisLength,
            double axisXDirX,
            double axisXDirY,
            double axisYDirX,
            double axisYDirY,
            double axisZDirX,
            double axisZDirY,
            bool drawAxisX,
            bool drawAxisY,
            bool drawAxisZ)
        {
            if (viewBase == null)
            {
                return 0;
            }

            double safeLength = Math.Max(TestAxisMinimumLength, axisLength);
            int created = 0;

            double xDirX = axisXDirX;
            double xDirY = axisXDirY;
            double yDirX = axisYDirX;
            double yDirY = axisYDirY;
            double zDirX = axisZDirX;
            double zDirY = axisZDirY;

            bool hasX = NormalizeVector2D(ref xDirX, ref xDirY);
            bool hasY = NormalizeVector2D(ref yDirX, ref yDirY);
            bool hasZ = NormalizeVector2D(ref zDirX, ref zDirY);

            if (drawAxisX && hasX)
            {
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerX,
                        centerY,
                        centerZ,
                        xDirX,
                        xDirY,
                        safeLength,
                        OtherPlaneContourColor,
                        "X"))
                {
                    created++;
                }
            }

            if (drawAxisY && hasY)
            {
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerX,
                        centerY,
                        centerZ,
                        yDirX,
                        yDirY,
                        safeLength,
                        SideNegativeContourColor,
                        "Y"))
                {
                    created++;
                }
            }

            if (drawAxisZ && hasZ)
            {
                if (TryCreateCenteredTestAxisLine(
                        viewBase,
                        centerX,
                        centerY,
                        centerZ,
                        zDirX,
                        zDirY,
                        safeLength,
                        SidePositiveContourColor,
                        "Z"))
                {
                    created++;
                }
            }

            return created;
        }

        private static bool TryGetMainPartPlanForTestAxes(List<ExplodedPartPlan> plans, out ExplodedPartPlan mainPlan)
        {
            mainPlan = null;
            if (plans == null || plans.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < plans.Count; i++)
            {
                if (plans[i] != null && plans[i].IsMainPart)
                {
                    mainPlan = plans[i];
                    return true;
                }
            }

            double bestScore = double.MinValue;
            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan candidate = plans[i];
                if (candidate == null)
                {
                    continue;
                }

                double score = candidate.ModelSizeScore > 0.0 ? candidate.ModelSizeScore : candidate.ApproxSize2D;
                if (score > bestScore)
                {
                    bestScore = score;
                    mainPlan = candidate;
                }
            }

            return mainPlan != null;
        }

        private static double ComputeMainPartAxisSpanForTest(
            bool hasBounds,
            double centerX,
            double centerY,
            double centerZ,
            double minX,
            double minY,
            double minZ,
            double maxX,
            double maxY,
            double maxZ,
            double axisX,
            double axisY,
            double axisZ,
            double fallbackSpan)
        {
            double span = hasBounds
                ? ComputeBoundsSpanAlongAxisForTest(
                    centerX,
                    centerY,
                    centerZ,
                    minX,
                    minY,
                    minZ,
                    maxX,
                    maxY,
                    maxZ,
                    axisX,
                    axisY,
                    axisZ)
                : 0.0;

            if (span <= 1e-6)
            {
                span = Math.Max(1.0, fallbackSpan);
            }

            return span * (1.0 + TestAxisPaddingRatio);
        }

        private static double ComputeBoundsSpanAlongAxisForTest(
            double centerX,
            double centerY,
            double centerZ,
            double minX,
            double minY,
            double minZ,
            double maxX,
            double maxY,
            double maxZ,
            double axisX,
            double axisY,
            double axisZ)
        {
            double[] xs = { minX, maxX };
            double[] ys = { minY, maxY };
            double[] zs = { minZ, maxZ };

            bool initialized = false;
            double minLocal = 0.0;
            double maxLocal = 0.0;

            for (int ix = 0; ix < xs.Length; ix++)
            {
                for (int iy = 0; iy < ys.Length; iy++)
                {
                    for (int iz = 0; iz < zs.Length; iz++)
                    {
                        double localValue = Dot3D(
                            xs[ix] - centerX,
                            ys[iy] - centerY,
                            zs[iz] - centerZ,
                            axisX,
                            axisY,
                            axisZ);

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

            return initialized ? Math.Max(0.0, maxLocal - minLocal) : 0.0;
        }

        private static bool TryCreateCenteredTestAxisLine(
            object viewBase,
            double centerX,
            double centerY,
            double centerZ,
            double directionX,
            double directionY,
            double sheetLength,
            int colorIndex,
            string axisName)
        {
            if (viewBase == null)
            {
                return false;
            }

            double safeLength = Math.Max(TestAxisMinimumLength, sheetLength);
            double halfLength = safeLength * 0.5;
            double labelOffset = Math.Max(3.0, safeLength * 0.04);
            double startX = centerX - (directionX * halfLength);
            double startY = centerY - (directionY * halfLength);
            double endX = centerX + (directionX * halfLength);
            double endY = centerY + (directionY * halfLength);
            string lineName = TestAxisNamePrefix + axisName + "_" + Guid.NewGuid().ToString("N");

            bool lineCreated = TryCreateNamedLine(
                viewBase,
                startX,
                startY,
                centerZ,
                endX,
                endY,
                centerZ,
                lineName,
                colorIndex,
                false);
            if (!lineCreated)
            {
                return false;
            }

            TryCreateAxisLabelText(
                viewBase,
                endX + (directionX * labelOffset),
                endY + (directionY * labelOffset),
                centerZ,
                axisName + "+",
                colorIndex);
            TryCreateAxisLabelText(
                viewBase,
                startX - (directionX * labelOffset),
                startY - (directionY * labelOffset),
                centerZ,
                axisName + "-",
                colorIndex);
            return true;
        }

        private static bool TryCreateAxisLabelText(
            object viewBase,
            double x,
            double y,
            double z,
            string labelText,
            int colorIndex)
        {
            if (viewBase == null || string.IsNullOrWhiteSpace(labelText))
            {
                return false;
            }

            object insertionPoint = CreatePoint(x, y, z);
            if (insertionPoint == null)
            {
                return false;
            }

            Type textType = ResolveTeklaType(
                "Tekla.Structures.Drawing.Text",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (textType == null)
            {
                return false;
            }

            Type textAttributesType = ResolveTeklaType(
                "Tekla.Structures.Drawing.Text+TextAttributes",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);

            object textAttributes = null;
            if (textAttributesType != null)
            {
                try
                {
                    textAttributes = Activator.CreateInstance(textAttributesType);
                }
                catch
                {
                    textAttributes = null;
                }
            }

            if (textAttributes != null)
            {
                SetPropertyValue(textAttributes, "TransparentBackground", true);

                object font = GetPropertyValue(textAttributes, "Font");
                if (font == null)
                {
                    Type fontType = ResolveTeklaType(
                        "Tekla.Structures.Drawing.FontAttributes",
                        "Tekla.Structures.Drawing",
                        TeklaDrawingDllCandidates);
                    if (fontType != null)
                    {
                        try
                        {
                            font = Activator.CreateInstance(fontType);
                        }
                        catch
                        {
                            font = null;
                        }
                    }
                }

                if (font != null)
                {
                    TrySetColorProperty(font, "Color", colorIndex);
                    TrySetColorProperty(font, "TrueColor", colorIndex);
                    SetPropertyValue(font, "Height", 4.0);
                    SetPropertyValue(textAttributes, "Font", font);
                }
            }

            object textObject = null;
            string objectName = TestAxisNamePrefix + labelText + "_" + Guid.NewGuid().ToString("N");
            if (textAttributes != null)
            {
                try
                {
                    textObject = Activator.CreateInstance(textType, new object[] { viewBase, insertionPoint, labelText, textAttributes });
                }
                catch
                {
                    textObject = null;
                }
            }

            if (textObject == null)
            {
                try
                {
                    textObject = Activator.CreateInstance(textType, new object[] { viewBase, insertionPoint, labelText });
                }
                catch
                {
                    textObject = null;
                }
            }

            if (textObject == null)
            {
                return false;
            }

            SetPropertyValue(textAttributes, "Name", objectName);
            if (textAttributes != null)
            {
                SetPropertyValue(textObject, "Attributes", textAttributes, "Tekla.Structures.Drawing.Text");
            }

            SetPropertyValue(textObject, "Name", objectName);

            bool? inserted = InvokeBoolMethod(textObject, "Insert");
            return inserted.HasValue && inserted.Value;
        }

        private static bool TryCreateNamedLine(
            object viewBase,
            double startX,
            double startY,
            double startZ,
            double endX,
            double endY,
            double endZ,
            string lineName,
            int colorIndex,
            bool dashed)
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
                ApplyNamedLineStyle(lineAttributes, colorIndex, dashed);
                SetPropertyValue(lineAttributes, "Name", lineName);
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

            SetPropertyValue(lineObject, "Name", lineName);
            bool? inserted = InvokeBoolMethod(lineObject, "Insert");
            return inserted.HasValue && inserted.Value;
        }

        private static bool ApplyNamedLineStyle(object lineAttributes, int colorIndex, bool dashed)
        {
            if (lineAttributes == null)
            {
                return false;
            }

            bool changed = false;
            object lineTypeAttributes = GetPropertyValue(lineAttributes, "Line");
            if (lineTypeAttributes != null)
            {
                if (dashed)
                {
                    changed |= TrySetLineTypeDashed(lineTypeAttributes);
                }

                changed |= TrySetColorProperty(lineTypeAttributes, "Color", colorIndex);
                changed |= TrySetColorProperty(lineTypeAttributes, "TrueColor", colorIndex);
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
                bool? modifiedView = InvokeBoolMethod(view, "Modify");
                if (modifiedView.HasValue && !modifiedView.Value)
                {
                    return false;
                }
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

            changed |= TryClearGhostFill(partAttributes);

            if (!changed)
            {
                return false;
            }

            bool attributesSet = SetPropertyValue(drawingObject, "Attributes", partAttributes, "Tekla.Structures.Drawing.Part")
                || SetPropertyValue(drawingObject, "Attributes", partAttributes);
            return attributesSet;
        }

        private static bool TryClearGhostFill(object partAttributes)
        {
            if (partAttributes == null)
            {
                return false;
            }

            bool changed = false;
            changed |= TryClearGhostHatchProperty(partAttributes, "FaceHatch");
            changed |= TryClearGhostHatchProperty(partAttributes, "SectionFaceHatch");
            changed |= TryClearGhostHatchProperty(partAttributes, "Hatch");
            return changed;
        }

        private static bool TryClearGhostHatchProperty(object owner, string propertyName)
        {
            if (owner == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            object hatch = GetPropertyValue(owner, propertyName);
            if (hatch == null)
            {
                return false;
            }

            bool changed = false;
            changed |= SetPropertyValue(hatch, "Name", string.Empty);
            changed |= TrySetColorProperty(hatch, "Color", 0);
            changed |= TrySetColorProperty(hatch, "BackgroundColor", 0);
            changed |= TrySetColorProperty(hatch, "TrueColor", 0);
            changed |= TrySetColorProperty(hatch, "TrueBackgroundColor", 0);
            changed |= SetPropertyValue(hatch, "DrawBackgroundColor", false);
            changed |= SetPropertyValue(owner, propertyName, hatch);
            return changed;
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

        private static bool TrySetLineTypeDashed(object lineTypeAttributes)
        {
            if (lineTypeAttributes == null)
            {
                return false;
            }

            if (!IsLineTypeAttributesObject(lineTypeAttributes))
            {
                return false;
            }

            Type lineTypeType = lineTypeAttributes.GetType();
            Assembly drawingAssembly = lineTypeType != null ? lineTypeType.Assembly : null;
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

            bool changed = SetPropertyValue(lineTypeAttributes, "Type", dashedValue);
            changed |= SetPropertyValue(lineTypeAttributes, "LineType", dashedValue);
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
            lblStatusTekla.Text = message ?? string.Empty;
            lblStatusTekla.ForeColor = color;
            lblStatusTekla.Visible = !string.IsNullOrWhiteSpace(message);
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

        private void ConfigureStatusIndicators()
        {
            ConfigureStatusIndicator(pnlTeklaIndicator);
            ConfigureStatusIndicator(pnlDrawingIndicator);
            SetIndicatorColor(pnlTeklaIndicator, Color.Silver);
            SetIndicatorColor(pnlDrawingIndicator, Color.Silver);
        }

        private void ConfigureStatusIndicator(Panel indicator)
        {
            if (indicator == null)
            {
                return;
            }

            ApplyCircularRegion(indicator);
            indicator.Resize += StatusIndicator_Resize;
        }

        private void StatusIndicator_Resize(object sender, EventArgs e)
        {
            Panel indicator = sender as Panel;
            if (indicator != null)
            {
                ApplyCircularRegion(indicator);
            }
        }

        private static void ApplyCircularRegion(Control control)
        {
            if (control == null || control.Width <= 0 || control.Height <= 0)
            {
                return;
            }

            GraphicsPath path = new GraphicsPath();
            path.AddEllipse(0, 0, control.Width - 1, control.Height - 1);
            Region oldRegion = control.Region;
            control.Region = new Region(path);
            if (oldRegion != null)
            {
                oldRegion.Dispose();
            }

            path.Dispose();
        }

        private static void SetIndicatorColor(Panel indicator, Color color)
        {
            if (indicator != null)
            {
                indicator.BackColor = color;
            }
        }

        private void SetLogExpanded(bool expanded)
        {
            logExpanded = expanded;
            pnlLog.Visible = expanded;
            btnToggleLog.Text = expanded ? "- Ocultar log" : "+ Mostrar log";
            ClientSize = new Size(ClientSize.Width, expanded ? ExpandedClientHeight : CollapsedClientHeight);
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
            TryLoadAssemblyFromCandidates(TeklaStructuresDllCandidates);
            Type modelType = ResolveTeklaModelType();
            return modelType != null ? Activator.CreateInstance(modelType) : null;
        }

        private static Type ResolveTeklaModelType()
        {
            TryLoadAssemblyFromCandidates(TeklaStructuresDllCandidates);
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

                try
                {
                    Assembly loadedAssembly = Assembly.LoadFrom(dllPath);
                    resolved = loadedAssembly.GetType(fullTypeName, false);
                    if (resolved != null)
                    {
                        return resolved;
                    }
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        private static Assembly TryLoadAssemblyFromCandidates(string[] candidatePaths)
        {
            if (candidatePaths == null)
            {
                return null;
            }

            for (int i = 0; i < candidatePaths.Length; i++)
            {
                string dllPath = candidatePaths[i];
                if (string.IsNullOrWhiteSpace(dllPath) || !File.Exists(dllPath))
                {
                    continue;
                }

                string simpleName = Path.GetFileNameWithoutExtension(dllPath);
                Assembly alreadyLoaded = FindLoadedAssembly(simpleName);
                if (alreadyLoaded != null)
                {
                    return alreadyLoaded;
                }

                try
                {
                    return Assembly.LoadFrom(dllPath);
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        private static Assembly FindLoadedAssembly(string simpleAssemblyName)
        {
            if (string.IsNullOrWhiteSpace(simpleAssemblyName))
            {
                return null;
            }

            Assembly[] loaded = AppDomain.CurrentDomain.GetAssemblies();
            for (int i = 0; i < loaded.Length; i++)
            {
                Assembly assembly = loaded[i];
                if (assembly == null)
                {
                    continue;
                }

                AssemblyName name = assembly.GetName();
                if (name != null
                    && string.Equals(name.Name, simpleAssemblyName, StringComparison.OrdinalIgnoreCase))
                {
                    return assembly;
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

        private static bool TryGetSelectedPlacementRectangle(
            object drawingHandler,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY,
            out int selectedObjectCount,
            out int boundedObjectCount)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;
            selectedObjectCount = 0;
            boundedObjectCount = 0;

            if (drawingHandler == null)
            {
                return false;
            }

            object selector = InvokeParameterlessMethod(drawingHandler, "GetDrawingObjectSelector");
            if (selector == null)
            {
                return false;
            }

            object selectedEnumerator = InvokeParameterlessMethod(selector, "GetSelected");
            if (selectedEnumerator == null)
            {
                return false;
            }

            MethodInfo moveNext = selectedEnumerator.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = selectedEnumerator.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return false;
            }

            bool hasBounds = false;
            while (true)
            {
                object moved = moveNext.Invoke(selectedEnumerator, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object selectedObject = current.GetValue(selectedEnumerator, null);
                if (selectedObject == null)
                {
                    continue;
                }

                selectedObjectCount++;

                double objMinX;
                double objMinY;
                double objMaxX;
                double objMaxY;

                bool gotBounds = TryGetExactSelectionBounds2D(selectedObject, out objMinX, out objMinY, out objMaxX, out objMaxY);

                if (!gotBounds)
                {
                    continue;
                }

                if (!hasBounds)
                {
                    minX = objMinX;
                    minY = objMinY;
                    maxX = objMaxX;
                    maxY = objMaxY;
                    hasBounds = true;
                }
                else
                {
                    minX = Math.Min(minX, objMinX);
                    minY = Math.Min(minY, objMinY);
                    maxX = Math.Max(maxX, objMaxX);
                    maxY = Math.Max(maxY, objMaxY);
                }

                boundedObjectCount++;
            }

            if (!hasBounds)
            {
                return false;
            }

            double width = Math.Abs(maxX - minX);
            double height = Math.Abs(maxY - minY);
            return width > 1e-6 && height > 1e-6;
        }

        private static bool TryPickPlacementRectangleByTwoPoints(
            object activeDrawing,
            object sheet,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY,
            out bool cancelled)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;
            cancelled = false;

            Type pickerType = ResolveTeklaType(
                "Tekla.Structures.Drawing.UI.Picker",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (pickerType == null)
            {
                return false;
            }

            if (activeDrawing == null)
            {
                return false;
            }

            object picker;
            try
            {
                picker = Activator.CreateInstance(
                    pickerType,
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    binder: null,
                    args: new object[] { activeDrawing },
                    culture: null);
            }
            catch
            {
                return false;
            }

            MethodInfo pickTwoPointsMethod = null;
            MethodInfo[] pickerMethods = pickerType.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < pickerMethods.Length; i++)
            {
                MethodInfo candidate = pickerMethods[i];
                if (!string.Equals(candidate.Name, "PickTwoPoints", StringComparison.Ordinal))
                {
                    continue;
                }

                ParameterInfo[] parameters = candidate.GetParameters();
                if (parameters.Length == 5)
                {
                    pickTwoPointsMethod = candidate;
                    break;
                }
            }

            if (pickTwoPointsMethod == null)
            {
                return false;
            }

            object[] args = new object[]
            {
                "Clique no primeiro canto da area",
                "Clique no canto oposto da area",
                null,
                null,
                null
            };

            try
            {
                pickTwoPointsMethod.Invoke(picker, args);
            }
            catch (TargetInvocationException ex)
            {
                string errorText = ex.InnerException != null
                    ? ex.InnerException.ToString()
                    : ex.ToString();
                if (!string.IsNullOrWhiteSpace(errorText)
                    && errorText.IndexOf("PickerInterruptedException", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cancelled = true;
                    return false;
                }

                return false;
            }
            catch
            {
                return false;
            }

            object firstPoint = args[2];
            object secondPoint = args[3];
            object pickedView = args[4];
            if (firstPoint == null || secondPoint == null)
            {
                return false;
            }

            object firstPointInSheet;
            object secondPointInSheet;
            if (!TryConvertPointToSheetCoordinates(pickedView, sheet, firstPoint, out firstPointInSheet)
                || !TryConvertPointToSheetCoordinates(pickedView, sheet, secondPoint, out secondPointInSheet))
            {
                return false;
            }

            double x1;
            double y1;
            double z1;
            double x2;
            double y2;
            double z2;
            if (!TryGetXYZ(firstPointInSheet, out x1, out y1, out z1)
                || !TryGetXYZ(secondPointInSheet, out x2, out y2, out z2))
            {
                return false;
            }

            minX = Math.Min(x1, x2);
            minY = Math.Min(y1, y2);
            maxX = Math.Max(x1, x2);
            maxY = Math.Max(y1, y2);
            return Math.Abs(maxX - minX) > 1e-6 && Math.Abs(maxY - minY) > 1e-6;
        }

        private static bool TryGetSheetPlacementRectangle(
            object sheet,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (sheet == null)
            {
                return false;
            }

            if (TryGetKnownPointPairBounds2D(sheet, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            if (TryGetDrawingObjectBounds2D(sheet, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            object origin = GetPropertyValue(sheet, "Origin") ?? GetPropertyValue(sheet, "InsertionPoint");
            double originX;
            double originY;
            double originZ;
            if (!TryGetXYZ(origin, out originX, out originY, out originZ))
            {
                originX = 0.0;
                originY = 0.0;
            }

            double width = ReadDoubleProperty(sheet, "Width", 0.0);
            if (!(width > 1e-6))
            {
                width = ReadDoubleProperty(sheet, "PaperWidth", 0.0);
            }

            double height = ReadDoubleProperty(sheet, "Height", 0.0);
            if (!(height > 1e-6))
            {
                height = ReadDoubleProperty(sheet, "PaperHeight", 0.0);
            }

            if (!(width > 1e-6) || !(height > 1e-6))
            {
                return false;
            }

            minX = originX;
            minY = originY;
            maxX = originX + width;
            maxY = originY + height;
            return true;
        }

        private static bool TryConvertPointToSheetCoordinates(
            object fromView,
            object sheet,
            object pointInView,
            out object pointInSheet)
        {
            pointInSheet = null;

            if (pointInView == null)
            {
                return false;
            }

            if (fromView == null || sheet == null || ReferenceEquals(fromView, sheet))
            {
                pointInSheet = pointInView;
                return true;
            }

            Type converterType = ResolveTeklaType(
                "Tekla.Structures.Drawing.Tools.DrawingCoordinateConverter",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (converterType != null)
            {
                MethodInfo[] methods = converterType.GetMethods(BindingFlags.Public | BindingFlags.Static);
                for (int i = 0; i < methods.Length; i++)
                {
                    MethodInfo method = methods[i];
                    if (!string.Equals(method.Name, "Convert", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    ParameterInfo[] parameters = method.GetParameters();
                    if (parameters.Length != 3)
                    {
                        continue;
                    }

                    if (parameters[2].ParameterType.IsAssignableFrom(pointInView.GetType()))
                    {
                        try
                        {
                            pointInSheet = method.Invoke(null, new object[] { fromView, sheet, pointInView });
                        }
                        catch
                        {
                            pointInSheet = null;
                        }

                        if (pointInSheet != null)
                        {
                            return true;
                        }
                    }
                }
            }

            double px;
            double py;
            double pz;
            double originX;
            double originY;
            double originZ;
            if (TryGetXYZ(pointInView, out px, out py, out pz)
                && TryGetXYZ(GetPropertyValue(fromView, "Origin"), out originX, out originY, out originZ))
            {
                pointInSheet = CreatePoint(originX + px, originY + py, originZ + pz);
                return pointInSheet != null;
            }

            return false;
        }

        private static bool TryGetExactSelectionBounds2D(
            object drawingObject,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (drawingObject == null)
            {
                return false;
            }

            if (TryGetObjectPointBounds2D(drawingObject, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            if (TryGetKnownPointPairBounds2D(drawingObject, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            if (TryGetStrictBoundingBoxBounds2D(drawingObject, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            return false;
        }

        private static bool TryGetObjectPointBounds2D(
            object drawingObject,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (drawingObject == null)
            {
                return false;
            }

            foreach (string propertyName in new[] { "Points", "Vertices", "Polygon", "ContourPoints", "ControlPoints", "PolygonPoints", "Corners" })
            {
                object value = GetPropertyValue(drawingObject, propertyName);
                if (TryGetNestedPointBounds2D(value, out minX, out minY, out maxX, out maxY))
                {
                    return true;
                }
            }

            foreach (string methodName in new[] { "GetPoints", "GetVertices", "GetPolygon", "GetContourPoints", "GetCorners" })
            {
                object value = InvokeParameterlessMethod(drawingObject, methodName);
                if (TryGetNestedPointBounds2D(value, out minX, out minY, out maxX, out maxY))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetNestedPointBounds2D(
            object value,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (value == null)
            {
                return false;
            }

            if (TryGetPointCollectionBounds2D(value, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            if (TryGetKnownPointPairBounds2D(value, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            foreach (string nestedPropertyName in new[] { "Points", "Vertices", "PolygonPoints", "ContourPoints", "ControlPoints", "Corners" })
            {
                object nestedValue = GetPropertyValue(value, nestedPropertyName);
                if (TryGetPointCollectionBounds2D(nestedValue, out minX, out minY, out maxX, out maxY))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetKnownPointPairBounds2D(
            object source,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (source == null)
            {
                return false;
            }

            foreach (string[] pair in new[]
            {
                new[] { "StartPoint", "EndPoint" },
                new[] { "Point1", "Point2" },
                new[] { "FirstPoint", "SecondPoint" },
                new[] { "LowerLeft", "UpperRight" },
                new[] { "BottomLeft", "TopRight" },
                new[] { "LeftBottom", "RightTop" },
                new[] { "MinimumPoint", "MaximumPoint" },
                new[] { "MinPoint", "MaxPoint" },
                new[] { "Minimum", "Maximum" }
            })
            {
                object first = GetPropertyValue(source, pair[0]);
                object second = GetPropertyValue(source, pair[1]);
                if (TryGetTwoPointBounds2D(first, second, out minX, out minY, out maxX, out maxY))
                {
                    return true;
                }
            }

            if (TryGetCornerSetBounds2D(source, out minX, out minY, out maxX, out maxY))
            {
                return true;
            }

            return false;
        }

        private static bool TryGetCornerSetBounds2D(
            object source,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (source == null)
            {
                return false;
            }

            List<object> corners = new List<object>();
            foreach (string propertyName in new[] { "LowerLeft", "LowerRight", "UpperLeft", "UpperRight", "TopLeft", "TopRight", "BottomLeft", "BottomRight" })
            {
                object point = GetPropertyValue(source, propertyName);
                if (point != null)
                {
                    corners.Add(point);
                }
            }

            return TryGetPointCollectionBounds2D(corners, out minX, out minY, out maxX, out maxY);
        }

        private static bool TryGetTwoPointBounds2D(
            object firstPoint,
            object secondPoint,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            double firstX;
            double firstY;
            double firstZ;
            double secondX;
            double secondY;
            double secondZ;
            if (!TryGetXYZ(firstPoint, out firstX, out firstY, out firstZ)
                || !TryGetXYZ(secondPoint, out secondX, out secondY, out secondZ))
            {
                return false;
            }

            minX = Math.Min(firstX, secondX);
            minY = Math.Min(firstY, secondY);
            maxX = Math.Max(firstX, secondX);
            maxY = Math.Max(firstY, secondY);
            return Math.Abs(maxX - minX) > 1e-6 && Math.Abs(maxY - minY) > 1e-6;
        }

        private static bool TryGetStrictBoundingBoxBounds2D(
            object drawingObject,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
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
                return false;
            }

            return TryGetKnownPointPairBounds2D(bounding, out minX, out minY, out maxX, out maxY);
        }

        private static bool TryGetPointCollectionBounds2D(
            object collection,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (collection == null || collection is string)
            {
                return false;
            }

            IEnumerable enumerable = collection as IEnumerable;
            if (enumerable == null)
            {
                return false;
            }

            bool hasAny = false;
            foreach (object item in enumerable)
            {
                if (item == null)
                {
                    continue;
                }

                double px;
                double py;
                double pz;
                if (!TryGetXYZ(item, out px, out py, out pz))
                {
                    object nestedPoint = GetPropertyValue(item, "Point") ?? GetPropertyValue(item, "Position");
                    if (!TryGetXYZ(nestedPoint, out px, out py, out pz))
                    {
                        continue;
                    }
                }

                if (!hasAny)
                {
                    minX = px;
                    minY = py;
                    maxX = px;
                    maxY = py;
                    hasAny = true;
                }
                else
                {
                    minX = Math.Min(minX, px);
                    minY = Math.Min(minY, py);
                    maxX = Math.Max(maxX, px);
                    maxY = Math.Max(maxY, py);
                }
            }

            return hasAny && Math.Abs(maxX - minX) > 1e-6 && Math.Abs(maxY - minY) > 1e-6;
        }

        private static bool TryFindBestIsometricViewInSheet(
            object sheet,
            out object bestView,
            out int inspectedViews,
            out int candidateViews,
            out string bestViewName,
            out double bestIsoScore)
        {
            bestView = null;
            inspectedViews = 0;
            candidateViews = 0;
            bestViewName = string.Empty;
            bestIsoScore = 0.0;

            if (sheet == null)
            {
                return false;
            }

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects");
            if (allObjects == null)
            {
                return false;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return false;
            }

            double bestScore = double.MinValue;
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

                inspectedViews++;

                string viewName = Convert.ToString(GetPropertyValue(drawingObject, "Name")) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(viewName)
                    && viewName.StartsWith(ExplodedViewPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                List<ExplodedPartPlan> viewPlans = CollectPartPlansFromView(drawingObject);
                int partCount = viewPlans != null ? viewPlans.Count : 0;
                if (partCount <= 0)
                {
                    continue;
                }

                candidateViews++;

                double centerX;
                double centerY;
                double sizeScore;
                if (!TryGetDrawingObjectCenter2D(drawingObject, out centerX, out centerY, out sizeScore))
                {
                    sizeScore = 1.0;
                }

                double isoScore = ComputeIsometricOrientationScore(drawingObject, viewName);
                double rankingScore = (isoScore * 100000.0) + (partCount * 1000.0) + sizeScore;
                if (rankingScore > bestScore)
                {
                    bestScore = rankingScore;
                    bestView = drawingObject;
                    bestViewName = viewName;
                    bestIsoScore = isoScore;
                }
            }

            return bestView != null;
        }

        private static double ComputeIsometricOrientationScore(object viewObject, string viewName)
        {
            double score = 0.0;
            if (viewObject != null)
            {
                object coordinateSystem = GetPropertyValue(viewObject, "ViewCoordinateSystem");
                if (coordinateSystem != null)
                {
                    double ax;
                    double ay;
                    double az;
                    if (TryGetXYZ(GetPropertyValue(coordinateSystem, "AxisX"), out ax, out ay, out az))
                    {
                        score += ComputeAxisSpread(ax, ay, az);
                    }

                    if (TryGetXYZ(GetPropertyValue(coordinateSystem, "AxisY"), out ax, out ay, out az))
                    {
                        score += ComputeAxisSpread(ax, ay, az);
                    }

                    if (TryGetXYZ(GetPropertyValue(coordinateSystem, "AxisZ"), out ax, out ay, out az))
                    {
                        score += ComputeAxisSpread(ax, ay, az) * 0.85;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(viewName)
                && viewName.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 1.25;
            }

            return score;
        }

        private static double ComputeAxisSpread(double x, double y, double z)
        {
            double ax = Math.Abs(x);
            double ay = Math.Abs(y);
            double az = Math.Abs(z);
            double sum = ax + ay + az;
            double dominant = Math.Max(ax, Math.Max(ay, az));
            return Math.Max(0.0, sum - dominant);
        }

        private static bool TryComputeFitIntoRectangle(
            List<ExplodedPartPlan> plans,
            bool ghostLineEnabled,
            double rectMinX,
            double rectMinY,
            double rectMaxX,
            double rectMaxY,
            bool allowGrowth,
            out double fitScaleFactor,
            out double anchorX,
            out double anchorY,
            out double fittedLayoutWidth,
            out double fittedLayoutHeight)
        {
            fitScaleFactor = 1.0;
            anchorX = 0.0;
            anchorY = 0.0;
            fittedLayoutWidth = 0.0;
            fittedLayoutHeight = 0.0;

            double baseMinX;
            double baseMinY;
            double baseMaxX;
            double baseMaxY;
            if (!TryGetPlannedLayoutBounds(
                    plans,
                    ghostLineEnabled,
                    1.0,
                    out baseMinX,
                    out baseMinY,
                    out baseMaxX,
                    out baseMaxY))
            {
                return false;
            }

            double baseWidth = Math.Max(1e-6, baseMaxX - baseMinX);
            double baseHeight = Math.Max(1e-6, baseMaxY - baseMinY);

            double fitPadding = FitRectanglePadding + FitRectangleBorderPadding;
            double targetWidth = Math.Max(1.0, Math.Abs(rectMaxX - rectMinX) - (2.0 * fitPadding));
            double targetHeight = Math.Max(1.0, Math.Abs(rectMaxY - rectMinY) - (2.0 * fitPadding));

            double rawFactor = Math.Min(targetWidth / baseWidth, targetHeight / baseHeight);
            if (!(rawFactor > 0.0))
            {
                rawFactor = 1.0;
            }

            if (!allowGrowth)
            {
                rawFactor = Math.Min(1.0, rawFactor);
            }

            fitScaleFactor = Math.Max(FitScaleFactorMin, Math.Min(FitScaleFactorMax, rawFactor));

            double layoutCenterBaseX = (baseMinX + baseMaxX) * 0.5;
            double layoutCenterBaseY = (baseMinY + baseMaxY) * 0.5;
            double rectCenterX = (rectMinX + rectMaxX) * 0.5;
            double rectCenterY = (rectMinY + rectMaxY) * 0.5;

            anchorX = rectCenterX - (layoutCenterBaseX * fitScaleFactor);
            anchorY = rectCenterY - (layoutCenterBaseY * fitScaleFactor);
            fittedLayoutWidth = baseWidth * fitScaleFactor;
            fittedLayoutHeight = baseHeight * fitScaleFactor;
            return true;
        }

        private static void ScalePlannedOffsets(List<ExplodedPartPlan> plans, double factor)
        {
            if (plans == null || plans.Count == 0)
            {
                return;
            }

            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                plan.OriginalOffsetX *= factor;
                plan.OriginalOffsetY *= factor;
                plan.OffsetX *= factor;
                plan.OffsetY *= factor;
            }
        }

        private static bool TryGetPlannedLayoutBounds(
            List<ExplodedPartPlan> plans,
            bool ghostLineEnabled,
            double visualScaleFactor,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (plans == null || plans.Count == 0)
            {
                return false;
            }

            bool hasAny = false;
            for (int i = 0; i < plans.Count; i++)
            {
                ExplodedPartPlan plan = plans[i];
                double halfSize = GetPlanHalfSize2D(plan, visualScaleFactor);

                if (ghostLineEnabled && !plan.IsMainPart)
                {
                    ExpandBounds(ref hasAny, ref minX, ref minY, ref maxX, ref maxY, plan.OriginalOffsetX, plan.OriginalOffsetY, halfSize);
                    ExpandBounds(ref hasAny, ref minX, ref minY, ref maxX, ref maxY, plan.OffsetX, plan.OffsetY, halfSize);
                }
                else
                {
                    ExpandBounds(ref hasAny, ref minX, ref minY, ref maxX, ref maxY, plan.OffsetX, plan.OffsetY, halfSize);
                }
            }

            return hasAny;
        }

        private static double GetPlanHalfSize2D(ExplodedPartPlan plan, double visualScaleFactor)
        {
            if (plan == null)
            {
                return 2.0;
            }

            double baseSize = plan.ApproxSize2D > 1e-6 ? plan.ApproxSize2D : 8.0;
            double scaled = baseSize * Math.Max(0.01, visualScaleFactor);
            return Math.Max(2.0, scaled * 0.5);
        }

        private static void ExpandBounds(
            ref bool hasAny,
            ref double minX,
            ref double minY,
            ref double maxX,
            ref double maxY,
            double centerX,
            double centerY,
            double halfSize)
        {
            double left = centerX - halfSize;
            double right = centerX + halfSize;
            double bottom = centerY - halfSize;
            double top = centerY + halfSize;

            if (!hasAny)
            {
                minX = left;
                minY = bottom;
                maxX = right;
                maxY = top;
                hasAny = true;
                return;
            }

            minX = Math.Min(minX, left);
            minY = Math.Min(minY, bottom);
            maxX = Math.Max(maxX, right);
            maxY = Math.Max(maxY, top);
        }

        private bool TryFitCreatedViewsIntoRectangle(
            List<object> createdViews,
            Dictionary<string, double> fitViewScales,
            double rectMinX,
            double rectMinY,
            double rectMaxX,
            double rectMaxY,
            double anchorZ,
            out double factorApplied,
            out int adjustedViews)
        {
            factorApplied = 1.0;
            adjustedViews = 0;

            if (createdViews == null || createdViews.Count == 0)
            {
                return false;
            }

            double fitPadding = FitRectanglePadding + FitRectangleBorderPadding;
            double targetWidth = Math.Max(1.0, Math.Abs(rectMaxX - rectMinX) - (2.0 * fitPadding));
            double targetHeight = Math.Max(1.0, Math.Abs(rectMaxY - rectMinY) - (2.0 * fitPadding));
            double targetCenterX = (rectMinX + rectMaxX) * 0.5;
            double targetCenterY = (rectMinY + rectMaxY) * 0.5;

            bool appliedAny = false;
            for (int pass = 0; pass < 8; pass++)
            {
                double currentMinX;
                double currentMinY;
                double currentMaxX;
                double currentMaxY;
                if (!TryGetViewsUnionBounds(createdViews, out currentMinX, out currentMinY, out currentMaxX, out currentMaxY))
                {
                    break;
                }

                double currentWidth = Math.Max(1e-6, currentMaxX - currentMinX);
                double currentHeight = Math.Max(1e-6, currentMaxY - currentMinY);
                double rawFactor = Math.Min(targetWidth / currentWidth, targetHeight / currentHeight);
                if (!(rawFactor > 0.0))
                {
                    break;
                }

                if (rawFactor >= 0.999)
                {
                    break;
                }

                double factor = Math.Max(0.10, Math.Min(0.999, rawFactor * 0.96));
                double currentCenterX = (currentMinX + currentMaxX) * 0.5;
                double currentCenterY = (currentMinY + currentMaxY) * 0.5;

                for (int i = 0; i < createdViews.Count; i++)
                {
                    object view = createdViews[i];
                    if (view == null)
                    {
                        continue;
                    }

                    double viewCenterX;
                    double viewCenterY;
                    double viewSize;
                    if (!TryGetDrawingObjectCenter2D(view, out viewCenterX, out viewCenterY, out viewSize))
                    {
                        continue;
                    }

                    string viewKey = GetDrawingObjectKey(view);
                    double viewScale = 0.0;
                    if (fitViewScales == null
                        || !fitViewScales.TryGetValue(viewKey, out viewScale)
                        || !(viewScale > 0.0))
                    {
                        viewScale = GetSourceViewScale(view);
                    }

                    double newScale = viewScale / factor;
                    TryForceViewScale(view, newScale);
                    if (fitViewScales != null)
                    {
                        fitViewScales[viewKey] = newScale;
                    }

                    double newCenterX = targetCenterX + ((viewCenterX - currentCenterX) * factor);
                    double newCenterY = targetCenterY + ((viewCenterY - currentCenterY) * factor);
                    if (ForceViewCenterToTarget(view, newCenterX, newCenterY, anchorZ))
                    {
                        adjustedViews++;
                    }
                }

                appliedAny = true;
                factorApplied *= factor;
            }

            return appliedAny;
        }

        private static bool TryGetViewsUnionBounds(
            List<object> views,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            minX = 0.0;
            minY = 0.0;
            maxX = 0.0;
            maxY = 0.0;

            if (views == null || views.Count == 0)
            {
                return false;
            }

            bool hasAny = false;
            for (int i = 0; i < views.Count; i++)
            {
                object view = views[i];
                if (view == null)
                {
                    continue;
                }

                double viewMinX;
                double viewMinY;
                double viewMaxX;
                double viewMaxY;
                if (!TryGetDrawingObjectBounds2D(view, out viewMinX, out viewMinY, out viewMaxX, out viewMaxY))
                {
                    continue;
                }

                if (!hasAny)
                {
                    minX = viewMinX;
                    minY = viewMinY;
                    maxX = viewMaxX;
                    maxY = viewMaxY;
                    hasAny = true;
                }
                else
                {
                    minX = Math.Min(minX, viewMinX);
                    minY = Math.Min(minY, viewMinY);
                    maxX = Math.Max(maxX, viewMaxX);
                    maxY = Math.Max(maxY, viewMaxY);
                }
            }

            return hasAny;
        }

        private static string GetDrawingObjectKey(object drawingObject)
        {
            if (drawingObject == null)
            {
                return Guid.NewGuid().ToString("N");
            }

            object rawName = GetPropertyValue(drawingObject, "Name");
            string name = rawName != null ? rawName.ToString() : string.Empty;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            return drawingObject.GetHashCode().ToString();
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

        private static int DeleteExistingGuideLines(object sheet)
        {
            if (sheet == null)
            {
                return 0;
            }
            List<object> linesToDelete = new List<object>();
            CollectHelperOverlaysFromContainer(sheet, linesToDelete);

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects");
            if (allObjects != null)
            {
                MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
                PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
                if (moveNext != null && current != null)
                {
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

                        CollectHelperOverlaysFromContainer(drawingObject, linesToDelete);
                    }
                }
            }

            int removed = 0;
            for (int i = 0; i < linesToDelete.Count; i++)
            {
                bool? deleted = InvokeBoolMethod(linesToDelete[i], "Delete");
                if (deleted.HasValue && deleted.Value)
                {
                    removed++;
                }
            }

            return removed;
        }

        private static int DeleteExistingTestAxes(object sheet)
        {
            if (sheet == null)
            {
                return 0;
            }

            List<object> objectsToDelete = new List<object>();
            CollectTestAxisObjectsFromContainer(sheet, objectsToDelete);

            object allObjects = InvokeParameterlessMethod(sheet, "GetAllObjects");
            if (allObjects != null)
            {
                MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
                PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
                if (moveNext != null && current != null)
                {
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

                        CollectTestAxisObjectsFromContainer(drawingObject, objectsToDelete);
                    }
                }
            }

            int removed = 0;
            for (int i = 0; i < objectsToDelete.Count; i++)
            {
                bool? deleted = InvokeBoolMethod(objectsToDelete[i], "Delete");
                if (deleted.HasValue && deleted.Value)
                {
                    removed++;
                }
            }

            return removed;
        }

        private static void CollectTestAxisObjectsFromContainer(object container, List<object> targets)
        {
            if (container == null || targets == null)
            {
                return;
            }

            object allObjects = InvokeParameterlessMethod(container, "GetAllObjects")
                ?? InvokeParameterlessMethod(container, "GetObjects");
            if (allObjects == null)
            {
                return;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return;
            }

            while (true)
            {
                object moved = moveNext.Invoke(allObjects, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(allObjects, null);
                if (!IsTestAxisObject(drawingObject))
                {
                    continue;
                }

                if (!targets.Contains(drawingObject))
                {
                    targets.Add(drawingObject);
                }
            }
        }

        private static bool IsTestAxisObject(object drawingObject)
        {
            if (drawingObject == null)
            {
                return false;
            }

            if (HasTestAxisNamePrefix(drawingObject))
            {
                return true;
            }

            object attributes = GetPropertyValue(drawingObject, "Attributes");
            return HasTestAxisNamePrefix(attributes);
        }

        private static bool HasTestAxisNamePrefix(object target)
        {
            if (target == null)
            {
                return false;
            }

            object rawName = GetPropertyValue(target, "Name");
            string name = rawName != null ? rawName.ToString() : string.Empty;
            return !string.IsNullOrWhiteSpace(name)
                && name.StartsWith(TestAxisNamePrefix, StringComparison.OrdinalIgnoreCase);
        }

        private static void CollectHelperOverlaysFromContainer(object container, List<object> targets)
        {
            if (container == null || targets == null)
            {
                return;
            }

            object allObjects = InvokeParameterlessMethod(container, "GetAllObjects")
                ?? InvokeParameterlessMethod(container, "GetObjects");
            if (allObjects == null)
            {
                return;
            }

            MethodInfo moveNext = allObjects.GetType().GetMethod("MoveNext", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo current = allObjects.GetType().GetProperty("Current", BindingFlags.Public | BindingFlags.Instance);
            if (moveNext == null || current == null)
            {
                return;
            }

            while (true)
            {
                object moved = moveNext.Invoke(allObjects, null);
                if (!(moved is bool) || !(bool)moved)
                {
                    break;
                }

                object drawingObject = current.GetValue(allObjects, null);
                if (!IsHelperOverlayObject(drawingObject))
                {
                    continue;
                }

                if (!IsGuideLineObject(drawingObject))
                {
                    continue;
                }

                if (!targets.Contains(drawingObject))
                {
                    targets.Add(drawingObject);
                }
            }
        }

        private static bool IsHelperOverlayObject(object drawingObject)
        {
            return IsLineObject(drawingObject) || IsAxisLabelTextObject(drawingObject);
        }

        private static bool IsLineObject(object drawingObject)
        {
            if (drawingObject == null)
            {
                return false;
            }

            Type type = drawingObject.GetType();
            string fullName = type != null ? type.FullName : null;
            if (string.Equals(fullName, "Tekla.Structures.Drawing.Line", StringComparison.Ordinal))
            {
                return true;
            }

            if (string.Equals(fullName, "Tekla.Structures.Drawing.Polyline", StringComparison.Ordinal))
            {
                return true;
            }

            string typeName = type != null ? type.Name : null;
            if (string.Equals(typeName, "Line", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return string.Equals(typeName, "Polyline", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAxisLabelTextObject(object drawingObject)
        {
            if (drawingObject == null)
            {
                return false;
            }

            Type type = drawingObject.GetType();
            string fullName = type != null ? type.FullName : null;
            string typeName = type != null ? type.Name : null;
            bool isText = string.Equals(fullName, "Tekla.Structures.Drawing.Text", StringComparison.Ordinal)
                || string.Equals(typeName, "Text", StringComparison.OrdinalIgnoreCase);
            if (!isText)
            {
                return false;
            }

            object rawText = GetPropertyValue(drawingObject, "TextString");
            string text = rawText != null ? rawText.ToString() : string.Empty;
            if (!IsAxisLabelText(text))
            {
                return false;
            }
            return true;
        }

        private static bool IsAxisLabelText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            switch (text.Trim().ToUpperInvariant())
            {
                case "X+":
                case "X-":
                case "Y+":
                case "Y-":
                case "Z+":
                case "Z-":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsGuideLineObject(object drawingObject)
        {
            if (drawingObject == null)
            {
                return false;
            }

            if (IsAxisLabelTextObject(drawingObject))
            {
                return true;
            }

            if (HasGuideLineNamePrefix(drawingObject))
            {
                return true;
            }

            object attributes = GetPropertyValue(drawingObject, "Attributes");
            if (HasGuideLineNamePrefix(attributes))
            {
                return true;
            }

            return MatchesGuideLineStyle(attributes)
                || MatchesLegacyFitRectangleStyle(attributes)
                || MatchesAxisOverlayLineStyle(attributes);
        }

        private static bool HasGuideLineNamePrefix(object target)
        {
            if (target == null)
            {
                return false;
            }

            object rawName = GetPropertyValue(target, "Name");
            string name = rawName != null ? rawName.ToString() : string.Empty;
            return !string.IsNullOrWhiteSpace(name)
                && (name.StartsWith(GuideLineNamePrefix, StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(FitRectangleNamePrefix, StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(TestAxisNamePrefix, StringComparison.OrdinalIgnoreCase));
        }

        private static bool MatchesGuideLineStyle(object attributes)
        {
            if (attributes == null)
            {
                return false;
            }

            object lineTypeAttributes = GetPropertyValue(attributes, "Line");
            if (lineTypeAttributes == null)
            {
                return false;
            }

            int colorIndex;
            if (!TryReadGuideLineColorIndex(lineTypeAttributes, out colorIndex))
            {
                return false;
            }

            if (!IsAxisOverlayColor(colorIndex))
            {
                return false;
            }

            object lineTypeValue = GetPropertyValue(lineTypeAttributes, "Type")
                ?? GetPropertyValue(lineTypeAttributes, "LineType");
            if (lineTypeValue == null)
            {
                return false;
            }

            string lineTypeText = lineTypeValue.ToString();
            return !string.IsNullOrWhiteSpace(lineTypeText)
                && lineTypeText.IndexOf("dashed", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool MatchesAxisOverlayLineStyle(object attributes)
        {
            if (attributes == null)
            {
                return false;
            }

            object lineTypeAttributes = GetPropertyValue(attributes, "Line");
            if (lineTypeAttributes == null)
            {
                return false;
            }

            int colorIndex;
            if (!TryReadGuideLineColorIndex(lineTypeAttributes, out colorIndex) || !IsAxisOverlayColor(colorIndex))
            {
                return false;
            }

            object lineTypeValue = GetPropertyValue(lineTypeAttributes, "Type")
                ?? GetPropertyValue(lineTypeAttributes, "LineType");
            string lineTypeText = lineTypeValue != null ? lineTypeValue.ToString() : string.Empty;

            if (string.IsNullOrWhiteSpace(lineTypeText))
            {
                return true;
            }

            return lineTypeText.IndexOf("dashed", StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static bool IsAxisOverlayColor(int colorIndex)
        {
            return colorIndex == OtherPlaneContourColor
                || colorIndex == SideNegativeContourColor
                || colorIndex == SidePositiveContourColor;
        }

        private static bool MatchesLegacyFitRectangleStyle(object attributes)
        {
            if (attributes == null)
            {
                return false;
            }

            object lineTypeAttributes = GetPropertyValue(attributes, "Line");
            if (lineTypeAttributes == null)
            {
                return false;
            }

            int colorIndex;
            if (!TryReadGuideLineColorIndex(lineTypeAttributes, out colorIndex))
            {
                return false;
            }

            return colorIndex == FitRectangleColor;
        }

        private static bool TryReadGuideLineColorIndex(object lineTypeAttributes, out int colorIndex)
        {
            colorIndex = 0;
            if (lineTypeAttributes == null)
            {
                return false;
            }

            object color = GetPropertyValue(lineTypeAttributes, "Color");
            if (TryConvertColorObjectToInt(color, out colorIndex))
            {
                return true;
            }

            object trueColor = GetPropertyValue(lineTypeAttributes, "TrueColor");
            return TryConvertColorObjectToInt(trueColor, out colorIndex);
        }

        private static bool TryConvertColorObjectToInt(object colorObject, out int colorIndex)
        {
            colorIndex = 0;
            if (colorObject == null)
            {
                return false;
            }

            Type valueType = colorObject.GetType();
            if (valueType.IsEnum)
            {
                colorIndex = Convert.ToInt32(colorObject);
                return true;
            }

            if (colorObject is int)
            {
                colorIndex = (int)colorObject;
                return true;
            }

            if (colorObject is short)
            {
                colorIndex = (short)colorObject;
                return true;
            }

            if (colorObject is byte)
            {
                colorIndex = (byte)colorObject;
                return true;
            }

            if (colorObject is long)
            {
                colorIndex = (int)(long)colorObject;
                return true;
            }

            object nestedColor = GetPropertyValue(colorObject, "Color");
            if (nestedColor != null && !ReferenceEquals(nestedColor, colorObject))
            {
                return TryConvertColorObjectToInt(nestedColor, out colorIndex);
            }

            return false;
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
                plans[i].OffsetFromXAxisX = 0.0;
                plans[i].OffsetFromXAxisY = 0.0;
                plans[i].OffsetFromYAxisX = 0.0;
                plans[i].OffsetFromYAxisY = 0.0;
                plans[i].OffsetFromZAxisX = 0.0;
                plans[i].OffsetFromZAxisY = 0.0;
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
            bool useOverrideDisplayAxes,
            double overrideAxisXx,
            double overrideAxisXy,
            double overrideAxisXz,
            double overrideAxisYx,
            double overrideAxisYy,
            double overrideAxisYz,
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
                useOverrideDisplayAxes,
                overrideAxisXx,
                overrideAxisXy,
                overrideAxisXz,
                overrideAxisYx,
                overrideAxisYy,
                overrideAxisYz,
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
            bool useOverrideDisplayAxes,
            double overrideAxisXx,
            double overrideAxisXy,
            double overrideAxisXz,
            double overrideAxisYx,
            double overrideAxisYy,
            double overrideAxisYz,
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
            if (useOverrideDisplayAxes)
            {
                axisXx = overrideAxisXx;
                axisXy = overrideAxisXy;
                axisXz = overrideAxisXz;
                axisYx = overrideAxisYx;
                axisYy = overrideAxisYy;
                axisYz = overrideAxisYz;
            }
            else if (!TryGetViewDisplayAxes(sourceView, out axisXx, out axisXy, out axisXz, out axisYx, out axisYy, out axisYz))
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
                bool canUseDrawing2DBase = !useOverrideDisplayAxes && plan.HasCenter2D && mainPlan.HasCenter2D;
                if (canUseDrawing2DBase)
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
                double xAxisOffsetX = 0.0;
                double xAxisOffsetY = 0.0;
                double yAxisOffsetX = 0.0;
                double yAxisOffsetY = 0.0;
                double zAxisOffsetX = 0.0;
                double zAxisOffsetY = 0.0;

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
                        double componentX = amount * sideOnZY * axisXSheetX;
                        double componentY = amount * sideOnZY * axisXSheetY;
                        xAxisOffsetX += componentX;
                        xAxisOffsetY += componentY;
                        explodeOffsetX += componentX;
                        explodeOffsetY += componentY;
                        movedByZY++;
                    }

                    if (usePlaneXY)
                    {
                        if (sideOnXY != 0)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalZFromMain) / maxAbsLocalZ);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            double componentX = amount * sideOnXY * axisZSheetX;
                            double componentY = amount * sideOnXY * axisZSheetY;
                            zAxisOffsetX += componentX;
                            zAxisOffsetY += componentY;
                            explodeOffsetX += componentX;
                            explodeOffsetY += componentY;
                            movedByXY++;
                        }
                    }

                    if (usePlaneXZ)
                    {
                        if (sideOnXZ != 0)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalYFromMain) / maxAbsLocalY);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            double componentX = amount * sideOnXZ * axisYSheetX;
                            double componentY = amount * sideOnXZ * axisYSheetY;
                            yAxisOffsetX += componentX;
                            yAxisOffsetY += componentY;
                            explodeOffsetX += componentX;
                            explodeOffsetY += componentY;
                            movedByXZ++;
                        }
                    }

                    if (usePlaneZY)
                    {
                        if (sideOnZY != 0 && !plan.IsOtherPlane)
                        {
                            double normalizedDepth = Math.Min(1.0, Math.Abs(plan.LocalXFromMain) / maxAbsLocalX);
                            double amount = baseExplodeSheet + (gainExplodeSheet * normalizedDepth);
                            double componentX = amount * sideOnZY * axisXSheetX;
                            double componentY = amount * sideOnZY * axisXSheetY;
                            xAxisOffsetX += componentX;
                            xAxisOffsetY += componentY;
                            explodeOffsetX += componentX;
                            explodeOffsetY += componentY;
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
                plan.OffsetFromXAxisX = xAxisOffsetX;
                plan.OffsetFromXAxisY = xAxisOffsetY;
                plan.OffsetFromYAxisX = yAxisOffsetX;
                plan.OffsetFromYAxisY = yAxisOffsetY;
                plan.OffsetFromZAxisX = zAxisOffsetX;
                plan.OffsetFromZAxisY = zAxisOffsetY;
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
                plan.OffsetX = plan.OriginalOffsetX;
                plan.OffsetY = plan.OriginalOffsetY;
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

        private object CreateSinglePartViewForFit(
            object sheet,
            object viewCoordinateSystem,
            object displayCoordinateSystem,
            object identifier,
            double sourceScale,
            string viewName,
            out bool autoIsoApplied)
        {
            ArrayList partList = new ArrayList();
            partList.Add(identifier);

            return CreatePartListViewForFit(
                sheet,
                viewCoordinateSystem,
                displayCoordinateSystem,
                partList,
                sourceScale,
                viewName,
                out autoIsoApplied);
        }

        private object CreatePartListViewForFit(
            object sheet,
            object viewCoordinateSystem,
            object displayCoordinateSystem,
            ArrayList partList,
            double sourceScale,
            string viewName,
            out bool autoIsoApplied)
        {
            autoIsoApplied = false;

            if (partList == null || partList.Count == 0)
            {
                return null;
            }

            Type viewType = ResolveTeklaType(
                "Tekla.Structures.Drawing.View",
                "Tekla.Structures.Drawing",
                TeklaDrawingDllCandidates);
            if (viewType == null)
            {
                return null;
            }

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
                // Neste fluxo carregamos o perfil para manter a formatacao visual e,
                // em seguida, forcamos a escala de fit.
                autoIsoApplied = TryLoadAttributesByName(viewAttributes, ExplodedViewAttributeFile);
                if (!autoIsoApplied)
                {
                    autoIsoApplied = TryLoadAttributesByName(view, ExplodedViewAttributeFile);
                }

                SetPropertyValue(viewAttributes, "Scale", sourceScale);
                SetPropertyValue(viewAttributes, "FixedViewPlacing", true);
                SetPropertyValue(view, "Attributes", viewAttributes, "Tekla.Structures.Drawing.View");
            }

            bool? inserted = InvokeBoolMethod(view, "Insert");
            if (!inserted.HasValue || !inserted.Value)
            {
                return null;
            }

            TryForceViewScale(view, sourceScale);

            return view;
        }

        private static bool TryForceViewScale(object view, double scale)
        {
            if (view == null || !(scale > 0.0))
            {
                return false;
            }

            bool changed = false;
            object viewAttributes = GetPropertyValue(view, "Attributes", "Tekla.Structures.Drawing.View")
                ?? GetPropertyValue(view, "Attributes");
            if (viewAttributes != null)
            {
                changed |= SetPropertyValue(viewAttributes, "Scale", scale);
                changed |= SetPropertyValue(viewAttributes, "ScaleX", scale);
                changed |= SetPropertyValue(viewAttributes, "ScaleY", scale);
                changed |= SetPropertyValue(view, "Attributes", viewAttributes, "Tekla.Structures.Drawing.View")
                    || SetPropertyValue(view, "Attributes", viewAttributes);
            }

            changed |= SetPropertyValue(view, "Scale", scale);

            if (changed)
            {
                InvokeBoolMethod(view, "Modify");
            }

            return changed;
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
            RotateViewToAngles(view, -35.264, 45.0);
        }

        private static void RotateViewToAngles(object view, double rotationXDegrees, double rotationYDegrees)
        {
            InvokeMethod(view, "RotateViewOnAxisX", rotationXDegrees);
            InvokeMethod(view, "RotateViewOnAxisY", rotationYDegrees);
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

        private static object CreateVector(double x, double y, double z)
        {
            Type vectorType = ResolveTeklaType(
                "Tekla.Structures.Geometry3d.Vector",
                "Tekla.Structures",
                TeklaStructuresDllCandidates);
            if (vectorType == null)
            {
                return null;
            }

            try
            {
                return Activator.CreateInstance(vectorType, new object[] { x, y, z });
            }
            catch
            {
                return null;
            }
        }

        private static object CreateGlobalCoordinateSystem()
        {
            return CreateCoordinateSystem(
                0.0,
                0.0,
                0.0,
                1.0,
                0.0,
                0.0,
                0.0,
                1.0,
                0.0);
        }

        private static object CreateCoordinateSystem(
            double originX,
            double originY,
            double originZ,
            double axisXx,
            double axisXy,
            double axisXz,
            double axisYx,
            double axisYy,
            double axisYz)
        {
            Type coordinateSystemType = ResolveTeklaType(
                "Tekla.Structures.Geometry3d.CoordinateSystem",
                "Tekla.Structures",
                TeklaStructuresDllCandidates);
            if (coordinateSystemType == null)
            {
                return null;
            }

            object origin = CreatePoint(originX, originY, originZ);
            object axisX = CreateVector(axisXx, axisXy, axisXz) ?? CreatePoint(axisXx, axisXy, axisXz);
            object axisY = CreateVector(axisYx, axisYy, axisYz) ?? CreatePoint(axisYx, axisYy, axisYz);
            if (origin == null || axisX == null || axisY == null)
            {
                return null;
            }

            try
            {
                return Activator.CreateInstance(coordinateSystemType, new object[] { origin, axisX, axisY });
            }
            catch
            {
                try
                {
                    object coordinateSystem = Activator.CreateInstance(coordinateSystemType);
                    if (coordinateSystem == null)
                    {
                        return null;
                    }

                    SetPropertyValue(coordinateSystem, "Origin", origin);
                    SetPropertyValue(coordinateSystem, "AxisX", axisX);
                    SetPropertyValue(coordinateSystem, "AxisY", axisY);
                    return coordinateSystem;
                }
                catch
                {
                    return null;
                }
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
            public double OffsetFromXAxisX;
            public double OffsetFromXAxisY;
            public double OffsetFromYAxisX;
            public double OffsetFromYAxisY;
            public double OffsetFromZAxisX;
            public double OffsetFromZAxisY;
            public double OffsetX;
            public double OffsetY;
        }
    }
}



