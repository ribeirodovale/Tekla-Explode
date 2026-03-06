using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public sealed class OrbitToolForm : Form
    {
        public const double DefaultPreviewYawRadians = 0.0;
        public const double DefaultPreviewPitchRadians = 0.0;

        public sealed class PreviewTriangle
        {
            public PreviewTriangle(
                double x1,
                double y1,
                double z1,
                double x2,
                double y2,
                double z2,
                double x3,
                double y3,
                double z3)
            {
                X1 = x1;
                Y1 = y1;
                Z1 = z1;
                X2 = x2;
                Y2 = y2;
                Z2 = z2;
                X3 = x3;
                Y3 = y3;
                Z3 = z3;
            }

            public double X1 { get; private set; }
            public double Y1 { get; private set; }
            public double Z1 { get; private set; }
            public double X2 { get; private set; }
            public double Y2 { get; private set; }
            public double Z2 { get; private set; }
            public double X3 { get; private set; }
            public double Y3 { get; private set; }
            public double Z3 { get; private set; }
        }

        public sealed class PreviewPart
        {
            public PreviewPart(
                string label,
                double minX,
                double minY,
                double minZ,
                double maxX,
                double maxY,
                double maxZ,
                bool isMainPart,
                IList<PreviewTriangle> triangles = null)
            {
                Label = string.IsNullOrWhiteSpace(label) ? "(sem id)" : label;
                MinX = minX;
                MinY = minY;
                MinZ = minZ;
                MaxX = maxX;
                MaxY = maxY;
                MaxZ = maxZ;
                IsMainPart = isMainPart;
                Triangles = triangles != null
                    ? new List<PreviewTriangle>(triangles)
                    : new List<PreviewTriangle>();
            }

            public string Label { get; private set; }

            public double MinX { get; private set; }

            public double MinY { get; private set; }

            public double MinZ { get; private set; }

            public double MaxX { get; private set; }

            public double MaxY { get; private set; }

            public double MaxZ { get; private set; }

            public bool IsMainPart { get; set; }

            public List<PreviewTriangle> Triangles { get; private set; }

            public double CenterX
            {
                get { return (MinX + MaxX) * 0.5; }
            }

            public double CenterY
            {
                get { return (MinY + MaxY) * 0.5; }
            }

            public double CenterZ
            {
                get { return (MinZ + MaxZ) * 0.5; }
            }
        }

        public OrbitToolForm(IList<PreviewPart> parts, string summaryText)
        {
            Text = "Ferramenta de orbita - Preview 3D";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(900, 620);
            MinimumSize = new Size(760, 520);
            BackColor = Color.FromArgb(238, 240, 244);

            Control previewSurface = CreatePreviewSurface(parts, summaryText);
            previewSurface.Dock = DockStyle.Fill;
            Controls.Add(previewSurface);
        }

        public static PreviewSurfaceControl CreatePreviewSurface(IList<PreviewPart> parts, string summaryText)
        {
            return new PreviewSurfaceControl(parts, summaryText);
        }

        public sealed class PreviewSurfaceControl : Panel
        {
            private readonly OrbitPreviewViewportControl viewport;

            public PreviewSurfaceControl(IList<PreviewPart> parts, string summaryText)
            {
                BackColor = Color.FromArgb(238, 240, 244);
                Margin = new Padding(0);
                Padding = new Padding(0);

                Panel headerPanel = new Panel();
                headerPanel.Dock = DockStyle.Top;
                headerPanel.Height = 64;
                headerPanel.BackColor = Color.FromArgb(245, 246, 248);

                Label lblTitle = new Label();
                lblTitle.AutoSize = true;
                lblTitle.Location = new Point(14, 10);
                lblTitle.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
                lblTitle.ForeColor = Color.FromArgb(48, 55, 68);
                lblTitle.Text = "Representacao 3D da vista selecionada";

                Label lblSummary = new Label();
                lblSummary.Location = new Point(14, 31);
                lblSummary.Size = new Size(730, 26);
                lblSummary.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point, 0);
                lblSummary.ForeColor = Color.FromArgb(86, 94, 109);
                lblSummary.Text = string.IsNullOrWhiteSpace(summaryText) ? "Sem dados." : summaryText;

                OrbitGizmoControl gizmo = new OrbitGizmoControl();
                gizmo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                gizmo.Location = new Point(812, 8);
                gizmo.Size = new Size(72, 50);

                headerPanel.Controls.Add(lblTitle);
                headerPanel.Controls.Add(lblSummary);
                headerPanel.Controls.Add(gizmo);

                viewport = new OrbitPreviewViewportControl(parts);
                viewport.Dock = DockStyle.Fill;

                Label lblHelp = new Label();
                lblHelp.Dock = DockStyle.Bottom;
                lblHelp.Height = 24;
                lblHelp.TextAlign = ContentAlignment.MiddleCenter;
                lblHelp.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point, 0);
                lblHelp.ForeColor = Color.FromArgb(80, 87, 100);
                lblHelp.Text = "Mouse esquerdo: orbita  |  Mouse direito: pan  |  Roda: zoom  |  Duplo clique: reset";

                Controls.Add(viewport);
                Controls.Add(lblHelp);
                Controls.Add(headerPanel);
            }

            public bool TryGetCameraAngles(out double yawRadians, out double pitchRadians)
            {
                if (viewport == null)
                {
                    yawRadians = 0.0;
                    pitchRadians = 0.0;
                    return false;
                }

                return viewport.TryGetCameraAngles(out yawRadians, out pitchRadians);
            }
        }

        private sealed class OrbitPreviewViewportControl : Control
        {
            private readonly bool useOrthographicProjection = true;

            private static readonly int[,] BoxEdges =
            {
                { 0, 1 }, { 1, 2 }, { 2, 3 }, { 3, 0 },
                { 4, 5 }, { 5, 6 }, { 6, 7 }, { 7, 4 },
                { 0, 4 }, { 1, 5 }, { 2, 6 }, { 3, 7 }
            };

            private static readonly int[,] BoxFaces =
            {
                { 0, 3, 2, 1 }, // bottom (-Z)
                { 4, 5, 6, 7 }, // top
                { 0, 1, 5, 4 }, // front
                { 2, 3, 7, 6 }, // back
                { 1, 2, 6, 5 }, // right
                { 0, 4, 7, 3 }  // left (-X)
            };

            private readonly List<PreviewPart> parts = new List<PreviewPart>();

            private MouseButtons dragButton = MouseButtons.None;
            private Point lastMouse;

            private double yawRadians = DefaultPreviewYawRadians;
            private double pitchRadians = DefaultPreviewPitchRadians;
            private float panPixelsX;
            private float panPixelsY;
            private double cameraDistance = 3000.0;
            private double minCameraDistance = 10.0;
            private double maxCameraDistance = 100000.0;

            private bool hasSceneBounds;
            private double sceneMinX;
            private double sceneMinY;
            private double sceneMinZ;
            private double sceneMaxX;
            private double sceneMaxY;
            private double sceneMaxZ;
            private double sceneCenterX;
            private double sceneCenterY;
            private double sceneCenterZ;
            private double sceneExtent = 1000.0;

            public OrbitPreviewViewportControl(IList<PreviewPart> sourceParts)
            {
                SetStyle(
                    ControlStyles.UserPaint
                    | ControlStyles.AllPaintingInWmPaint
                    | ControlStyles.OptimizedDoubleBuffer
                    | ControlStyles.ResizeRedraw,
                    true);
                BackColor = Color.FromArgb(229, 232, 236);
                TabStop = true;

                if (sourceParts != null)
                {
                    for (int i = 0; i < sourceParts.Count; i++)
                    {
                        PreviewPart part = sourceParts[i];
                        if (part != null)
                        {
                            parts.Add(part);
                        }
                    }
                }

                RebuildSceneBounds();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (LinearGradientBrush background = new LinearGradientBrush(
                    ClientRectangle,
                    Color.FromArgb(238, 240, 243),
                    Color.FromArgb(220, 223, 228),
                    90f))
                {
                    g.FillRectangle(background, ClientRectangle);
                }

                DrawReferenceGrid(g);

                if (parts.Count == 0 || !hasSceneBounds)
                {
                    DrawEmptyMessage(g);
                    return;
                }

                DrawSceneAxes(g);
                DrawPartBoxes(g);

                using (SolidBrush marker = new SolidBrush(Color.FromArgb(220, 42, 47, 57)))
                {
                    g.FillEllipse(
                        marker,
                        (ClientSize.Width / 2f) + panPixelsX - 2f,
                        (ClientSize.Height / 2f) + panPixelsY - 2f,
                        4f,
                        4f);
                }
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);

                if (e.Button == MouseButtons.Left || e.Button == MouseButtons.Right)
                {
                    dragButton = e.Button;
                    lastMouse = e.Location;
                    Focus();
                }
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);

                if (dragButton == MouseButtons.None)
                {
                    return;
                }

                int dx = e.X - lastMouse.X;
                int dy = e.Y - lastMouse.Y;
                lastMouse = e.Location;

                if (dragButton == MouseButtons.Left)
                {
                    yawRadians += dx * 0.010;
                    pitchRadians += dy * 0.010;

                    if (pitchRadians < -1.45)
                    {
                        pitchRadians = -1.45;
                    }
                    else if (pitchRadians > 1.45)
                    {
                        pitchRadians = 1.45;
                    }
                }
                else if (dragButton == MouseButtons.Right)
                {
                    panPixelsX += dx;
                    panPixelsY += dy;
                }

                Invalidate();
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                base.OnMouseUp(e);
                dragButton = MouseButtons.None;
            }

            protected override void OnMouseWheel(MouseEventArgs e)
            {
                base.OnMouseWheel(e);

                double zoomFactor = e.Delta > 0 ? 0.88 : 1.14;
                cameraDistance *= zoomFactor;
                ClampCameraDistance();
                Invalidate();
            }

            protected override void OnDoubleClick(EventArgs e)
            {
                base.OnDoubleClick(e);
                ResetCamera();
                Invalidate();
            }

            public bool TryGetCameraAngles(out double yawRadians, out double pitchRadians)
            {
                yawRadians = this.yawRadians;
                pitchRadians = this.pitchRadians;
                return true;
            }

            private void ResetCamera()
            {
                yawRadians = DefaultPreviewYawRadians;
                pitchRadians = DefaultPreviewPitchRadians;
                panPixelsX = 0f;
                panPixelsY = 0f;
                cameraDistance = Math.Max(sceneExtent * 2.8, 80.0);
                ClampCameraDistance();
            }

            private void RebuildSceneBounds()
            {
                hasSceneBounds = false;
                sceneMinX = 0.0;
                sceneMinY = 0.0;
                sceneMinZ = 0.0;
                sceneMaxX = 0.0;
                sceneMaxY = 0.0;
                sceneMaxZ = 0.0;
                sceneExtent = 1000.0;
                sceneCenterX = 0.0;
                sceneCenterY = 0.0;
                sceneCenterZ = 0.0;

                for (int i = 0; i < parts.Count; i++)
                {
                    PreviewPart part = parts[i];
                    if (!hasSceneBounds)
                    {
                        sceneMinX = part.MinX;
                        sceneMinY = part.MinY;
                        sceneMinZ = part.MinZ;
                        sceneMaxX = part.MaxX;
                        sceneMaxY = part.MaxY;
                        sceneMaxZ = part.MaxZ;
                        hasSceneBounds = true;
                        continue;
                    }

                    sceneMinX = Math.Min(sceneMinX, part.MinX);
                    sceneMinY = Math.Min(sceneMinY, part.MinY);
                    sceneMinZ = Math.Min(sceneMinZ, part.MinZ);
                    sceneMaxX = Math.Max(sceneMaxX, part.MaxX);
                    sceneMaxY = Math.Max(sceneMaxY, part.MaxY);
                    sceneMaxZ = Math.Max(sceneMaxZ, part.MaxZ);
                }

                if (!hasSceneBounds)
                {
                    minCameraDistance = 10.0;
                    maxCameraDistance = 100000.0;
                    cameraDistance = 600.0;
                    return;
                }

                sceneCenterX = (sceneMinX + sceneMaxX) * 0.5;
                sceneCenterY = (sceneMinY + sceneMaxY) * 0.5;
                sceneCenterZ = (sceneMinZ + sceneMaxZ) * 0.5;

                double sizeX = sceneMaxX - sceneMinX;
                double sizeY = sceneMaxY - sceneMinY;
                double sizeZ = sceneMaxZ - sceneMinZ;
                sceneExtent = Math.Max(1.0, Math.Max(sizeX, Math.Max(sizeY, sizeZ)));

                minCameraDistance = Math.Max(5.0, sceneExtent * 0.20);
                maxCameraDistance = Math.Max(5000.0, sceneExtent * 40.0);
                cameraDistance = Math.Max(sceneExtent * 2.8, 80.0);
                ClampCameraDistance();
            }

            private void ClampCameraDistance()
            {
                if (cameraDistance < minCameraDistance)
                {
                    cameraDistance = minCameraDistance;
                }
                else if (cameraDistance > maxCameraDistance)
                {
                    cameraDistance = maxCameraDistance;
                }
            }

            private void DrawEmptyMessage(Graphics g)
            {
                const string message = "Sem dados 3D para exibir.\nSelecione uma vista no desenho e abra novamente.";
                using (StringFormat format = new StringFormat())
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(88, 95, 108)))
                {
                    format.Alignment = StringAlignment.Center;
                    format.LineAlignment = StringAlignment.Center;
                    g.DrawString(message, Font, brush, ClientRectangle, format);
                }
            }

            private void DrawReferenceGrid(Graphics g)
            {
                if (!hasSceneBounds)
                {
                    return;
                }

                double gridSize = sceneExtent * 0.60;
                double step = Math.Max(sceneExtent / 12.0, 1.0);
                int lineCount = 10;
                using (Pen gridPen = new Pen(Color.FromArgb(72, 130, 138, 150), 1f))
                {
                    for (int i = -lineCount; i <= lineCount; i++)
                    {
                        double offset = i * step;
                        DrawProjectedSegment(
                            g,
                            new Vec3(sceneCenterX - gridSize, sceneCenterY + offset, sceneCenterZ),
                            new Vec3(sceneCenterX + gridSize, sceneCenterY + offset, sceneCenterZ),
                            gridPen);
                        DrawProjectedSegment(
                            g,
                            new Vec3(sceneCenterX + offset, sceneCenterY - gridSize, sceneCenterZ),
                            new Vec3(sceneCenterX + offset, sceneCenterY + gridSize, sceneCenterZ),
                            gridPen);
                    }
                }
            }

            private void DrawSceneAxes(Graphics g)
            {
                double axisLength = Math.Max(sceneExtent * 0.70, 30.0);
                Vec3 center = new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ);

                using (Pen xPen = new Pen(Color.FromArgb(196, 43, 45), 2f))
                using (Pen yPen = new Pen(Color.FromArgb(38, 168, 72), 2f))
                using (Pen zPen = new Pen(Color.FromArgb(52, 88, 214), 2f))
                using (SolidBrush xBrush = new SolidBrush(Color.FromArgb(196, 43, 45)))
                using (SolidBrush yBrush = new SolidBrush(Color.FromArgb(38, 168, 72)))
                using (SolidBrush zBrush = new SolidBrush(Color.FromArgb(52, 88, 214)))
                using (Font axisFont = new Font("Segoe UI", 8.5f, FontStyle.Bold, GraphicsUnit.Point, 0))
                {
                    DrawProjectedAxisWithLabels(
                        g,
                        center,
                        new Vec3(sceneCenterX - axisLength, sceneCenterY, sceneCenterZ),
                        new Vec3(sceneCenterX + axisLength, sceneCenterY, sceneCenterZ),
                        xPen);
                    DrawProjectedAxisWithLabels(
                        g,
                        center,
                        new Vec3(sceneCenterX, sceneCenterY - axisLength, sceneCenterZ),
                        new Vec3(sceneCenterX, sceneCenterY + axisLength, sceneCenterZ),
                        yPen);
                    DrawProjectedAxisWithLabels(
                        g,
                        center,
                        new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ - axisLength),
                        new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ + axisLength),
                        zPen);

                    DrawAxisLabel(g, axisFont, xBrush, "X-", center, new Vec3(sceneCenterX - axisLength, sceneCenterY, sceneCenterZ));
                    DrawAxisLabel(g, axisFont, xBrush, "X+", center, new Vec3(sceneCenterX + axisLength, sceneCenterY, sceneCenterZ));
                    DrawAxisLabel(g, axisFont, yBrush, "Y-", center, new Vec3(sceneCenterX, sceneCenterY - axisLength, sceneCenterZ));
                    DrawAxisLabel(g, axisFont, yBrush, "Y+", center, new Vec3(sceneCenterX, sceneCenterY + axisLength, sceneCenterZ));
                    DrawAxisLabel(g, axisFont, zBrush, "Z-", center, new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ - axisLength));
                    DrawAxisLabel(g, axisFont, zBrush, "Z+", center, new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ + axisLength));
                }
            }

            private void DrawProjectedAxisWithLabels(
                Graphics g,
                Vec3 center,
                Vec3 from,
                Vec3 to,
                Pen pen)
            {
                PointF centerScreen;
                PointF fromScreen;
                PointF toScreen;
                double centerDepth;
                double fromDepth;
                double toDepth;
                bool hasCenter = TryProject(center, out centerScreen, out centerDepth);
                bool hasFrom = TryProject(from, out fromScreen, out fromDepth);
                bool hasTo = TryProject(to, out toScreen, out toDepth);

                if (hasFrom && hasTo)
                {
                    g.DrawLine(pen, fromScreen, toScreen);
                    return;
                }

                if (hasCenter && hasFrom)
                {
                    g.DrawLine(pen, centerScreen, fromScreen);
                }

                if (hasCenter && hasTo)
                {
                    g.DrawLine(pen, centerScreen, toScreen);
                }
            }

            private void DrawAxisLabel(
                Graphics g,
                Font font,
                Brush brush,
                string label,
                Vec3 center,
                Vec3 target)
            {
                PointF centerScreen;
                PointF targetScreen;
                double centerDepth;
                double targetDepth;
                if (!TryProject(center, out centerScreen, out centerDepth)
                    || !TryProject(target, out targetScreen, out targetDepth))
                {
                    return;
                }

                float dirX = targetScreen.X - centerScreen.X;
                float dirY = targetScreen.Y - centerScreen.Y;
                float length = (float)Math.Sqrt((dirX * dirX) + (dirY * dirY));
                if (length < 1e-3f)
                {
                    return;
                }

                dirX /= length;
                dirY /= length;

                float outwardOffset = Math.Max(9f, Math.Min(18f, length * 0.08f));
                float labelX = targetScreen.X + (dirX * outwardOffset);
                float labelY = targetScreen.Y + (dirY * outwardOffset);

                SizeF size = g.MeasureString(label, font);
                float drawX = labelX - (size.Width * 0.5f);
                float drawY = labelY - (size.Height * 0.5f);

                using (SolidBrush halo = new SolidBrush(Color.FromArgb(180, 245, 247, 250)))
                {
                    g.FillRectangle(
                        halo,
                        drawX - 1.5f,
                        drawY - 0.5f,
                        size.Width + 3f,
                        size.Height + 1f);
                }

                g.DrawString(label, font, brush, drawX, drawY);
            }

            private void DrawPartBoxes(Graphics g)
            {
                List<BoxFaceRenderInfo> allFaces = new List<BoxFaceRenderInfo>();
                for (int i = 0; i < parts.Count; i++)
                {
                    PreviewPart part = parts[i];
                    if (part != null && part.Triangles != null && part.Triangles.Count > 0)
                    {
                        CollectPartTriangleFaces(part, allFaces);
                    }
                    else
                    {
                        CollectPartBoxFaces(part, allFaces);
                    }
                }

                RenderDepthBufferedFaces(g, allFaces);
            }

            private void CollectPartBoxFaces(PreviewPart part, List<BoxFaceRenderInfo> targetFaces)
            {
                if (part == null || targetFaces == null)
                {
                    return;
                }

                Vec3[] corners =
                {
                    new Vec3(part.MinX, part.MinY, part.MinZ),
                    new Vec3(part.MaxX, part.MinY, part.MinZ),
                    new Vec3(part.MaxX, part.MaxY, part.MinZ),
                    new Vec3(part.MinX, part.MaxY, part.MinZ),
                    new Vec3(part.MinX, part.MinY, part.MaxZ),
                    new Vec3(part.MaxX, part.MinY, part.MaxZ),
                    new Vec3(part.MaxX, part.MaxY, part.MaxZ),
                    new Vec3(part.MinX, part.MaxY, part.MaxZ)
                };

                PointF[] projected = new PointF[corners.Length];
                bool[] visible = new bool[corners.Length];
                double[] projectedDepths = new double[corners.Length];
                Vec3[] cameraCorners = new Vec3[corners.Length];
                int visibleCount = 0;

                for (int i = 0; i < corners.Length; i++)
                {
                    if (!TryTransformToCamera(corners[i], out cameraCorners[i]))
                    {
                        visible[i] = false;
                        projected[i] = PointF.Empty;
                        projectedDepths[i] = 0.0;
                        continue;
                    }

                    visible[i] = TryProjectCamera(cameraCorners[i], out projected[i], out projectedDepths[i]);
                    if (visible[i])
                    {
                        visibleCount++;
                    }
                }

                if (visibleCount < 2)
                {
                    return;
                }

                Color edgeColor = part.IsMainPart
                    ? Color.FromArgb(92, 64, 8)
                    : Color.FromArgb(18, 28, 46);
                float edgeThickness = part.IsMainPart ? 2.6f : 1.5f;
                Color fillColor = part.IsMainPart
                    ? Color.FromArgb(210, 164, 32)
                    : Color.FromArgb(55, 79, 122);

                for (int face = 0; face < BoxFaces.GetLength(0); face++)
                {
                    int a = BoxFaces[face, 0];
                    int b = BoxFaces[face, 1];
                    int c = BoxFaces[face, 2];
                    int d = BoxFaces[face, 3];
                    if (!visible[a] || !visible[b] || !visible[c] || !visible[d])
                    {
                        continue;
                    }

                    Vec3 pa = cameraCorners[a];
                    Vec3 pb = cameraCorners[b];
                    Vec3 pc = cameraCorners[c];
                    double ux = pb.X - pa.X;
                    double uy = pb.Y - pa.Y;
                    double uz = pb.Z - pa.Z;
                    double vx = pc.X - pa.X;
                    double vy = pc.Y - pa.Y;
                    double vz = pc.Z - pa.Z;
                    double nx = (uy * vz) - (uz * vy);
                    double ny = (uz * vx) - (ux * vz);
                    double nz = (ux * vy) - (uy * vx);
                    double normalLength = Math.Sqrt((nx * nx) + (ny * ny) + (nz * nz));
                    if (normalLength < 1e-9)
                    {
                        continue;
                    }

                    nx /= normalLength;
                    ny /= normalLength;
                    nz /= normalLength;

                    // Camera olha para +Z em coordenadas de camera; face frontal aponta para -Z.
                    if (Math.Abs(nz) <= 1e-6)
                    {
                        continue;
                    }

                    PointF[] polygon =
                    {
                        projected[a],
                        projected[b],
                        projected[c],
                        projected[d]
                    };
                    double depth = (projectedDepths[a] + projectedDepths[b] + projectedDepths[c] + projectedDepths[d]) * 0.25;

                    // Luz principal vinda da camera para manter leitura de frente/trás.
                    double facingToCamera = Math.Abs(nz);
                    double lambert = 0.45 + (0.55 * facingToCamera);
                    double depthDarkening = Math.Max(0.82, 1.0 - (Math.Max(0.0, depth) / Math.Max(1.0, sceneExtent * 6.0)));
                    double shade = lambert * depthDarkening;
                    Color shadedFillColor = ScaleColor(fillColor, shade);

                    double[] vertexDepths =
                    {
                        projectedDepths[a],
                        projectedDepths[b],
                        projectedDepths[c],
                        projectedDepths[d]
                    };

                    targetFaces.Add(
                        new BoxFaceRenderInfo(
                            polygon,
                            depth,
                            shadedFillColor,
                            edgeColor,
                            edgeThickness,
                            vertexDepths));
                }
            }

            private void CollectPartTriangleFaces(PreviewPart part, List<BoxFaceRenderInfo> targetFaces)
            {
                if (part == null || targetFaces == null || part.Triangles == null || part.Triangles.Count == 0)
                {
                    return;
                }

                Color edgeColor = part.IsMainPart
                    ? Color.FromArgb(92, 64, 8)
                    : Color.FromArgb(18, 28, 46);
                float edgeThickness = part.IsMainPart ? 2.2f : 1.2f;
                Color fillColor = part.IsMainPart
                    ? Color.FromArgb(210, 164, 32)
                    : Color.FromArgb(55, 79, 122);

                for (int i = 0; i < part.Triangles.Count; i++)
                {
                    PreviewTriangle tri = part.Triangles[i];
                    if (tri == null)
                    {
                        continue;
                    }

                    Vec3 w0 = new Vec3(tri.X1, tri.Y1, tri.Z1);
                    Vec3 w1 = new Vec3(tri.X2, tri.Y2, tri.Z2);
                    Vec3 w2 = new Vec3(tri.X3, tri.Y3, tri.Z3);

                    Vec3 c0;
                    Vec3 c1;
                    Vec3 c2;
                    if (!TryTransformToCamera(w0, out c0)
                        || !TryTransformToCamera(w1, out c1)
                        || !TryTransformToCamera(w2, out c2))
                    {
                        continue;
                    }

                    PointF p0;
                    PointF p1;
                    PointF p2;
                    double z0;
                    double z1;
                    double z2;
                    if (!TryProjectCamera(c0, out p0, out z0)
                        || !TryProjectCamera(c1, out p1, out z1)
                        || !TryProjectCamera(c2, out p2, out z2))
                    {
                        continue;
                    }

                    double ux = c1.X - c0.X;
                    double uy = c1.Y - c0.Y;
                    double uz = c1.Z - c0.Z;
                    double vx = c2.X - c0.X;
                    double vy = c2.Y - c0.Y;
                    double vz = c2.Z - c0.Z;
                    double nx = (uy * vz) - (uz * vy);
                    double ny = (uz * vx) - (ux * vz);
                    double nz = (ux * vy) - (uy * vx);
                    double normalLength = Math.Sqrt((nx * nx) + (ny * ny) + (nz * nz));
                    if (normalLength < 1e-9)
                    {
                        continue;
                    }

                    nx /= normalLength;
                    ny /= normalLength;
                    nz /= normalLength;

                    double facingToCamera = Math.Abs(nz);
                    double lambert = 0.45 + (0.55 * facingToCamera);
                    double depth = (z0 + z1 + z2) / 3.0;
                    double depthDarkening = Math.Max(0.82, 1.0 - (Math.Max(0.0, depth) / Math.Max(1.0, sceneExtent * 6.0)));
                    double shade = lambert * depthDarkening;
                    Color shadedFillColor = ScaleColor(fillColor, shade);

                    PointF[] polygon = { p0, p1, p2 };
                    double[] vertexDepths = { z0, z1, z2 };
                    targetFaces.Add(
                        new BoxFaceRenderInfo(
                            polygon,
                            depth,
                            shadedFillColor,
                            edgeColor,
                            edgeThickness,
                            vertexDepths));
                }
            }

            private void RenderDepthBufferedFaces(Graphics g, List<BoxFaceRenderInfo> faces)
            {
                if (g == null || faces == null || faces.Count == 0 || ClientSize.Width <= 0 || ClientSize.Height <= 0)
                {
                    return;
                }

                int width = ClientSize.Width;
                int height = ClientSize.Height;
                int pixelCount = width * height;
                float[] zBuffer = new float[pixelCount];
                int[] colorBuffer = new int[pixelCount];

                for (int i = 0; i < pixelCount; i++)
                {
                    zBuffer[i] = float.PositiveInfinity;
                    colorBuffer[i] = 0;
                }

                for (int i = 0; i < faces.Count; i++)
                {
                    BoxFaceRenderInfo face = faces[i];
                    if (face.Points == null
                        || face.Points.Length < 3
                        || face.VertexDepths == null
                        || face.VertexDepths.Length != face.Points.Length)
                    {
                        continue;
                    }

                    int fillArgb = face.FillColor.ToArgb();
                    for (int tri = 1; tri < face.Points.Length - 1; tri++)
                    {
                        RasterizeTriangle(
                            face.Points[0], face.VertexDepths[0],
                            face.Points[tri], face.VertexDepths[tri],
                            face.Points[tri + 1], face.VertexDepths[tri + 1],
                            fillArgb,
                            width,
                            height,
                            zBuffer,
                            colorBuffer);
                    }
                }

                for (int i = 0; i < faces.Count; i++)
                {
                    BoxFaceRenderInfo face = faces[i];
                    if (face.Points == null
                        || face.Points.Length < 2
                        || face.VertexDepths == null
                        || face.VertexDepths.Length != face.Points.Length)
                    {
                        continue;
                    }

                    int edgeArgb = face.EdgeColor.ToArgb();
                    for (int edge = 0; edge < face.Points.Length; edge++)
                    {
                        int next = (edge + 1) % face.Points.Length;
                        RasterizeDepthLine(
                            face.Points[edge],
                            face.VertexDepths[edge],
                            face.Points[next],
                            face.VertexDepths[next],
                            edgeArgb,
                            width,
                            height,
                            zBuffer,
                            colorBuffer);
                    }
                }

                using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    Rectangle rect = new Rectangle(0, 0, width, height);
                    BitmapData data = bitmap.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    try
                    {
                        System.Runtime.InteropServices.Marshal.Copy(colorBuffer, 0, data.Scan0, colorBuffer.Length);
                    }
                    finally
                    {
                        bitmap.UnlockBits(data);
                    }

                    g.DrawImageUnscaled(bitmap, 0, 0);
                }
            }

            private static void RasterizeTriangle(
                PointF p0,
                double z0,
                PointF p1,
                double z1,
                PointF p2,
                double z2,
                int colorArgb,
                int width,
                int height,
                float[] zBuffer,
                int[] colorBuffer)
            {
                float minXf = Math.Min(p0.X, Math.Min(p1.X, p2.X));
                float maxXf = Math.Max(p0.X, Math.Max(p1.X, p2.X));
                float minYf = Math.Min(p0.Y, Math.Min(p1.Y, p2.Y));
                float maxYf = Math.Max(p0.Y, Math.Max(p1.Y, p2.Y));

                int minX = Math.Max(0, (int)Math.Floor(minXf));
                int maxX = Math.Min(width - 1, (int)Math.Ceiling(maxXf));
                int minY = Math.Max(0, (int)Math.Floor(minYf));
                int maxY = Math.Min(height - 1, (int)Math.Ceiling(maxYf));
                if (minX > maxX || minY > maxY)
                {
                    return;
                }

                double area = EdgeFunction(p0, p1, p2);
                if (Math.Abs(area) < 1e-8)
                {
                    return;
                }

                for (int y = minY; y <= maxY; y++)
                {
                    for (int x = minX; x <= maxX; x++)
                    {
                        PointF p = new PointF(x + 0.5f, y + 0.5f);
                        double w0 = EdgeFunction(p1, p2, p);
                        double w1 = EdgeFunction(p2, p0, p);
                        double w2 = EdgeFunction(p0, p1, p);

                        bool inside = area > 0.0
                            ? (w0 >= 0.0 && w1 >= 0.0 && w2 >= 0.0)
                            : (w0 <= 0.0 && w1 <= 0.0 && w2 <= 0.0);
                        if (!inside)
                        {
                            continue;
                        }

                        w0 /= area;
                        w1 /= area;
                        w2 /= area;
                        float depth = (float)((w0 * z0) + (w1 * z1) + (w2 * z2));

                        int index = (y * width) + x;
                        if (depth < zBuffer[index])
                        {
                            zBuffer[index] = depth;
                            colorBuffer[index] = colorArgb;
                        }
                    }
                }
            }

            private static void RasterizeDepthLine(
                PointF p0,
                double z0,
                PointF p1,
                double z1,
                int colorArgb,
                int width,
                int height,
                float[] zBuffer,
                int[] colorBuffer)
            {
                int steps = (int)Math.Max(Math.Abs(p1.X - p0.X), Math.Abs(p1.Y - p0.Y));
                if (steps <= 0)
                {
                    return;
                }

                for (int i = 0; i <= steps; i++)
                {
                    float t = (float)i / steps;
                    int x = (int)Math.Round(p0.X + ((p1.X - p0.X) * t));
                    int y = (int)Math.Round(p0.Y + ((p1.Y - p0.Y) * t));
                    if (x < 0 || x >= width || y < 0 || y >= height)
                    {
                        continue;
                    }

                    float depth = (float)(z0 + ((z1 - z0) * t));
                    int index = (y * width) + x;
                    if (depth <= zBuffer[index] + 0.02f)
                    {
                        colorBuffer[index] = colorArgb;
                    }
                }
            }

            private static double EdgeFunction(PointF a, PointF b, PointF c)
            {
                return ((c.X - a.X) * (b.Y - a.Y)) - ((c.Y - a.Y) * (b.X - a.X));
            }

            private void DrawProjectedSegment(Graphics g, Vec3 from, Vec3 to, Pen pen)
            {
                PointF fromScreen;
                PointF toScreen;
                double fromDepth;
                double toDepth;
                if (!TryProject(from, out fromScreen, out fromDepth)
                    || !TryProject(to, out toScreen, out toDepth))
                {
                    return;
                }

                g.DrawLine(pen, fromScreen, toScreen);
            }

            private bool TryProject(Vec3 world, out PointF screen, out double depth)
            {
                Vec3 cameraPoint;
                if (!TryTransformToCamera(world, out cameraPoint))
                {
                    screen = PointF.Empty;
                    depth = 0.0;
                    return false;
                }

                return TryProjectCamera(cameraPoint, out screen, out depth);
            }

            private bool TryTransformToCamera(Vec3 world, out Vec3 cameraPoint)
            {
                double x = world.X - sceneCenterX;
                double y = world.Y - sceneCenterY;
                double z = world.Z - sceneCenterZ;

                double cosYaw = Math.Cos(yawRadians);
                double sinYaw = Math.Sin(yawRadians);
                double xRot = (cosYaw * x) + (sinYaw * z);
                double zYaw = (-sinYaw * x) + (cosYaw * z);

                double cosPitch = Math.Cos(pitchRadians);
                double sinPitch = Math.Sin(pitchRadians);
                double yRot = (cosPitch * y) - (sinPitch * zYaw);
                double zRot = (sinPitch * y) + (cosPitch * zYaw);

                cameraPoint = new Vec3(xRot, yRot, zRot);
                return true;
            }

            private bool TryProjectCamera(Vec3 cameraPoint, out PointF screen, out double depth)
            {
                double xRot = cameraPoint.X;
                double yRot = cameraPoint.Y;
                double zRot = cameraPoint.Z;

                double cameraRelativeDepth = zRot + cameraDistance;
                depth = cameraRelativeDepth;
                if (!useOrthographicProjection)
                {
                    double nearPlane = Math.Max(0.1, minCameraDistance * 0.05);
                    if (cameraRelativeDepth <= nearPlane)
                    {
                        screen = PointF.Empty;
                        return false;
                    }
                }
                else
                {
                    // Em ortografica mantemos a profundidade relativa da camera para z-buffer estavel.
                    // Apenas descartamos geometrias atras da camera para evitar inversoes frente/tras.
                    if (cameraRelativeDepth <= 1e-4)
                    {
                        screen = PointF.Empty;
                        return false;
                    }
                }

                float screenX;
                float screenY;
                if (useOrthographicProjection)
                {
                    double pixelsPerUnitBase = Math.Max(
                        0.001,
                        (Math.Min(ClientSize.Width, ClientSize.Height) * 0.42) / Math.Max(1.0, sceneExtent));
                    double zoomFactor = (sceneExtent * 2.8) / Math.Max(1.0, cameraDistance);
                    if (zoomFactor < 0.08)
                    {
                        zoomFactor = 0.08;
                    }
                    else if (zoomFactor > 8.0)
                    {
                        zoomFactor = 8.0;
                    }

                    double pixelsPerUnit = pixelsPerUnitBase * zoomFactor;
                    screenX = (float)((ClientSize.Width * 0.5) + panPixelsX + (xRot * pixelsPerUnit));
                    screenY = (float)((ClientSize.Height * 0.5) + panPixelsY - (yRot * pixelsPerUnit));
                }
                else
                {
                    double focal = Math.Max(120.0, Math.Min(ClientSize.Width, ClientSize.Height) * 0.95);
                    screenX = (float)((ClientSize.Width * 0.5) + panPixelsX + ((xRot * focal) / depth));
                    screenY = (float)((ClientSize.Height * 0.5) + panPixelsY - ((yRot * focal) / depth));
                }

                screen = new PointF(screenX, screenY);
                return true;
            }

            private static Color ScaleColor(Color color, double factor)
            {
                if (factor < 0.0)
                {
                    factor = 0.0;
                }
                else if (factor > 1.6)
                {
                    factor = 1.6;
                }

                int r = (int)Math.Round(color.R * factor);
                int g = (int)Math.Round(color.G * factor);
                int b = (int)Math.Round(color.B * factor);
                if (r < 0) r = 0; else if (r > 255) r = 255;
                if (g < 0) g = 0; else if (g > 255) g = 255;
                if (b < 0) b = 0; else if (b > 255) b = 255;
                return Color.FromArgb(color.A, r, g, b);
            }

            private double GetCameraDepth(Vec3 world)
            {
                double x = world.X - sceneCenterX;
                double y = world.Y - sceneCenterY;
                double z = world.Z - sceneCenterZ;

                double cosYaw = Math.Cos(yawRadians);
                double sinYaw = Math.Sin(yawRadians);
                double zYaw = (-sinYaw * x) + (cosYaw * z);

                double cosPitch = Math.Cos(pitchRadians);
                double sinPitch = Math.Sin(pitchRadians);
                double zRot = (sinPitch * y) + (cosPitch * zYaw);
                return zRot + cameraDistance;
            }

            private struct Vec3
            {
                public Vec3(double x, double y, double z)
                {
                    X = x;
                    Y = y;
                    Z = z;
                }

                public double X;
                public double Y;
                public double Z;
            }

            private struct BoxFaceRenderInfo
            {
                public BoxFaceRenderInfo(
                    PointF[] points,
                    double depth,
                    Color fillColor,
                    Color edgeColor,
                    float edgeThickness,
                    double[] vertexDepths)
                {
                    Points = points;
                    Depth = depth;
                    FillColor = fillColor;
                    EdgeColor = edgeColor;
                    EdgeThickness = edgeThickness;
                    VertexDepths = vertexDepths;
                }

                public PointF[] Points;
                public double Depth;
                public Color FillColor;
                public Color EdgeColor;
                public float EdgeThickness;
                public double[] VertexDepths;
            }
        }

        private sealed class OrbitGizmoControl : Control
        {
            public OrbitGizmoControl()
            {
                SetStyle(
                    ControlStyles.UserPaint
                    | ControlStyles.AllPaintingInWmPaint
                    | ControlStyles.OptimizedDoubleBuffer
                    | ControlStyles.ResizeRedraw,
                    true);
                BackColor = Color.FromArgb(245, 246, 248);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                float centerX = ClientSize.Width * 0.5f;
                float centerY = ClientSize.Height * 0.5f;
                float radius = Math.Min(ClientSize.Width, ClientSize.Height) * 0.35f;

                DrawGizmoRing(g, centerX, centerY, radius * 2.1f, radius * 0.95f, 0f, Color.FromArgb(42, 170, 72));
                DrawGizmoRing(g, centerX, centerY, radius * 2.1f, radius * 0.95f, 90f, Color.FromArgb(52, 88, 214));
                DrawGizmoRing(g, centerX, centerY, radius * 2.1f, radius * 0.95f, -60f, Color.FromArgb(196, 43, 45));
            }

            private static void DrawGizmoRing(
                Graphics g,
                float centerX,
                float centerY,
                float width,
                float height,
                float rotation,
                Color color)
            {
                GraphicsState state = g.Save();
                g.TranslateTransform(centerX, centerY);
                g.RotateTransform(rotation);

                using (Pen pen = new Pen(color, 2f))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    g.DrawArc(pen, -width / 2f, -height / 2f, width, height, 18f, 324f);
                }

                g.Restore(state);
            }
        }
    }
}
