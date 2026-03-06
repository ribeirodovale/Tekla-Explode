using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public sealed class OrbitToolForm : Form
    {
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
                bool isMainPart)
            {
                Label = string.IsNullOrWhiteSpace(label) ? "(sem id)" : label;
                MinX = minX;
                MinY = minY;
                MinZ = minZ;
                MaxX = maxX;
                MaxY = maxY;
                MaxZ = maxZ;
                IsMainPart = isMainPart;
            }

            public string Label { get; private set; }

            public double MinX { get; private set; }

            public double MinY { get; private set; }

            public double MinZ { get; private set; }

            public double MaxX { get; private set; }

            public double MaxY { get; private set; }

            public double MaxZ { get; private set; }

            public bool IsMainPart { get; set; }

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

            OrbitPreviewViewportControl viewport = new OrbitPreviewViewportControl(parts);
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

        private sealed class OrbitPreviewViewportControl : Control
        {
            private static readonly int[,] BoxEdges =
            {
                { 0, 1 }, { 1, 2 }, { 2, 3 }, { 3, 0 },
                { 4, 5 }, { 5, 6 }, { 6, 7 }, { 7, 4 },
                { 0, 4 }, { 1, 5 }, { 2, 6 }, { 3, 7 }
            };

            private readonly List<PreviewPart> parts = new List<PreviewPart>();

            private MouseButtons dragButton = MouseButtons.None;
            private Point lastMouse;

            private double yawRadians = -0.55;
            private double pitchRadians = 0.40;
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

            private void ResetCamera()
            {
                yawRadians = -0.55;
                pitchRadians = 0.40;
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

                using (Pen xPen = new Pen(Color.FromArgb(196, 43, 45), 2f))
                using (Pen yPen = new Pen(Color.FromArgb(38, 168, 72), 2f))
                using (Pen zPen = new Pen(Color.FromArgb(52, 88, 214), 2f))
                {
                    DrawProjectedSegment(
                        g,
                        new Vec3(sceneCenterX - axisLength, sceneCenterY, sceneCenterZ),
                        new Vec3(sceneCenterX + axisLength, sceneCenterY, sceneCenterZ),
                        xPen);
                    DrawProjectedSegment(
                        g,
                        new Vec3(sceneCenterX, sceneCenterY - axisLength, sceneCenterZ),
                        new Vec3(sceneCenterX, sceneCenterY + axisLength, sceneCenterZ),
                        yPen);
                    DrawProjectedSegment(
                        g,
                        new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ - axisLength),
                        new Vec3(sceneCenterX, sceneCenterY, sceneCenterZ + axisLength),
                        zPen);
                }
            }

            private void DrawPartBoxes(Graphics g)
            {
                List<PreviewPart> ordered = new List<PreviewPart>(parts);
                ordered.Sort(
                    delegate (PreviewPart left, PreviewPart right)
                    {
                        double leftDepth = GetCameraDepth(new Vec3(left.CenterX, left.CenterY, left.CenterZ));
                        double rightDepth = GetCameraDepth(new Vec3(right.CenterX, right.CenterY, right.CenterZ));
                        return rightDepth.CompareTo(leftDepth);
                    });

                for (int i = 0; i < ordered.Count; i++)
                {
                    DrawPartBox(g, ordered[i]);
                }
            }

            private void DrawPartBox(Graphics g, PreviewPart part)
            {
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
                int visibleCount = 0;

                for (int i = 0; i < corners.Length; i++)
                {
                    double depth;
                    visible[i] = TryProject(corners[i], out projected[i], out depth);
                    if (visible[i])
                    {
                        visibleCount++;
                    }
                }

                if (visibleCount < 2)
                {
                    return;
                }

                Color lineColor = part.IsMainPart
                    ? Color.FromArgb(210, 164, 32)
                    : Color.FromArgb(55, 79, 122);
                float thickness = part.IsMainPart ? 2.3f : 1.2f;

                using (Pen boxPen = new Pen(lineColor, thickness))
                {
                    boxPen.StartCap = LineCap.Round;
                    boxPen.EndCap = LineCap.Round;

                    for (int edge = 0; edge < BoxEdges.GetLength(0); edge++)
                    {
                        int a = BoxEdges[edge, 0];
                        int b = BoxEdges[edge, 1];
                        if (!visible[a] || !visible[b])
                        {
                            continue;
                        }

                        g.DrawLine(boxPen, projected[a], projected[b]);
                    }
                }
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

                depth = zRot + cameraDistance;
                double nearPlane = Math.Max(0.1, minCameraDistance * 0.05);
                if (depth <= nearPlane)
                {
                    screen = PointF.Empty;
                    return false;
                }

                double focal = Math.Max(120.0, Math.Min(ClientSize.Width, ClientSize.Height) * 0.95);
                float screenX = (float)((ClientSize.Width * 0.5) + panPixelsX + ((xRot * focal) / depth));
                float screenY = (float)((ClientSize.Height * 0.5) + panPixelsY - ((yRot * focal) / depth));

                screen = new PointF(screenX, screenY);
                return true;
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
