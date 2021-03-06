﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace fncGTD4Outlook.Controles
{
    class TagControl : Panel
    {
        [Browsable(true)]
        public Color _BackColor { get; set; }

        public TagControl(string texto)
        {
            this.DoubleBuffered = true;

            Label miTexto = new Label();
            miTexto.Text = texto;
            miTexto.AutoSize = true;
            //miTexto.Location = new Point(5, 5);
            miTexto.Left = (this.ClientSize.Width - miTexto.Width) / 2;
            miTexto.Top = (this.ClientSize.Height - miTexto.Height) / 2;
            //Size = new System.Drawing.Size(43, 18),
            //miTexto.BorderStyle = BorderStyle.FixedSingle;
            this.Controls.Add(miTexto);

            this.Controls.Add(
                new Label
                {
                    Location = new Point(miTexto.Width + 10, 5),
                    AutoSize = true,
                    Font = new Font(this.Font.Name, 10, FontStyle.Bold),
                    Text = "x"
                });

            this.Height = 45;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (var graphicsPath = _getRoundRectangle(this.ClientRectangle))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(_BackColor))
                    e.Graphics.FillPath(brush, graphicsPath);
                using (var pen = new Pen(_BackColor, 1.0f))
                    e.Graphics.DrawPath(pen, graphicsPath);
                TextRenderer.DrawText(e.Graphics, Text, this.Font, this.ClientRectangle, this.ForeColor);
            }
        }

        private GraphicsPath _getRoundRectangle(Rectangle rectangle)
        {
            int cornerRadius = 20; // change this value according to your needs
            int diminisher = 1;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(rectangle.X, rectangle.Y, cornerRadius, cornerRadius, 180, 90);
            path.AddArc(rectangle.X + rectangle.Width - cornerRadius - diminisher, rectangle.Y, cornerRadius, cornerRadius, 270, 90);
            path.AddArc(rectangle.X + rectangle.Width - cornerRadius - diminisher, rectangle.Y + rectangle.Height - cornerRadius - diminisher, cornerRadius, cornerRadius, 0, 90);
            path.AddArc(rectangle.X, rectangle.Y + rectangle.Height - cornerRadius - diminisher, cornerRadius, cornerRadius, 90, 90);
            path.CloseAllFigures();
            return path;
        }

    }
}
