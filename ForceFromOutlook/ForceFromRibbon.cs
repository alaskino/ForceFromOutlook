using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ForceFromOutlook
{
    [ComVisible(true)]
    public class ForceFromRibbon :
        Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        private static Bitmap _iconOn;
        private static Bitmap _iconOff;


        public ForceFromRibbon()
        {
        }


        // =========================================================
        // CARICAMENTO XML
        // =========================================================

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Ribbon1.xml");
        }


        public void Ribbon_Load(
            Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }


        // =========================================================
        // GRUPPO
        // =========================================================

        public string GetGroupLabel(
            Office.IRibbonControl control)
        {
            return Language.IsItalian
                ? "Mittente"
                : "Sender";
        }


        // =========================================================
        // LABEL PULSANTE
        // =========================================================

        public string GetLabel(
            Office.IRibbonControl control)
        {
            if (Language.IsItalian)
            {
                return ThisAddIn.ForceFromEnabled
                    ? "Forza DA - ON"
                    : "Forza DA - OFF";
            }

            return ThisAddIn.ForceFromEnabled
                ? "Force From - ON"
                : "Force From - OFF";
        }


        // =========================================================
        // STATO PREMUTO
        // =========================================================

        public bool GetPressed(
            Office.IRibbonControl control)
        {
            return ThisAddIn.ForceFromEnabled;
        }


        // =========================================================
        // CLICK
        // =========================================================

        public void ToggleForceFrom(
            Office.IRibbonControl control,
            bool pressed)
        {
            ThisAddIn.ForceFromEnabled = pressed;

            // Aggiorna immediatamente icona, testo e stato
            if (_ribbon != null)
            {
                _ribbon.Invalidate();
            }
        }


        // =========================================================
        // TOOLTIP
        // =========================================================

        public string GetSuperTip(
            Office.IRibbonControl control)
        {
            string account = "";

            try
            {
                account =
                    Globals.ThisAddIn
                        .GetDefaultAccountDisplayName();
            }
            catch
            {
            }


            if (Language.IsItalian)
            {
                if (ThisAddIn.ForceFromEnabled)
                {
                    string text =
                        "ATTIVO\n\n" +
                        "Forza ogni messaggio a essere inviato " +
                        "dall'account predefinito del profilo Outlook.";

                    if (!string.IsNullOrWhiteSpace(account))
                    {
                        text += "\n\nAccount: " + account;
                    }

                    text +=
                        "\n\nSviluppato da Mathieu Licata (Alaska)";

                    return text;
                }

                return
                    "DISATTIVO\n\n" +
                    "Outlook può utilizzare liberamente " +
                    "l'account scelto nel campo Da.\n\n" +
                    "Sviluppato da Mathieu Licata (Alaska)";
            }


            if (ThisAddIn.ForceFromEnabled)
            {
                string text =
                    "ENABLED\n\n" +
                    "Forces every message to be sent from " +
                    "the default account of the Outlook profile.";

                if (!string.IsNullOrWhiteSpace(account))
                {
                    text += "\n\nAccount: " + account;
                }

                text +=
                    "\n\nDeveloped by Mathieu Licata (Alaska)";

                return text;
            }

            return
                "DISABLED\n\n" +
                "Outlook can freely use the account selected " +
                "in the From field.\n\n" +
                "Developed by Mathieu Licata (Alaska)";
        }


        // =========================================================
        // ICONA DINAMICA
        // =========================================================

        public object GetImage(
            Office.IRibbonControl control)
        {
            if (ThisAddIn.ForceFromEnabled)
            {
                if (_iconOn == null)
                    _iconOn = CreateIcon(true);

                return _iconOn;
            }

            if (_iconOff == null)
                _iconOff = CreateIcon(false);

            return _iconOff;
        }


        private static Bitmap CreateIcon(bool enabled)
        {
            const int size = 64;

            Bitmap bitmap =
                new Bitmap(
                    size,
                    size,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode =
                    SmoothingMode.AntiAlias;

                g.InterpolationMode =
                    InterpolationMode.HighQualityBicubic;

                g.PixelOffsetMode =
                    PixelOffsetMode.HighQuality;

                g.Clear(Color.Transparent);


                Color dark =
                    enabled
                        ? Color.FromArgb(25, 96, 145)
                        : Color.FromArgb(100, 100, 100);

                Color accent =
                    enabled
                        ? Color.FromArgb(20, 170, 85)
                        : Color.FromArgb(165, 165, 165);


                // -------------------------------------------------
                // BUSTA
                // -------------------------------------------------

                using (Pen envelopePen =
                    new Pen(dark, 5f))
                {
                    envelopePen.StartCap =
                        LineCap.Round;

                    envelopePen.EndCap =
                        LineCap.Round;

                    Rectangle envelope =
                        new Rectangle(
                            6,
                            13,
                            45,
                            34);

                    g.DrawRoundedRectangle(
                        envelopePen,
                        envelope,
                        6);


                    // diagonali della busta
                    g.DrawLine(
                        envelopePen,
                        8,
                        17,
                        28,
                        32);

                    g.DrawLine(
                        envelopePen,
                        28,
                        32,
                        49,
                        17);
                }


                // -------------------------------------------------
                // CERCHIO STATUS
                // -------------------------------------------------

                Rectangle circle =
                    new Rectangle(
                        34,
                        32,
                        28,
                        28);

                using (SolidBrush circleBrush =
                    new SolidBrush(accent))
                {
                    g.FillEllipse(
                        circleBrush,
                        circle);
                }


                // -------------------------------------------------
                // FRECCIA VERSO L'ALTO / DA
                // -------------------------------------------------

                using (Pen arrowPen =
                    new Pen(Color.White, 4f))
                {
                    arrowPen.StartCap =
                        LineCap.Round;

                    arrowPen.EndCap =
                        LineCap.Round;

                    g.DrawLine(
                        arrowPen,
                        48,
                        53,
                        48,
                        40);

                    g.DrawLine(
                        arrowPen,
                        48,
                        40,
                        42,
                        46);

                    g.DrawLine(
                        arrowPen,
                        48,
                        40,
                        54,
                        46);
                }
            }

            return bitmap;
        }


        // =========================================================
        // LEGGE L'XML EMBEDDED
        // =========================================================

        private static string GetResourceText(
            string resourceFileName)
        {
            Assembly assembly =
                Assembly.GetExecutingAssembly();

            string[] resourceNames =
                assembly.GetManifestResourceNames();

            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith(
                    resourceFileName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    using (Stream stream =
                        assembly.GetManifestResourceStream(
                            resourceName))
                    {
                        if (stream == null)
                            return null;

                        using (StreamReader reader =
                            new StreamReader(stream))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }

            return null;
        }
    }


    // =============================================================
    // ESTENSIONE GRAFICA PER RETTANGOLO ARROTONDATO
    // =============================================================

    internal static class GraphicsExtensions
    {
        internal static void DrawRoundedRectangle(
            this Graphics graphics,
            Pen pen,
            Rectangle bounds,
            int radius)
        {
            int diameter = radius * 2;

            using (GraphicsPath path =
                new GraphicsPath())
            {
                path.AddArc(
                    bounds.Left,
                    bounds.Top,
                    diameter,
                    diameter,
                    180,
                    90);

                path.AddArc(
                    bounds.Right - diameter,
                    bounds.Top,
                    diameter,
                    diameter,
                    270,
                    90);

                path.AddArc(
                    bounds.Right - diameter,
                    bounds.Bottom - diameter,
                    diameter,
                    diameter,
                    0,
                    90);

                path.AddArc(
                    bounds.Left,
                    bounds.Bottom - diameter,
                    diameter,
                    diameter,
                    90,
                    90);

                path.CloseFigure();

                graphics.DrawPath(
                    pen,
                    path);
            }
        }
    }
}