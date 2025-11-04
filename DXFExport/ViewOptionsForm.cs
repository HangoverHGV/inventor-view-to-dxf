using Inventor;
using System;
using System.Drawing;
using System.IO;
using System.Reflection.Emit;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Environment = System.Environment;

namespace ViewOptionsFormNamespace
{
    public class ViewOptionsForm : Form
    {
        private ComboBox cmbOrientation;
        private ComboBox cmbViewStyle;
        private NumericUpDown nudScale;
        private Button btnOk;
        private Button btnCancel;

        // Stores the path chosen by the user when clicking OK
        public string SaveFilePath { get; private set; }

        // initial file name to show in SaveFileDialog (without extension)
        private readonly string _initialFileName;

        private readonly (string Label, ViewOrientationTypeEnum Value)[] _orientations =
        {
            ("Front", ViewOrientationTypeEnum.kFrontViewOrientation),
            ("Right", ViewOrientationTypeEnum.kRightViewOrientation),
            ("Top", ViewOrientationTypeEnum.kTopViewOrientation),
            ("Left", ViewOrientationTypeEnum.kLeftViewOrientation),
            ("Back", ViewOrientationTypeEnum.kBackViewOrientation),
            ("Bottom", ViewOrientationTypeEnum.kBottomViewOrientation),
            ("Iso Top Right", ViewOrientationTypeEnum.kIsoTopRightViewOrientation),
            ("Iso Bottom Right", ViewOrientationTypeEnum.kIsoBottomRightViewOrientation)
        };

        private readonly (string Label, DrawingViewStyleEnum Value)[] _viewStyles =
        {
            ("Hidden line", DrawingViewStyleEnum.kHiddenLineDrawingViewStyle),
            ("Shaded", DrawingViewStyleEnum.kShadedDrawingViewStyle)
        };

        public ViewOrientationTypeEnum SelectedOrientation
        {
            get
            {
                var idx = Math.Max(0, Math.Min(cmbOrientation.SelectedIndex, _orientations.Length - 1));
                return _orientations[idx].Value;
            }
        }

        public DrawingViewStyleEnum SelectedViewStyle
        {
            get
            {
                var idx = Math.Max(0, Math.Min(cmbViewStyle.SelectedIndex, _viewStyles.Length - 1));
                return _viewStyles[idx].Value;
            }
        }

        public double Scale => (double)nudScale.Value;

        // New constructor accepts an optional initial file name (without extension)
        public ViewOptionsForm(string initialFileName = "drawing")
        {
            _initialFileName = string.IsNullOrWhiteSpace(initialFileName)
                ? "drawing"
                : System.IO.Path.GetFileNameWithoutExtension(initialFileName);

            Text = "View Options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(800, 500); // window already increased
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            // Shared UI sizing
            var uiFont = new System.Drawing.Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            int leftLabel = 20;
            int leftControl = 230;
            int controlWidth = 500;
            int top = 20;
            int vSpacing = 72;
            int controlHeight = 40;

            var lblOrientation = new System.Windows.Forms.Label { Text = "Orientation:", Left = leftLabel, Top = top + 10, AutoSize = true, Font = uiFont };
            cmbOrientation = new ComboBox
            {
                Left = leftControl,
                Top = top,
                Width = controlWidth,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = uiFont,
                Height = controlHeight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            foreach (var o in _orientations) cmbOrientation.Items.Add(o.Label);

            top += vSpacing;

            var lblViewStyle = new System.Windows.Forms.Label { Text = "View style:", Left = leftLabel, Top = top + 10, AutoSize = true, Font = uiFont };
            cmbViewStyle = new ComboBox
            {
                Left = leftControl,
                Top = top,
                Width = controlWidth,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = uiFont,
                Height = controlHeight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            foreach (var s in _viewStyles) cmbViewStyle.Items.Add(s.Label);

            top += vSpacing;

            // Scale label and control — label sized to control height and vertically centered so it doesn't overlap nudScale
            var lblScale = new System.Windows.Forms.Label
            {
                Text = "Scale (1.0 = 1:1):",
                Left = leftLabel,
                Top = top,
                Width = 200,
                Height = controlHeight,
                AutoSize = false,
                Font = uiFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            nudScale = new NumericUpDown
            {
                Left = leftControl,
                Top = top,
                Width = 200,
                Height = controlHeight,
                DecimalPlaces = 2,
                Minimum = 0.01M,
                Maximum = 1000M,
                Increment = 0.1M,
                Value = 1.00M,
                Font = uiFont,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            // Safe default selections (ensure index exists)
            cmbOrientation.SelectedIndex = cmbOrientation.Items.Count > 1 ? 1 : 0;
            cmbViewStyle.SelectedIndex = cmbViewStyle.Items.Count > 0 ? 0 : -1;

            // Buttons at bottom right, larger
            int btnWidth = 110;
            int btnHeight = 44;
            int btnY = ClientSize.Height - 80;
            btnOk = new Button
            {
                Text = "OK",
                // Do not set DialogResult here; we'll set it after the user chooses a file.
                Left = leftControl + controlWidth - (btnWidth * 2) - 12,
                Top = btnY,
                Width = btnWidth,
                Height = btnHeight,
                Font = uiFont,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            // Wire up click to prompt for save location before closing
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = leftControl + controlWidth - btnWidth,
                Top = btnY,
                Width = btnWidth,
                Height = btnHeight,
                Font = uiFont,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[] {
                lblOrientation, cmbOrientation,
                lblViewStyle, cmbViewStyle,
                lblScale, nudScale,
                btnOk, btnCancel
            });
        }

        private void BtnOk_Click(object? sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title = "Save exported file",
                Filter = "DXF files (*.dxf)|*.dxf|All files (*.*)|*.*",
                DefaultExt = "dxf",
                FileName = _initialFileName + ".dxf",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            var result = dlg.ShowDialog(this);
            if (result != DialogResult.OK)
            {
                // User cancelled the save dialog — do not close the form.
                return;
            }

            // Store chosen path and close the form with OK result
            SaveFilePath = dlg.FileName;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}