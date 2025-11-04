using Inventor;
using System;
using System.Drawing;
using System.Reflection.Emit;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace ViewOptionsFormNamespace
{
    public class ViewOptionsForm : Form
    {
        private ComboBox cmbOrientation;
        private ComboBox cmbViewStyle;
        private NumericUpDown nudScale;
        private Button btnOk;
        private Button btnCancel;

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

        public ViewOptionsForm()
        {
            Text = "View Options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(360, 160);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            var lblOrientation = new System.Windows.Forms.Label { Text = "Orientation:", Left = 12, Top = 12, AutoSize = true };
            cmbOrientation = new ComboBox { Left = 110, Top = 8, Width = 230, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var o in _orientations) cmbOrientation.Items.Add(o.Label);

            var lblViewStyle = new System.Windows.Forms.Label { Text = "View style:", Left = 12, Top = 48, AutoSize = true };
            cmbViewStyle = new ComboBox { Left = 110, Top = 44, Width = 230, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var s in _viewStyles) cmbViewStyle.Items.Add(s.Label);

            var lblScale = new System.Windows.Forms.Label { Text = "Scale (1.0 = 1:1):", Left = 12, Top = 84, AutoSize = true };
            nudScale = new NumericUpDown
            {
                Left = 150,
                Top = 80,
                Width = 80,
                DecimalPlaces = 2,
                Minimum = 0.01M,
                Maximum = 1000M,
                Increment = 0.1M,
                Value = 1.00M
            };

            // Safe default selections (ensure index exists)
            cmbOrientation.SelectedIndex = cmbOrientation.Items.Count > 1 ? 1 : 0;
            cmbViewStyle.SelectedIndex = cmbViewStyle.Items.Count > 0 ? 0 : -1;

            btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 200, Width = 70, Top = 112 };
            btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 280, Width = 70, Top = 112 };

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[] { lblOrientation, cmbOrientation, lblViewStyle, cmbViewStyle, lblScale, nudScale, btnOk, btnCancel });
        }
    }
}