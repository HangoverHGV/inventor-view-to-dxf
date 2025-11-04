using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Inventor;
using ViewOptionsFormNamespace; // Replace with the actual namespace containing ViewOptionsForm


namespace InventorDXFExport
{
    // Replace these GUIDs with your own for production.
    [Guid("D5AAFB69-0EC4-44CE-89B2-C089F0492425")]
    [ProgId("DXFExport.AddInServer")]
    [ComVisible(true)]
    public class AddInServer : ApplicationAddInServer
    {
        private Inventor.Application _invApp;
        private ButtonDefinition _btnExport;
        private const string clientId = "{D5AAFB69-0EC4-44CE-89B2-C089F0492425}";

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _invApp = addInSiteObject.Application;

            var ctrlDefs = _invApp.CommandManager.ControlDefinitions;
            try { ctrlDefs["ViewToDXF_Button"]?.Delete(); } catch { }

            _btnExport = ctrlDefs.AddButtonDefinition(
                "Export View to DXF",
                "ViewToDXF_Button",
                CommandTypesEnum.kEditMaskCmdType,
                clientId,
                "Create a drawing from the active part and export a single view to DXF.",
                "Creates a drawing (ANSI mm), places the chosen view, removes all other views and saves as DXF.",
                null, null);

            _btnExport.OnExecute += BtnExport_OnExecute;

            try
            {
                var ribbons = _invApp.UserInterfaceManager.Ribbons;
                Ribbon ribbon;
                try { ribbon = ribbons["Part"]; } catch { ribbon = ribbons.Count > 0 ? ribbons[1] : null; }

                if (ribbon != null)
                {
                    RibbonTab tab = null;
                    try { tab = ribbon.RibbonTabs["ViewToDXF_Tab"]; } catch { }
                    if (tab == null)
                        // call the Add overload with minimum args to avoid overload ambiguity
                        tab = ribbon.RibbonTabs.Add("ViewToDXF", "ViewToDXF_Tab", clientId);

                    RibbonPanel panel = null;
                    try { panel = tab.RibbonPanels["Export_Panel"]; } catch { }
                    if (panel == null)
                        // use the simple overload to avoid mismatched parameter interpretation
                        panel = tab.RibbonPanels.Add("Export", "Export_Panel", clientId);

                    panel.CommandControls.AddButton(_btnExport, true);
                }
            }
            catch { }
        }

        public void Deactivate()
        {
            if (_btnExport != null)
            {
                try { _btnExport.OnExecute -= BtnExport_OnExecute; } catch { }
                try { _btnExport.Delete(); } catch { }
                _btnExport = null;
            }
            _invApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int CommandID) { }

        public object Automation => null;

        private void BtnExport_OnExecute(NameValueMap Context)
        {
            try
            {
                var invApp = _invApp;

                if (!(invApp.ActiveDocument is PartDocument partDoc))
                {
                    System.Windows.Forms.MessageBox.Show("Active document is not a part document. Open a part and try again.", "ViewToDXF", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                string templatePath = invApp.FileManager.GetTemplateFile(
                    DocumentTypeEnum.kDrawingDocumentObject,
                    SystemOfMeasureEnum.kMetricSystemOfMeasure,
                    DraftingStandardEnum.kANSI_DraftingStandard);

                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("The drawing template file was not found:\n" + templatePath,
                        "ViewToDXF", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                var drawingDoc = (DrawingDocument)invApp.Documents.Add(
                    DocumentTypeEnum.kDrawingDocumentObject,
                    templatePath,
                    true);

                var sheet = drawingDoc.Sheets[1];

                try { sheet.TitleBlock?.Delete(); } catch { }
                try { sheet.Border?.Delete(); } catch { }

                // Show the WinForms dialog to get orientation, style and scale from the user
                ViewOrientationTypeEnum orientation;
                DrawingViewStyleEnum viewStyle;
                double scale;

                using (var dlg = new ViewOptionsFormNamespace.ViewOptionsForm())
                {
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        // User cancelled — clean up by closing the drawing (without saving)
                        try { drawingDoc.Close(false); } catch { }
                        return;
                    }

                    orientation = dlg.SelectedOrientation;
                    viewStyle = dlg.SelectedViewStyle;
                    scale = dlg.Scale;
                }

                var tg = invApp.TransientGeometry;
                var insertPt = tg.CreatePoint2d(10, 10);

                var baseView = sheet.DrawingViews.AddBaseView(
                    partDoc as _Document,
                    insertPt,
                    scale,
                    orientation,
                    viewStyle
                );

                var toDelete = new List<DrawingView>();
                foreach (DrawingView v in sheet.DrawingViews)
                    if (!ReferenceEquals(v, baseView))
                        toDelete.Add(v);
                foreach (var v in toDelete) v.Delete();

                drawingDoc.Activate();

                string dxfPath;
                if (!string.IsNullOrEmpty(partDoc.FullFileName))
                    dxfPath = System.IO.Path.ChangeExtension(partDoc.FullFileName, ".dxf");
                else
                {
                    var baseName = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName ?? "Export");
                    dxfPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), baseName + ".dxf");
                }

                try
                {
                    drawingDoc.SaveAs(dxfPath, true);
                    System.Windows.Forms.MessageBox.Show("Saved drawing as DXF:\n" + dxfPath, "ViewToDXF", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
                catch (COMException saveEx)
                {
                    System.Windows.Forms.MessageBox.Show("SaveAs to DXF failed: " + saveEx.Message + "\nEnsure DWG/DXF translator is installed.",
                        "ViewToDXF", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            catch (COMException comEx)
            {
                System.Windows.Forms.MessageBox.Show("Unable to use Inventor: " + comEx.Message, "ViewToDXF", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "ViewToDXF", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}