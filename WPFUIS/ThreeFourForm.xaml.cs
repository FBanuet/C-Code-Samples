using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;
using DATools.Util; 

namespace DATools.WPFUIS
{
    /// <summary>
    /// Lógica de interacción para ThreeFourForm.xaml
    /// </summary>
    public partial class ThreeFourForm : Window
    {
        Autodesk.Revit.UI.UIApplication _UIApp;
        Autodesk.Revit.ApplicationServices.Application _revitApp;
        Guid BLD34 = new Guid("bc9e34ab-fbfc-4948-8264-2f25a2058be8");


        public ThreeFourForm(Autodesk.Revit.ApplicationServices.Application revitApp)
        {
            InitializeComponent();
            _revitApp = revitApp;
        }

        private void Btn_FolderClick(object sender, RoutedEventArgs e)
        {
            var dialog = new WinForms.FolderBrowserDialog();
            if (WinForms.DialogResult.OK == dialog.ShowDialog())
            {
                SourceFolder.Text = dialog.SelectedPath;
            }
        }


        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if(string.IsNullOrEmpty(SourceFolder.Text))
            {
                MessageBox.Show("SELECCIONE DE FAVOR UNA RUTA ESPECIFICA DENTRO DE BIM360");
                SourceFolder.Focus();
                return;
            }
            if(Directory.Exists(SourceFolder.Text) == false)
            {
                MessageBox.Show("Sourcefolder does not exists!");
                SourceFolder.Focus();
                return;
            }

            string targetFolder = SourceFolder.Text;
            if(Directory.Exists(targetFolder) == false)
            {
                Directory.CreateDirectory(targetFolder);
            }

            
            ExportOnFile(BLD34, targetFolder);
            RAPWPF.WpfApplication.DoEvents();
            Autodesk.Revit.UI.TaskDialog.Show("INFIO", "EXPORTACIÓN EXITOSA!");
            Close();
            

        }

        private void ExportOnFile(Guid ModelGuid, string targetFolder)
        {
            Guid GlobalProjectGuid = new Guid("9352eecd-7567-4ebb-87a6-aacf7a964199");
            OpenOptions opt = new OpenOptions();
            NavisworksExportOptions nweOptions = new NavisworksExportOptions();
            Guid BLD34 = ModelGuid;
            DefaultOpenFromCloudCallback defaultOpen = new DefaultOpenFromCloudCallback();
            string viewNames = "";


            try
            {
                ModelPath mpath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(GlobalProjectGuid, BLD34);
                Document doc = _revitApp.OpenDocumentFile(mpath, opt, defaultOpen);
                Nwc3DViewExporter Exporter = new Nwc3DViewExporter();
                List<View3D> vistasExport = Exporter.GetBIMviews(doc);

                foreach (View3D vistaMNSR in vistasExport)
                {
                    string clave = vistaMNSR.Name;
                    viewNames += clave + Environment.NewLine;
                    nweOptions.ExportScope = NavisworksExportScope.View;
                    nweOptions.ViewId = vistaMNSR.Id;
                    nweOptions.ExportRoomGeometry = false;
                    nweOptions.ExportLinks = true;
                    doc.Export(targetFolder,clave + ".nwc", nweOptions);
                    doc.Close(false);

                }
                
            }
            catch(Exception e)
            {
                MessageBox.Show("INFO ERROR REVIT AXM MANSUR", e.Message);
                
            }
        }

        
       
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
