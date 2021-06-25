using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using DATools.Util;
using DATools.UIForm;
using DATools.WPFUIS;

namespace DATools.Commands
{

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class _3_4_NWC : IExternalCommand, IOpenFromCloudCallback
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            ThreeFourForm dialog = new ThreeFourForm(commandData.Application.Application);
            dialog.ShowDialog();


            return Result.Succeeded;
        }

        public OpenConflictResult OnOpenConflict(OpenConflictScenario scenario)
        {
            throw new NotImplementedException();
        }
    }
}

