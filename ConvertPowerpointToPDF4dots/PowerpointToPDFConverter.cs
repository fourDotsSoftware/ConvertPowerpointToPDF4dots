using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace ConvertPowerpointToPDF4dots
{
    public class PowerpointToPDFConverter
    {        
        public string err = "";

        public bool ConvertToPDF(string filepath,string outfilepath)
        {
            err = "";
            
            object oDocuments = null;
            object doc = null;

            try
            {
                OfficeHelper.CreatePowerPointApplication();

                oDocuments = OfficeHelper.PPApp.GetType().InvokeMember("Presentations", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.PPApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);

                OfficeHelper.PPApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.PPApp, null);
                */

                System.Threading.Thread.Sleep(200);

                /*
                string fp=System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(filepath),
                    System.IO.Path.GetFileNameWithoutExtension(filepath)+".pdf"
                    );                
                */

                doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, doc, new object[] { outfilepath, 32 });

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Powerpoint to PDF") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }                
    }
}