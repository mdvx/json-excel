using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace JsonExcel
{
    //[ComVisible(true)]
    //[ClassInterface(ClassInterfaceType.AutoDual)]
    //public class ComLibrary
    //{
    //    public string ComLibraryHello()
    //    {
    //        return "Hello from JsonExcel.ComLibrary";
    //    }
    //    public double Add(double x, double y)
    //    {
    //        return x + y;
    //    }
    //}
    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            //ComServer.DllUnregisterServer();
        }
    }

}
