using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ImageInserter
{
    public partial class ThisAddIn
    {
        //アドイン起動時にリボンの読み込みを行う処理
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
#if false
            // 設定前のカルチャを表示
            System.Diagnostics.Debug.WriteLine("CurrentCulture: {0}", System.Threading.Thread.CurrentThread.CurrentCulture.Name);
            System.Diagnostics.Debug.WriteLine("CurrentUICulture: {0}", System.Threading.Thread.CurrentThread.CurrentUICulture.Name);
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            // 設定後のカルチャを表示
            System.Diagnostics.Debug.WriteLine("CurrentCulture: {0}", System.Threading.Thread.CurrentThread.CurrentCulture.Name);
            System.Diagnostics.Debug.WriteLine("CurrentUICulture: {0}", System.Threading.Thread.CurrentThread.CurrentUICulture.Name);
#endif

            return base.CreateRibbonExtensibilityObject();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

#region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
#endregion
    }
}
