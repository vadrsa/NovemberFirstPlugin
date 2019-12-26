using CefSharp;
using CefSharp.Wpf;
using DevExpress.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
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
using Uniconta.API.Service;
using Uniconta.API.System;
using Uniconta.ClientTools;
using Uniconta.ClientTools.Controls;
using Uniconta.ClientTools.DataModel;
using Uniconta.ClientTools.Page;
using Uniconta.Common;
using Uniconta.DataModel;
using UnicontaClient.Pages;
using UnicontaClient.Pages.Creditor.Payments;

namespace NovemberFirstPlugin
{
    public class DebtorGrid : UnicontaDataGrid
    {
        public override Type TableType
        {
            get
            {
                return typeof(DebtorClient);
            }
        }
    }



    public partial class NFPage : GridBasePage
    {
        BackgroundWorker backgroundWorker;
        bool availableTrans;
        Uniconta.Common.Language loc;
        string ActionName;


        public override string NameOfControl
        {
            get
            {
                return "NovemberFirstPage";/* should be a unique name*/
            }
        }

        public List<PluginControl> RegisterControls()
        {
            var ctrls = new List<PluginControl>();
            ctrls.Add(new PluginControl()
            {
                UniqueName = "NovemberFirstPage",
                PageType = typeof(DebtorClient),
                AllowMultipleOpen = false,
                PageHeader = "November First"
            });
            return ctrls;
        }

        List<TreeRibbon> CreateRibbonItems()
        {
            var ribbonItems = new List<TreeRibbon>();
            var editRowItem = new TreeRibbon();
            var localized = Uniconta.ClientTools.Localization.lookup("Refresh", loc);
            ActionName = localized;
            editRowItem.Name = localized;
            editRowItem.ActionName = localized;
            editRowItem.LargeGlyph = LargeIcon.Refresh.ToString();
            ribbonItems.Add(editRowItem);
            return ribbonItems;
        }


        public NFPage(CrudAPI api) : base(api, string.Empty)
        {
            try
            {
                InitializeComponent();
                loc = Uniconta.ClientTools.Localization.LanguageCode;
                localMenu.AddRibbonItems(CreateRibbonItems());
                localMenu.OnItemClicked += LocalMenu_OnItemClickedAsync;
                //var serviceProvider = (IServiceProvider)myWebBrowser.Document;
                //if (serviceProvider != null)
                //{
                //    Guid serviceGuid = new Guid("0002DF05-0000-0000-C000-000000000046");
                //    Guid iid = typeof(SHDocVw.WebBrowser).GUID;
                //    var webBrowserPtr = (SHDocVw.WebBrowser)serviceProvider
                //        .QueryService(ref serviceGuid, ref iid);
                //    if (webBrowserPtr != null)
                //    {
                //        webBrowserPtr.NewWindow2 += webBrowser1_NewWindow2;
                //    }
                //}
                this.Loaded += NFPage_Loaded;
                backgroundWorker = new BackgroundWorker();
                availableTrans = false;

                //this.myWebBrowser.Navigated += MyWebBrowser_Navigated;

            }
            catch (Exception excMsg)
            {
                NovemberFirstPlugin.Logger.logMessage(excMsg.Message);
            }
        }

        private void MyWebBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            //SetSilent(myWebBrowser, true);
        }

        public static void SetSilent(WebBrowser browser, bool silent)
        {
            if (browser == null)
                throw new ArgumentNullException("browser");
            IOleServiceProvider sp = browser.Document as IOleServiceProvider;
            if (sp != null)
            {
                Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
                Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");

                object webBrowser;
                sp.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out webBrowser);
                if (webBrowser != null)
                {
                    webBrowser.GetType().InvokeMember("Silent", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { silent });
                }
            }
        }


        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
        }
        //private void webBrowser1_NewWindow2(ref object ppDisp, ref bool Cancel)
        //{
        //}

        //[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        //[Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
        //internal interface IServiceProvider
        //{
        //    [return: MarshalAs(UnmanagedType.IUnknown)]
        //    object QueryService(ref Guid guidService, ref Guid riid);
        //}

       

        private void LocalMenu_OnItemClickedAsync(string ActionType)
        {
            NovemberFirstPlugin.Logger.logMessage(ActionType);
            if (ActionType == ActionName)
            {
                CheckAvailableTransaction();
                try
                {
                    ShowResults();
                }
                catch (Exception exc)
                {
                    NovemberFirstPlugin.Logger.logMessage(exc.Message);
                }
            }
        }


        private void NFPage_Loaded(object sender, RoutedEventArgs e)
        {
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.RunWorkerAsync();
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ShowResults();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            CheckAvailableTransaction();
        }

        public void ShowResults()
        {
            if (availableTrans)
            {
                OpenUrl();
                SendWebhook();
            }
            else
            {
                UnicontaMessageBox.Show("No Open Transaction Found To Process", "", MessageBoxButton.OK, MessageBoxImage.Information);
                OpenUrl(true);
            }
        }
        public void CheckAvailableTransaction()
        {
            try
            {
                IEnumerable<CreditorTransPayment> creditorsList = api.Query<CreditorTransPayment>().GetAwaiter().GetResult().ToList();
                if (creditorsList.Count() > 0)
                {
                    if (LoadPayments(creditorsList))
                        availableTrans = true;
                    else
                        availableTrans = false;
                }
                else
                    availableTrans = false;
            }
            catch (Exception excMsg)
            {
                NovemberFirstPlugin.Logger.logMessage(excMsg.Message);
            }
        }


        public void OpenUrl(bool noTrans = false)
        {
            if (noTrans)
            {
                var completeURL = "https://wedoio.dk/wp-content/uploads//2019/12/n1-logo.png";
                Uri uri = new Uri(completeURL, UriKind.RelativeOrAbsolute);
                if (!uri.IsAbsoluteUri)
                {
                    UnicontaMessageBox.Show("The Address URI must be absolute. For example, 'http://www.microsoft.com'", " ", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                this.myWebBrowser.Address = uri.ToString();

            }
            else
            {
                WaitDialogForm frm = new WaitDialogForm();
                frm.Show();
                frm.Caption = "Fetching Open Transactions";
                Thread.Sleep(5000);
                // https://app-api.novemberfirst.com/api/economic/login?id= 
                var novemberFirstUrl = api.CompanyEntity.GetUserField("NovemberFirstURL").ToString();
                // var customerNumber = api.CompanyEntity.GetUserField("CustomerNumber").ToString();
                var completeURL = novemberFirstUrl;// + customerNumber;
                Uri uri = new Uri(completeURL, UriKind.RelativeOrAbsolute);
                if (!uri.IsAbsoluteUri)
                {
                    UnicontaMessageBox.Show("The Address URI must be absolute. For example, 'http://www.microsoft.com'", " ", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                this.myWebBrowser.Address = uri.ToString();

                frm.Close();
            }
        }


        public void SendWebhook()
        {
            try
            {
                string autoCreationParam = "";
                string companyID = api.CompanyEntity.CompanyId.ToString();
                string URI = api.CompanyEntity.GetUserField("Wedoio").ToString();
                bool auto = (bool)api.CompanyEntity.GetUserField("Autobook");
                if (auto)
                    autoCreationParam = "True";
                else
                    autoCreationParam = "False";
                string GLSelected = "";
                var GL = (api.CompanyEntity.GetUserField("GL") as GLDailyJournalClient);
                if (GL != null)
                    GLSelected = GL.RowId.ToString();
                else
                    GLSelected = "0";
                string key = CreateMD5(URI + "CompanyID" + companyID);
                //string myParameters = "CompanyID=" + companyID + "&Key=" + key + "?GL=" + GLSelected + "&Auto=" + autoCreationParam;
                string myParametersURL = "/" + companyID + "/" + key + "?GL=" + GLSelected + "&Auto=" + autoCreationParam;
                URI = URI + myParametersURL;
                using (WebClient wc = new WebClient())
                {
                    string HtmlResult = wc.UploadString(URI, null);
                }
            }
            catch (Exception msg)
            {
                NovemberFirstPlugin.Logger.logMessage("An Error Occured While Sending The Webhook = " + msg.Message);
            }
        }

        public string CreateMD5(string input)
        {
            // Use input string to calculate MD5 hash
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);
                // Convert the byte array to hexadecimal string
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }
                return sb.ToString();
            }
        }




        bool LoadPayments(IEnumerable<CreditorTransPayment> lst)
        {
            var today = BasePage.GetSystemDefaultDate().Date;
            var company = api.CompanyEntity;
            bool openTransaction = false;
            foreach (var rec in lst)
            {
                if (rec.PaymentId == string.Empty || rec.PaymentAmount > 0 || !string.IsNullOrEmpty(rec.Payment))
                    openTransaction = true;
                else
                    openTransaction = false;
            }
            if (openTransaction)
                return true;
            else
                return false;
        }

    }
}
