using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace SBOAddonProject
{
    class Program
    {
        public static SAPbouiCOM.Application sboApp = null;
        public static SAPbobsCOM.Company oCompany = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                sboApp = (SAPbouiCOM.Application)Application.SBO_Application;
                oCompany = (SAPbobsCOM.Company)sboApp.Company.GetDICompany();
                //sboApp.MessageBox("Welcome to : " + oCompany.CompanyName.ToString());
                //sboApp.SetStatusBarMessage("Welcome to : " + oCompany.CompanyName.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
