using ERPBotHubDLL;
using SAPbouiCOM.Framework;
using System;

namespace EBHMaterialPlanning
{
    class Program
    {
        #region Declaretion
        public static SAPbouiCOM.Application SBO_App = null;
        public static SAPbobsCOM.Company oCompany;
        public static ERPBotHub_Commonfile EBH = new ERPBotHub_Commonfile();
        #endregion

        #region MainEvent
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
                EBH.SAPConnection(System.Reflection.Assembly.GetEntryAssembly().GetName().Name); oCompany = EBH.oCompany; SBO_App = EBH.SBO_App;
                PlanService.Store = new PlanStore();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                PlanForm.Start();
                ProductionHook.Start();
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
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                default:
                    break;
            }
        }
        #endregion
    }
}
