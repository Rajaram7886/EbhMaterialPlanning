using ERPBotHubDLL;
using SAPbouiCOM.Framework;
using System;

namespace EBHMaterialPlanning
{
    class Menu
    {
        ERPBotHub_Commonfile EBH = new ERPBotHub_Commonfile();
        public void AddMenuItems()
        {
            try
            {
                //Addon Menu Setup
                EBH.AddMenu("43520", SAPbouiCOM.BoMenuType.mt_POPUP, "Mrp_ERPBotHub", "EBH Material Planning", -1, "Menu.png");
                EBH.AddMenu("Mrp_ERPBotHub", SAPbouiCOM.BoMenuType.mt_STRING, "frmMrpPlan", "Material Request Planning", -1, "Menu.png");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("EBH Material Planning menu : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction) return;
                if (pVal.MenuUID == "ErpDB") // ErpDB
                {
                    DBSetup activeForm = new DBSetup();
                    activeForm.Show();
                }
                //Addon Menu
                if (pVal.MenuUID == "frmMrpPlan") MaterialPlanForm.Show();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("EBH Material Planning : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

    }
}
