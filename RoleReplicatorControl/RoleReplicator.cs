using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace RoleReplicatorControl
{
    public partial class RoleReplicator : PluginControlBase, IGitHubPlugin
    {
        public string RepositoryName => "RoleReplicator";
        private AppInsights ai;
        private const string aiEndpoint = "https://dc.services.visualstudio.com/v2/track";

        private const string aiKey = "f521c450-81df-45cb-aedc-700df1b55541";
        public string UserName => "santoshCRM";

        public RoleReplicator()
        {
            InitializeComponent();
            ai = new AppInsights(aiEndpoint, aiKey, Assembly.GetExecutingAssembly());
            ai.WriteEvent("Control Loaded");
        }


        private void btn_retUsers_Click(object sender, EventArgs e)
        {
            ExecuteMethod(GetUsers);
        }


        private void drp_users_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (drp_users.SelectedIndex == 0)
            {
                lstview_sRole.DataSource = new List<systemUser>();
                lstbox_queues.DataSource = new List<Queue>();
                lstbox_Teams.DataSource = new List<Team>();


                return;
            }

            ExecuteMethod(GetUserDetail);
            DisplayUsers();

        }

        private void btn_copyRole_Click(object sender, EventArgs e)
        {
            ExecuteMethod(CopyRole);
        }

        private void txt_filter_TextChanged(object sender, EventArgs e)
        {
            DisplayUsers();
        }

        private void grdview_user_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow row in grdview_user.SelectedRows)
            {
                if (row.Cells[0].ValueType == typeof(bool))
                {
                    if ((bool)row.Cells[0].Value == false)
                    {
                        row.Cells[0].Value = true;
                        systemUser user = DestinationUsers.Where(x => x.SystemUserID.Equals((row.DataBoundItem as systemUser).SystemUserID)).First();
                        user.Select = true;

                    }
                    else
                    {
                        row.Cells[0].Value = !(bool)row.Cells[0].Value;
                        systemUser user = DestinationUsers.Where(x => x.SystemUserID.Equals((row.DataBoundItem as systemUser).SystemUserID)).First();
                        user.Select = false;
                    }
                }
            }
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            base.CloseTool();
        }

        private void RoleReplicator_Load(object sender, EventArgs e)
        {
            ////ExecuteMethod(GetUsers);
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {

            base.UpdateConnection(newService, detail, actionName, parameter);

            LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);

            ExecuteMethod(GetUsers);
        }

        private void chkBU_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBU.Checked)
            {
                lblBU.Text = "Change BU";


                if (drp_users.SelectedIndex > 0)
                {
                    lblBU.Text += " to " + ((systemUser)drp_users.SelectedItem).BusinessUnit;
                }
            }
            else
            {
                lblBU.Text = "Keep BU";
            }
            DisplayUsers();
        }

    }
}
