using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace RoleReplicatorControl
{
    public partial class RoleReplicator : PluginControlBase

    {
        private List<systemUser> DestinationUsers = new List<systemUser>();
        void GetUsers()
        {
            Helper.createConn(Service);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving Users",
                Work = (wrk, e) =>

                    e.Result = Helper.getUsers(),
                PostWorkCallBack = e =>
                {
                    DestinationUsers = (List<systemUser>)e.Result;
                    DestinationUsers.Sort((p, q) => p.FullName.CompareTo(q.FullName));

                    ConfigUserList(new List<systemUser>((List<systemUser>)e.Result));
                    ConfigUserGrid();
                },
                ProgressChanged = e => { }
            });
        }

        private void ConfigUserGrid()
        {
            grdview_user.DataSource = DestinationUsers;
            grdview_user.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            grdview_user.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            grdview_user.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            grdview_user.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            grdview_user.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }

        private void ConfigUserList(List<systemUser> users)
        {
            systemUser Select = new systemUser();
            Select.FullName = "--Select--";
            Select.SystemUserID = Guid.Empty;
            users.Insert(0, Select);
            drp_users.DataSource = users;
            drp_users.DisplayMember = "FullName";
            drp_users.ValueMember = "SystemUserID";
        }

        void GetUserDetail()
        {
            Helper.createConn(Service);
            var user = (systemUser)drp_users.SelectedItem;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving User Roles",
                Work = (wrk, e) =>
                {
                    var roles = Helper.getSecurityRole(user.SystemUserID);

                    wrk.ReportProgress(33, "Retrieving Teams");
                    var teams = Helper.getTeambyUser(user.SystemUserID, user.BusinessUnit);
                    wrk.ReportProgress(66, "Retrieving Queues");
                    var queues = Helper.getQueuebyUser(user.SystemUserID, user.BusinessUnit);

                    e.Result = new object[] { roles, teams, queues };
                },
                PostWorkCallBack = e =>
                {
                    var result = (object[])e.Result;
                    lstview_sRole.DataSource = result[0];
                    lstbox_Teams.DataSource = result[1];
                    lstbox_queues.DataSource = result[2];

                    ConfigDetails();
                },
                ProgressChanged = e => SetWorkingMessage(e.UserState.ToString())
            });

        }

        private void ConfigDetails()
        {
            lstview_sRole.DisplayMember = "Name";
            lstview_sRole.ValueMember = "RoldId";
            lstbox_Teams.DisplayMember = "Name";
            lstbox_Teams.ValueMember = "TeamId";
            lstbox_queues.DisplayMember = "Name";
            lstbox_queues.ValueMember = "QueueId";
        }
    }
}
