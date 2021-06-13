using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
namespace RoleReplicatorControl
{
    public partial class RoleReplicator : PluginControlBase
    {
        public RoleReplicator()
        {
            InitializeComponent();


        }


        private void btn_retUsers_Click(object sender, EventArgs e)
        {
            ExecuteMethod(GetUsers);
        }


        private void drp_users_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (drp_users.SelectedIndex == 0)
            {
                return;
            }

            ExecuteMethod(GetUserDetail);
            //bool flag = this.drp_users.SelectedIndex > 0;
            //if (flag)
            //{
            //    systemUser user = (systemUser)drp_users.SelectedItem;
            //    List<SecurityRole> userRoles = Helper.getSecurityRole(user.SystemUserID);
            //    Helper.SecurityRoles = userRoles;
            //    lstview_sRole.DataSource = Helper.SecurityRoles;
            //    lstview_sRole.DisplayMember = "Name";
            //    lstview_sRole.ValueMember = "RoldId";

            //    List<Team> userTeams = Helper.getTeambyUser(user.SystemUserID, user.BusinessUnit);
            //    Helper.Teams = userTeams;
            //    lstbox_Teams.DataSource = Helper.Teams;
            //    lstbox_Teams.DisplayMember = "Name";
            //    lstbox_Teams.ValueMember = "TeamId";

            //    List<Queue> userQueues = Helper.getQueuebyUser(user.SystemUserID, user.BusinessUnit);
            //    Helper.Queues = userQueues;
            //    lstbox_queues.DataSource = Helper.Queues;
            //    lstbox_queues.DisplayMember = "Name";
            //    lstbox_queues.ValueMember = "QueueId";

            //    List<systemUser> _user = DestinationUsers.Where(u => u.Select == true).ToList();
            //    foreach (systemUser row in _user)
            //    {
            //        row.Select = false;
            //    }
            //    lbl_bu.Text = user.BusinessUnit;
            //    grdview_user.DataSource = DestinationUsers.Where(x => x.BusinessUnit == user.BusinessUnit).ToList();

        }


        private void btn_copyRole_Click(object sender, EventArgs e)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Executing request...",
                Work = (bw, m) =>
                {
                    if (Service != null)
                    {
                        Guid[] selecteUsers = new Guid[DestinationUsers.Where(u => u.Select == true).Count()];
                        string[] selecteUserBU = new string[DestinationUsers.Where(u => u.Select == true).Count()];
                        int x = 0;
                        foreach (systemUser row in DestinationUsers.Where(u => u.Select == true))
                        {
                            selecteUsers[x] = row.SystemUserID;
                            selecteUserBU[x] = row.BusinessUnit;
                            x++;
                        }
                        if (selecteUsers.Count() > 0)
                        {
                            Helper.copyRole(selecteUsers, selecteUserBU, chk_role.Checked, chk_team.Checked, chk_queue.Checked);
                        }
                    }
                },
                PostWorkCallBack = m =>
                {
                    if (m.Error == null)
                    {

                    }
                    else
                    {
                        MessageBox.Show(this, "An error occured: " + m.Error.Message, "Error", MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                }
            });



        }

        private void txt_filter_TextChanged(object sender, EventArgs e)
        {
            if (Service != null)
            {
                if (!string.IsNullOrEmpty(txt_filter.Text))
                {
                    grdview_user.DataSource = DestinationUsers.Where(x => x.FullName.ToUpper().Contains(txt_filter.Text.ToUpper()) && x.BusinessUnit == ((systemUser)drp_users.SelectedItem).BusinessUnit).ToList();
                }
                else
                {
                    grdview_user.DataSource = DestinationUsers;
                }
            }
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
                        systemUser user = DestinationUsers.Where(x => x.SystemUserID.Equals(row.Cells[1].Value)).First();
                        user.Select = true;

                    }
                    else
                    {
                        row.Cells[0].Value = !(bool)row.Cells[0].Value;
                        systemUser user = DestinationUsers.Where(x => x.SystemUserID.Equals(row.Cells[1].Value)).First();
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
            ExecuteMethod(GetUsers);
        }
    }
}
