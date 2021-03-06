﻿using System;
using System.Collections.Generic;
using System.Linq;
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
            grdview_user.DataSource = chkBU.Checked ? DestinationUsers : DestinationUsers.Where(su => su.BusinessUnit == user.BusinessUnit).ToList();
            lblBU.Text = chkBU.Checked ? "Change BU to " + user.BusinessUnit : "Keep BU";

        }

        private void ConfigDetails()
        {
            lstview_sRole.DisplayMember = "Name";
            lstview_sRole.ValueMember = "RoleId";
            lstbox_Teams.DisplayMember = "Name";
            lstbox_Teams.ValueMember = "TeamId";
            lstbox_queues.DisplayMember = "Name";
            lstbox_queues.ValueMember = "QueueId";
        }

        void CopyRole()
        {
            if (drp_users.SelectedItem == null || drp_users.SelectedIndex == 0)
            {
                MessageBox.Show(
                    "Please select a source user before initiating the copy",
                    "Select Users to Copy From",
                    MessageBoxButtons.OK);
                return;
            }
            List<systemUser> destUsers = DestinationUsers.Where(usr => usr.Select && usr.SystemUserID != ((systemUser)drp_users.SelectedItem).SystemUserID).ToList();
            if (!destUsers.Any())
            {
                MessageBox.Show(
                    "Please select one or more destination users before initiating the copy",
                    "Select Users to Copy To",
                    MessageBoxButtons.OK);

                return;
            }

            if (!(chk_role.Checked || chk_queue.Checked || chk_team.Checked))
            {
                MessageBox.Show(
                    "Please select one or more of the Role, Team or Queue boxes to initiate the copy",
                    "Select Type of Copy",
                    MessageBoxButtons.OK);

                return;
            }

            List<SecurityRole> roles = (List<SecurityRole>)lstview_sRole.DataSource;
            List<Team> teams = (List<Team>)lstbox_Teams.DataSource;
            List<Queue> queues = (List<Queue>)lstbox_queues.DataSource;

            string caption = "Remove and copy " +
                (chk_role.Checked && roles.Any() ? "Roles, " : string.Empty) +
                (chk_team.Checked ? "Teams, " : string.Empty) +
                (chk_queue.Checked ? "Queues, " : string.Empty);
            caption = caption.Substring(0, caption.Length - 2) + " for " + Environment.NewLine;
            caption += String.Join(Environment.NewLine, destUsers.Select(usr => usr.FullName));


            if (MessageBox.Show(caption, "Confirm Copy", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            Guid buId = ((systemUser)drp_users.SelectedItem).BUId;


            WorkAsync(new WorkAsyncInfo
            {
                Message = "Executing request...",
                Work = (bw, m) =>
                {
                    string error = string.Empty;
                    bw.ReportProgress(-1, "Getting destination user details");
                    Helper.getUserDetail(destUsers);

                    if (chkBU.Checked)
                    {
                        bw.ReportProgress(-1, "Changing BUs");
                        string response = Helper.ChangeBUs(destUsers, buId);
                        if (!string.IsNullOrEmpty(response)) error += "Error in Changing BUs: " + response;
                    }

                    if (chk_role.Checked && roles.Any() && string.IsNullOrEmpty(error))
                    {
                        bw.ReportProgress(-1, "Removing Roles");
                        Helper.removeuserRole(destUsers);
                        bw.ReportProgress(-1, "Adding Roles");

                        Helper.adduserRole(destUsers, roles);
                        ai.WriteEvent("Roles Copied", roles.Count * destUsers.Count);
                    }

                    if (chk_team.Checked)
                    {
                        bw.ReportProgress(-1, "Removing Teams");
                        Helper.removeuserTeam(destUsers);
                        if (teams.Any())
                        {
                            bw.ReportProgress(-1, "Adding Teams");
                            Helper.adduserTeam(destUsers, teams);
                            ai.WriteEvent("Teams Copied", teams.Count * destUsers.Count);
                        }
                    }

                    if (chk_queue.Checked && queues.Any())
                    {
                        bw.ReportProgress(-1, "Removing Queues");
                        Helper.removeuserQueue(destUsers);

                        if (queues.Any())
                        {
                            bw.ReportProgress(-1, "Adding Queues");
                            Helper.adduserQueue(destUsers, queues);
                            ai.WriteEvent("Queues Copied", queues.Count * destUsers.Count);
                        }
                    }

                },
                PostWorkCallBack = m =>
                {
                    if (m.Error == null)
                    {
                        ai.WriteEvent("Users updated", destUsers.Count);
                        MessageBox.Show(
                            "All changes made successfully, please confirm",
                            "Success",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(this, "An error occured: " + m.Error.Message, "Error", MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                    GetUsers();
                },
                ProgressChanged = e => SetWorkingMessage(e.UserState.ToString())
            });


        }

        void DisplayUsers()
        {
            var userList = DestinationUsers;

            if (drp_users.SelectedIndex > 0 && !chkBU.Checked)
            {
                var selectedUser = ((systemUser)drp_users.SelectedItem);
                userList = userList.Where(usr => usr.BusinessUnit == selectedUser.BusinessUnit).ToList();

            }

            if (!string.IsNullOrEmpty(txt_filter.Text))
            {
                userList = userList.Where(usr => (usr.FullName.ToUpper().Contains(txt_filter.Text.ToUpper())
                           || usr.Domainname.ToUpper().Contains(txt_filter.Text.ToUpper()))).ToList();
            }

            grdview_user.DataSource = userList;

        }
    }
}
