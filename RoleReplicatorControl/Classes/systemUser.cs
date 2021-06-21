using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
namespace RoleReplicatorControl
{
    internal class systemUser
    {
        public bool Select { get; set; }


        [DisplayName("User Name")]
        public string FullName { get; set; }

        [DisplayName("Email")]

        public string Domainname { get; set; }
        [DisplayName("Business Unit")]
        public string BusinessUnit { get; set; }
        [DisplayName("ID")]
        public Guid SystemUserID { get; set; }



        public List<Team> Teams = new List<Team>();

        public List<SecurityRole> Roles = new List<SecurityRole>();

        public List<Queue> Queues = new List<Queue>();

        internal void AddAssoc(Entity entity)
        {

            if (entity.HasValue("RoleID"))
            { 

                if (!Roles.Any(rl => rl.RoleId == (Guid)((AliasedValue)entity.Attributes["RoleID"]).Value))
                {
                    Roles.Add(
                        new SecurityRole
                        {
                            RoleId = (Guid)((AliasedValue)entity.Attributes["RoleID"]).Value,
                            Name = ((AliasedValue)entity.Attributes["RoleName"]).Value.ToString()
                        });
                }
            }
            if (entity.HasValue("QueueID"))
            {
                if (!Queues.Any(q => q.QueueId == (Guid)((AliasedValue)entity.Attributes["QueueID"]).Value))
                {
                    Queues.Add(
                       new Queue
                       {
                           QueueId = (Guid)((AliasedValue)entity.Attributes["QueueID"]).Value,
                           Name = ((AliasedValue)entity.Attributes["QueueName"]).Value.ToString()
                       });
                }
            }

            if (entity.HasValue("TeamID"))
            {
                if (!Teams.Any(tm => tm.TeamId == (Guid)((AliasedValue)entity.Attributes["TeamID"]).Value))
                {
                    Teams.Add(
                       new Team
                       {
                           TeamId = (Guid)((AliasedValue)entity.Attributes["TeamID"]).Value,
                           Name = ((AliasedValue)entity.Attributes["TeamName"]).Value.ToString()
                       });
                }
            }
        }
    }
}
