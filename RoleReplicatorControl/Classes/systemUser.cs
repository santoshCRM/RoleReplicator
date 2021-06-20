using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
namespace RoleReplicatorControl
{
    internal class systemUser
    {
        public bool Select { get; set; }

        public Guid SystemUserID { get; set; }

        public string FullName { get; set; }

        public string Domainname { get; set; }

        public string BusinessUnit { get; set; }



        public List<Team> Teams = new List<Team>();

        public List<SecurityRole> Roles = new List<SecurityRole>();

        public List<Queue> Queues = new List<Queue>();

        internal void AddAssoc(Entity entity)
        {

            if (entity.HasValue("RoleID"))// && !string.IsNullOrEmpty(entity.GetAttributeValue<string>("RoleID"))
            //    !string.IsNullOrEmpty(entity.GetAttributeValue<string>("RoleID")))
            {// entity.Attributes["RoleID"].ToString()))

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
            // throw new NotImplementedException();
        }


        //private EntityReference _businessUnit;

        //public EntityReference BusinessUnit
        //{
        //    get { return _businessUnit; }
        //    set { _businessUnit = value; }
        //}




    }
}
