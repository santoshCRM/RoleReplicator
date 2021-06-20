using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Windows.Forms;
using System.Xml;
namespace RoleReplicatorControl
{
    public static class Helper
    {


        private static List<SecurityRole> _securityRoles;
        private static List<systemUser> _systemUser;
        internal static List<systemUser> SystemUser
        {
            get => Helper._systemUser;
            set => Helper._systemUser = value;
        }
        private static List<systemUser> _systemUsergrid;
        private static List<systemUser> SystemUsergrid
        {
            get => Helper._systemUsergrid;
            set => Helper._systemUsergrid = value;
        }
        internal static List<SecurityRole> SecurityRoles
        {
            get => Helper._securityRoles;
            set => Helper._securityRoles = value;
        }
        private static List<Team> _teams;
        internal static List<Team> Teams
        {
            get => Helper._teams;
            set => Helper._teams = value;
        }
        private static List<Queue> _queues;

        public static List<Queue> Queues
        {
            get => Helper._queues;
            set => Helper._queues = value;
        }

        private static IOrganizationService _serviceProxy;
        public static IOrganizationService ServiceProxy
        {
            get => Helper._serviceProxy;
            set => Helper._serviceProxy = value;
        }
        private static string _OrganizationUri;
        public static string OrganizationUri
        {
            get => Helper._OrganizationUri;
            set => Helper._OrganizationUri = value;
        }
        private static ClientCredentials _Credentials = null;
        public static ClientCredentials Credentials
        {
            get => Helper._Credentials;
            set => Helper._Credentials = value;
        }
        internal static void createConn(IOrganizationService Service)
        {

            _serviceProxy = Service;

        }

        public static bool HasValue(this Entity entity, string field)
        {
            if (entity.Attributes.Contains(field))
            {
                return true;
            }

            return false;
        }
        internal static List<systemUser> getUsers()
        {

            SystemUser = new List<systemUser>();
            SystemUsergrid = new List<systemUser>();
            int fetchCount = 5000;// Initialize the page number.
            int pageNumber = 1;// Initialize the number of records.
            string pagingCookie = null;
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='systemuser'>
                  <attribute name='fullname' />
                  <attribute name='businessunitid' />
                  <attribute name='domainname' />
                  <attribute name='systemuserid' />
                  <order attribute='fullname' descending='false' />
                  <filter type='and'>
                      <condition attribute='isdisabled' operator='eq' value='0' />
                      <condition attribute='accessmode' operator='ne' value='3' />
                      <condition attribute='accessmode' operator='ne' value='4' />
                  </filter>
              </entity>
            </fetch>";

            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXML, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection returnCollection = ((RetrieveMultipleResponse)_serviceProxy.Execute(fetchRequest1)).EntityCollection;

                SystemUser.AddRange(
                    returnCollection.Entities
                        .Select(
                            ent => new systemUser
                            {
                                Domainname = ent.Attributes["domainname"].ToString(),
                                FullName = ent.Attributes["fullname"].ToString(),
                                BusinessUnit = ((EntityReference)ent.Attributes["businessunitid"]).Name,
                                SystemUserID = ent.Id
                            }));

                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;
                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }


            return SystemUser;
        }

        internal static void getUserDetail(List<systemUser> destUsers)
        {
            var fetchXml = $@"
<fetch>
  <entity name='systemuser'>
    <attribute name='fullname' />
    <attribute name='systemuserid' />
    <filter type='or'>
      <condition attribute='systemuserid' operator='in'>
        {string.Join(string.Empty, destUsers.Select(usr => "<value>" + usr.SystemUserID + "</value>" + Environment.NewLine))} 
      </condition>
    </filter>
    <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' link-type='outer'>
      <link-entity name='role' from='roleid' to='roleid' intersect='true'>
        <attribute name='name' alias='RoleName'/>
        <attribute name='roleid' alias='RoleID'/>
      </link-entity>
    </link-entity>
    <link-entity name='teammembership' from='systemuserid' to='systemuserid' link-type='outer'>
      <link-entity name='team' from='teamid' to='teamid' link-type='outer' intersect='true'>
        <attribute name='teamid' alias='TeamID'/>
        <attribute name='name' alias='TeamName' />
      </link-entity>
    </link-entity>
    <link-entity name='queuemembership' from='systemuserid' to='systemuserid'>
      <link-entity name='queue' from='queueid' to='queueid' intersect='true'>
        <attribute name='queuetypecode' />
        <attribute name='queueid' alias='QueueID'/>
        <attribute name='name' alias='QueueName'/>
        <attribute name='msdyn_queuetype' />
        <filter>
          <condition attribute='queueviewtype' operator='eq' value='1'/>
          <condition attribute='name' operator='not-like' value='&lt;'/>
        </filter>
      </link-entity>
    </link-entity>
  </entity>
</fetch>";
            int fetchCount = 5000;// Initialize the page number.
            int pageNumber = 1;// Initialize the number of records.
            string pagingCookie = null;

            var returnUsers = new List<systemUser>();
            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXml, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection returnCollection = ((RetrieveMultipleResponse)ServiceProxy.Execute(fetchRequest1)).EntityCollection;

                foreach (Entity entity in returnCollection.Entities)
                {
                    destUsers.First(usr => usr.SystemUserID == entity.Id).AddAssoc(entity);
                }
                //returnUsers.AddRange(
                //    returnCollection.Entities
                //        .Select(
                //            ent => new systemUser
                //            {
                //                Domainname = ent.Attributes["domainname"].ToString(),
                //                FullName = ent.Attributes["fullname"].ToString(),
                //                BusinessUnit = ((EntityReference)ent.Attributes["businessunitid"]).Name,
                //                SystemUserID = ent.Id
                //            }));

                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;
                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }

        }



        internal static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }
        internal static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }
        internal static List<SecurityRole> getSecurityRole(Guid systemUserId)
        {
            EntityCollection returnCollection;
            List<SecurityRole> accessRoles = new List<SecurityRole>();
            int fetchCount = 5000;// Initialize the page number.
            int pageNumber = 1;// Initialize the number of records.
            string pagingCookie = null;
            string fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>"
  + "<entity name='role'>"
    + "<attribute name='name' />"
    + "<attribute name='businessunitid' />"
    + "<attribute name='roleid' />"
    + "<order attribute='name' descending='false' />"
    + "<link-entity name='systemuserroles' from='roleid' to='roleid' visible='false' intersect='true'>"
      + "<link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='ab'>"
        + "<filter type='and'>"
          + "<condition attribute='systemuserid' operator='eq' uitype='systemuser'  value='{" + systemUserId + "}' />"
        + "</filter>"
      + "</link-entity>"
    + "</link-entity>"
  + "</entity>"
+ "</fetch>";
            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXML, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                returnCollection = ((RetrieveMultipleResponse)_serviceProxy.Execute(fetchRequest1)).EntityCollection;
                foreach (Entity ent in returnCollection.Entities)
                {
                    SecurityRole SecurityRole = new SecurityRole();

                    SecurityRole.Name = ent.Attributes["name"].ToString();
                    SecurityRole.BusinessUnit = (EntityReference)ent.Attributes["businessunitid"];
                    SecurityRole.RoleId = ent.Id;
                    accessRoles.Add(SecurityRole);
                }

                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;
                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }


            return accessRoles;
        }
        internal static List<Team> getTeambyUser(Guid systemUserId, string buName)
        {
            List<Team> userTeams = new List<Team>();
            int fetchCount = 5000;// Initialize the page number.
            int pageNumber = 1;// Initialize the number of records.
            string pagingCookie = null;
            string fetchXML =
                "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>"
  + "<entity name='team'>"
    + "<attribute name='name' />"
    + "<attribute name='businessunitid' />"
    + "<attribute name='teamid' />"
    + "<attribute name='teamtype' />"
     + "<filter>"
      + "<condition attribute='name' operator='neq' value='" + buName + "' />"
    + "</filter>"
    + "<order attribute='name' descending='false' />"
    + "<link-entity name='teammembership' from='teamid' to='teamid' visible='false' intersect='true'>"
      + "<link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='ac'>"
        + "<filter type='and'>"
           + "<condition attribute='systemuserid' operator='eq' uitype='systemuser'  value='{" + systemUserId + "}' />"
        + "</filter>"
      + "</link-entity>"
    + "</link-entity>"
  + "</entity>"
+ "</fetch>";
            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXML, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection returnCollection = ((RetrieveMultipleResponse)_serviceProxy.Execute(fetchRequest1)).EntityCollection;
                foreach (Entity ent in returnCollection.Entities)
                {
                    Team team = new Team();
                    team.Name = ent.Attributes["name"].ToString();
                    team.BusinessUnit = (EntityReference)ent.Attributes["businessunitid"];
                    team.TeamId = ent.Id;
                    userTeams.Add(team);
                }

                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;
                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }


            return userTeams;
        }
        internal static List<Queue> getQueuebyUser(Guid systemUserId, string buName)
        {
            List<Queue> userQueues = new List<Queue>();
            RetrieveUserQueuesRequest userQueuesReq = new RetrieveUserQueuesRequest
            {
                UserId = systemUserId,
                IncludePublic = false,
            };
            RetrieveUserQueuesResponse res = (RetrieveUserQueuesResponse)_serviceProxy.Execute(userQueuesReq);
            foreach (Entity ent in res.EntityCollection.Entities)
            {
                if (((OptionSetValue)ent.Attributes["queueviewtype"]).Value.ToString() == "1" && !ent.Attributes["name"].ToString().StartsWith("<"))
                {
                    Queue queue = new Queue();
                    queue.Name = ent.Attributes["name"].ToString();
                    queue.QueueId = ent.Id;
                    userQueues.Add(queue);
                }
            }
            return userQueues;
        }

        internal static void removeuserRole(List<systemUser> destUsers)// Guid[] toUserID, string[] BUs)
        {
            ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
            {
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                Requests = new OrganizationRequestCollection()
            };

            foreach (var user in destUsers)
            {
                if (user.Roles.Any())
                {
                    DisassociateRequest de = new DisassociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("systemuserroles_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (SecurityRole role in user.Roles)
                    {
                        de.RelatedEntities.Add(new EntityReference("role", role.RoleId));

                    }
                    multiReq.Requests.Add(de);
                }
            }
            if (multiReq.Requests.Any())
            {
                ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
            }
        }
        internal static void adduserRole(List<systemUser> destUsers, List<SecurityRole> roles)
        {
            try
            {
                ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                foreach (var user in destUsers)
                {
                    AssociateRequest ar = new AssociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("systemuserroles_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (SecurityRole role in roles)
                    {
                        ar.RelatedEntities.Add(new EntityReference("role", role.RoleId));
                    }
                    multiReq.Requests.Add(ar);

                }

                if (multiReq.Requests.Any())
                {
                    ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }

        }

        internal static void adduserTeam(List<systemUser> destUsers, List<Team> teams)
        {
            try
            {
                ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                foreach (var user in destUsers)
                {
                    AssociateRequest ar = new AssociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("teammembership_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (Team team in teams)
                    {
                        ar.RelatedEntities.Add(new EntityReference("team", team.TeamId));
                    }
                    multiReq.Requests.Add(ar);

                }

                if (multiReq.Requests.Any())
                {
                    ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }
        }
        internal static void removeuserTeam(List<systemUser> destUsers)
        {
            ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
            {
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                Requests = new OrganizationRequestCollection()
            };

            foreach (var user in destUsers)
            {
                if (user.Teams.Any())
                {
                    DisassociateRequest de = new DisassociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("teammembership_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (Team team in user.Teams)
                    {
                        de.RelatedEntities.Add(new EntityReference("team", team.TeamId));

                    }
                    multiReq.Requests.Add(de);
                }
            }
            if (multiReq.Requests.Any())
            {
                ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
            }
        }
        internal static void adduserQueue(List<systemUser> destUsers, List<Queue> queues)
        {
            try
            {
                ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                foreach (var user in destUsers)
                {
                    AssociateRequest ar = new AssociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("queuemembership_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (Queue queue in queues)
                    {
                        ar.RelatedEntities.Add(new EntityReference("team", queue.QueueId));
                    }
                    multiReq.Requests.Add(ar);

                }

                if (multiReq.Requests.Any())
                {
                    ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }
        }

        internal static void removeuserQueue(List<systemUser> destUsers)
        {
            ExecuteMultipleRequest multiReq = new ExecuteMultipleRequest
            {
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                Requests = new OrganizationRequestCollection()
            };

            foreach (var user in destUsers)
            {
                if (user.Teams.Any())
                {
                    DisassociateRequest de = new DisassociateRequest
                    {
                        Target = new EntityReference("systemuser", user.SystemUserID),
                        Relationship = new Relationship("queuemembership_association"),
                        RelatedEntities = new EntityReferenceCollection()
                    };
                    foreach (Queue queue in user.Queues)
                    {
                        de.RelatedEntities.Add(new EntityReference("queue", queue.QueueId));

                    }
                    multiReq.Requests.Add(de);
                }
            }
            if (multiReq.Requests.Any())
            {
                ExecuteMultipleResponse response = (ExecuteMultipleResponse)ServiceProxy.Execute(multiReq);
            }
        }
        internal static void removeuserTeam(Guid[] toUserID, string[] BUs)
        {
            for (int x = 0; x < toUserID.Count(); x++)
            {
                Guid userid = toUserID[x];
                List<Team> userTeamss = getTeambyUser(userid, BUs[x]);
                foreach (Team team in userTeamss)
                {
                    if (team.TeamId != Guid.Empty)
                    {

                        _serviceProxy.Execute(new RemoveMembersTeamRequest
                        {
                            TeamId = team.TeamId,
                            MemberIds = toUserID
                        });

                    }

                }
            }
        }
        private static void adduserTeam(Guid[] toUserID)
        {
            foreach (Team team in Helper.Teams)
            {
                if (team.TeamId != Guid.Empty || toUserID[0] != Guid.Empty)
                {

                    _serviceProxy.Execute(new AddMembersTeamRequest
                    {
                        TeamId = team.TeamId,
                        MemberIds = toUserID
                    });

                }

            }
        }

        internal static void removeuserQueue(Guid[] toUserID, string[] BUs)
        {
            for (int x = 0; x < toUserID.Count(); x++)
            {
                Guid userid = toUserID[x];
                List<Queue> userQueues = getQueuebyUser(userid, BUs[x]);
                foreach (Queue queue in userQueues)
                {
                    _serviceProxy.Disassociate(
                  "systemuser",
                  userid,
                  new Relationship("queuemembership_association"),
                  new EntityReferenceCollection { new EntityReference("queue", queue.QueueId) });

                }


            }
        }
        private static void adduserQueue(Guid[] toUserID)
        {
            try
            {
                foreach (Queue queue in Helper.Queues)
                {

                    foreach (Guid userid in toUserID)
                    {
                        _serviceProxy.Associate(
                "systemuser",
                userid,
                new Relationship("queuemembership_association"),
                new EntityReferenceCollection { new EntityReference("queue", queue.QueueId) });
                    }

                }
            }
            catch (Exception ex)
            {


            }

        }

    }
}
