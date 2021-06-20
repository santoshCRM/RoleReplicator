using Microsoft.Xrm.Sdk;
using System;

namespace RoleReplicatorControl
{
    internal class Team
    {


        Guid _teamId;

        public Guid TeamId
        {
            get => _teamId;
            set => _teamId = value;
        }


        string _name;

        public string Name
        {
            get => _name;
            set => _name = value;
        }
        private EntityReference _businessUnit;

        public EntityReference BusinessUnit
        {
            get => _businessUnit;
            set => _businessUnit = value;
        }
    }
}
