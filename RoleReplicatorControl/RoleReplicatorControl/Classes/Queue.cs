using System;

namespace RoleReplicatorControl
{
    public class Queue
    {
        Guid _queueId;

        public Guid QueueId
        {
            get => _queueId;
            set => _queueId = value;
        }

        string _name;

        public string Name
        {
            get => _name;
            set => _name = value;
        }

    }
}
