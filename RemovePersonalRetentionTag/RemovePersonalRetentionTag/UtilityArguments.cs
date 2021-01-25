// RemovePrivateFlag
//
// Author: Torsten Schlopsnies, Thomas Stensitzki
//
// Based on: http://dotnetfollower.com/wordpress/2012/03/c-simple-command-line-arguments-parser/
//
// Published under MIT license

using Microsoft.Exchange.WebServices.Data;

namespace UtilityArguments
{
    /// <summary>
    /// Description of UtilityArguments.
    /// </summary>
    public class UtilityArguments : InputArguments
    {
        protected bool GetSwitchValue(string key)
        {
            if (ContainsKey(key, out _))
            {
                return true;
            }
            return false;
        }

        public bool Help
        {
            get
            {
                return GetSwitchValue("-help");
            }
        }
        
        public string Usage
        {
            get
            {
                return GetValue("-usage");
            }
        }

        public string Mailbox
        {
            get
            {
                return GetValue("-mailbox");
            }
        }
        
        public string Foldername
        {
            get
            {
                return GetValue("-foldername");
            }
        }

        public bool LogOnly
        {
            get
            {
                return GetSwitchValue("-logonly");
            }
        }

        public bool IgnoreCertificate
        {
            get
            {
                return GetSwitchValue("-ignorecertificate");
            }
        }

        public string URL
        {
            get
            {
                return GetValue("-url");
            }
        }

        public bool AllowRedirection
        {
            get
            {
                return GetSwitchValue("-allowredirection");
            }
        }

        public string User
        {
            get
            {
                return GetValue("-user");
            }
        }

        public string Password
        {
            get
            {
                return GetValue("-password");
            }
        }

        public bool Impersonate
        {
            get
            {
                return GetSwitchValue("-impersonate");
            }
        }

        public bool Archive
        {
            get
            {
                return GetSwitchValue("-archive");
            }
        }

        public string RetentionId
        {
            get
            {
                return GetValue("-retentionid");
            }
        }

        public UtilityArguments(string[] args) : base(args)
        {
        }
    }
}