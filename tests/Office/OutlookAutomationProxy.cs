// Test-only CLR surface for the existing mutable PowerShell Outlook fixtures.
// The published module performs planning; this bridge supplies objects, not policy.
using System;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace Winkit.Tests
{
    public sealed class OutlookAutomationProxy
    {
        private readonly PSObject value;
        private readonly PSObject session;
        private readonly object[] identities;

        public OutlookAutomationProxy(object value, object session, object[] identities)
        {
            this.value = PSObject.AsPSObject(value);
            this.session = PSObject.AsPSObject(session);
            this.identities = identities ?? new object[0];
        }

        private object Read(string name)
        {
            PSPropertyInfo property = value.Properties[name];
            if (property == null) throw new MissingMemberException(name);
            object result = property.Value;
            while (result is PSObject && !(((PSObject)result).BaseObject is PSCustomObject))
                result = ((PSObject)result).BaseObject;
            return result;
        }

        private object Wrap(object item)
        {
            if (item == null) return null;
            PSObject wrapped = PSObject.AsPSObject(item);
            if (wrapped.BaseObject is PSCustomObject)
                return new OutlookAutomationProxy(wrapped, session, identities);
            return wrapped.BaseObject;
        }

        private object Call(string name, params object[] arguments)
        {
            PSMethodInfo method = value.Methods[name];
            if (method == null) throw new MissingMethodException(name);
            return Wrap(method.Invoke(arguments));
        }

        public string Name { get { return (string)Read("Name"); } }
        public string FolderPath { get { return (string)Read("FolderPath"); } }
        public string StoreID { get { return (string)Read("StoreID"); } }
        public string EntryID { get { return (string)Read("EntryID"); } }
        public string FilePath { get { return (string)Read("FilePath"); } }
        public string DisplayName { get { return (string)Read("DisplayName"); } }
        public int DefaultItemType { get { return Convert.ToInt32(Read("DefaultItemType")); } }
        public int Count { get { return Convert.ToInt32(Read("Count")); } }
        public object Folders { get { return Wrap(Read("Folders")); } }
        public object Items { get { return Read("Items"); } }
        public object Values { get { return Read("Values"); } }
        public object Stores { get { return Wrap(Read("Stores")); } }
        public object DefaultStore { get { return Wrap(Read("DefaultStore")); } }
        public object Root { get { return Wrap(Read("Root")); } }
        public int ExchangeStoreType { get { return Convert.ToInt32(Read("ExchangeStoreType")); } }
        public int Added { get { return Convert.ToInt32(Read("Added")); } }
        public object Detached { get { return Read("Detached"); } }
        public object PropertyAccessor { get { return Wrap(Read("PropertyAccessor")); } }
        public object Store { get { return Wrap(session.Properties["DefaultStore"].Value); } }
        public object Application { get { return this; } }
        // Planning fixtures supply explicit standard identities; use GetDefaultFolder.
        // Outlook 2007 MAPI discovery is tested separately by the module's fixtures.
        public string Version { get { return "16.0"; } }
        public object Item(int index) { return Call("Item", index); }
        public object GetProperty(string name) { return Call("GetProperty", name); }
        public object GetFolderFromID(string entryId, string storeId)
        {
            return Call("GetFolderFromID", entryId, storeId);
        }

        public object GetStoreFromID(string storeId) { return Call("GetStoreFromID", storeId); }
        public object GetRootFolder() { return Call("GetRootFolder"); }
        public void AddStore(string path) { Call("AddStore", path); }
        public void RemoveStore(object root) { Call("RemoveStore", root); }

        public object GetDefaultFolder(int kind)
        {
            string[] names = {
                "DeletedItems", "Outbox", "SentItems", "Inbox", "Calendar", "Contacts",
                "Journal", "Notes", "Tasks", "Drafts", "AllPublicFolders", "Conflicts",
                "SyncIssues", "LocalFailures", "ServerFailures", "Junk", "RssFeeds",
                "ToDo", "ManagedEmail", "SuggestedContacts"
            };
            int[] kinds = { 3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 18, 19, 20, 21, 22, 23, 25, 28, 29, 30 };
            int index = Array.IndexOf(kinds, kind);
            foreach (object item in identities)
            {
                PSObject identity = PSObject.AsPSObject(item);
                if (index < 0 || (string)identity.Properties["Kind"].Value != names[index]) continue;
                string state = (string)identity.Properties["State"].Value;
                if (state == "Absent") return null;
                if (state != "Resolved") throw new COMException("Synthetic unavailable identity", unchecked((int)0x80004004));
                string entryId = (string)identity.Properties["EntryID"].Value;
                PSPropertyInfo storeProperty = identity.Properties["StoreID"];
                string storeId = storeProperty == null ? StoreID : (string)storeProperty.Value;
                return Wrap(session.Methods["GetFolderFromID"].Invoke(entryId, storeId));
            }
            return null;
        }
    }
}
