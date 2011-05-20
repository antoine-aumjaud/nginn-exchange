using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeIntegration.Interfaces;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using NGinnBPM.MessageBus;
using NGinnBPM.MessageBus.Impl;
using NLog;
using System.Threading;

namespace ExchangeIntegration.Service
{
    /// <summary>
    /// Receives all email messages by polling a
    /// specified folder periodically.
    /// </summary>
    public class PollingEmailReceiver : IStartableService
    {
        public string User { get; set; }
        public string Password { get; set; }
        public string ImpersonatedUser { get; set; }
        public string EWSUrl { get; set; }
        public string IncomingFolder { get; set; }
        public bool DeleteRetrievedEmails { get; set; }
        public string MoveToFolder { get; set; }
        public string TargetEndpoint { get; set; }
        public TimeSpan MessageRetentionPeriod { get; set; }
        public int PollingIntervalSeconds { get; set; }
        public IMessageBus MessageBus { get; set; }
        public WellKnownFolderName Root { get; set; }

        private string _fromFolderId;
        private string _moveToFolderId;
        private System.Threading.AutoResetEvent _waitEvent;
        private IExchangeConnect exConnect;
        private Logger log = LogManager.GetCurrentClassLogger();

        public PollingEmailReceiver()
        {
            Root = WellKnownFolderName.MsgFolderRoot;
        }

        void PollProc()
        {
            ExchangeService es;
            
            //var f = Folder.Bind(es, WellKnownFolderName.Root);

            

        }

        private Folder FindFolder(Folder searchRoot, string name)
        {
            Folder ret = null;
            var v = new FolderView(100);
            v.Traversal = FolderTraversal.Shallow;
            var r = searchRoot.FindFolders(new SearchFilter.Exists(FolderSchema.DisplayName), v);
            log.Info("Searching for {0} in {1}", name, searchRoot.DisplayName);
            foreach (Folder f in r.Folders)
            {
                log.Info("{0}/{1}", searchRoot.DisplayName, f.DisplayName);
                if (f.DisplayName == name)
                    ret = f;
            }
            return ret;
        }

        private Folder FindFolderByPath(Folder searchRoot, string path, bool create)
        {
            if (string.IsNullOrEmpty(path) || path == ".")
                return searchRoot;
            if (path.StartsWith("/"))
                path = path.Substring(1);
            var v = new FolderView(100);
            v.Traversal = FolderTraversal.Shallow;
            int idx = path.IndexOf('/');
            if (idx < 0) idx = path.Length;
            string name = path.Substring(0, idx);
            Folder f = FindFolder(searchRoot, name);
            if (f == null && create)
            {
                f = new Folder(searchRoot.Service);
                f.DisplayName = name;
                f.Save(searchRoot.Id);
            }
            if (f == null) return null;
            if (path.Length <= idx) return f;
            return FindFolderByPath(f, path.Substring(idx + 1), create);
        }


        public bool IsRunning
        {
            get { throw new NotImplementedException(); }
        }

        public void Start()
        {
            exConnect = new ExchangeConnect
            {
                User = User,
                Password = Password,
                EWSUrl = EWSUrl
            };

            var es = exConnect.Connect();
            Folder f = Folder.Bind(es, Root);
            var f1 = FindFolderByPath(f, IncomingFolder, false);
            if (f1 == null)
                throw new Exception("Input folder not found: " + IncomingFolder);
            _fromFolderId = f1.Id.UniqueId;
            f1 = FindFolderByPath(f, MoveToFolder, true);
            if (f1 == null)
                throw new Exception("Destination folder not found: " + MoveToFolder);
            _moveToFolderId = f1.Id.UniqueId;

            var inf = Folder.Bind(es, new FolderId(_fromFolderId));
            ItemView iv = new ItemView(100);
            iv.Traversal = ItemTraversal.Shallow;
            var r = inf.FindItems(iv);
            log.Info("Found {0} items", r.Items.Count);
            foreach (var it in r.Items)
            {
                ProcessItem(it);
            }
            log.Info("Finished");
        }

        protected void ProcessItem(Item it)
        {
            it.Move(new FolderId(_moveToFolderId));
            MessageBus.NewMessage(new DeleteItem { ItemId = it.Id.UniqueId }).SetDeliveryDate(DateTime.Now + MessageRetentionPeriod).Publish();
        }

        public void Stop()
        {
            
        }
    }
}
