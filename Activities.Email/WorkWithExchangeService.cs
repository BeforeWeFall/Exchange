using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Activities.Email
{
    internal class WorkWithExchangeService
    {
        private ExchangeService _service { get; }

        public WorkWithExchangeService(string Login, string Password, string domain, string url)
        {
            _service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            _service.Credentials = new WebCredentials(Login, Password, domain ?? "");

            _service.Url =   new Uri("https://" + url + "/EWS/Exchange.asmx");

            //Path.Combine("", url, )

            //service.AutodiscoverUrl(Login);

        }

        public WorkWithExchangeService(ExchangeService service)
        {
            this._service = service;

            //service.AutodiscoverUrl(Login);
        }
        public Folder SearchOneFolder(string folderName)
        {
            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folderName);

            FolderView folderView = new FolderView(1)
            {
                Traversal = FolderTraversal.Deep
            };

            FindFoldersResults results = _service.FindFolders(WellKnownFolderName.Root, searchFilter, folderView);
            return results.Folders.Count > 0 ? results.Folders[0] : throw new Exception("Не удалось найти указанную папку"); // локализацию добавь
        }

    }
}
