using BR.Core;
using BR.Core.Attributes;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Activities.Email.Exchange
{
    [BR.Core.Attributes.Path("Email.Exchange")]
    public class GetAttachments : Activity
    {
        [LocalizableScreenName(nameof(Resources.Login_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Login_Description), typeof(Resources))]
        [IsRequired]
        public string Login { get; set; }

        [LocalizableScreenName(nameof(Resources.Password_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Password_Description), typeof(Resources))]
        [IsRequired]
        public string Password { get; set; }

        [LocalizableDescription(nameof(Resources.URL_Description), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string URL { get; set; }

        [LocalizableScreenName(nameof(Resources.Domain_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Domain_Description), typeof(Resources))]
        public string Domain { get; set; }

        //[LocalizableScreenName(nameof(Resources.OldFolderName_Name), typeof(Resources))]
        //[LocalizableDescription(nameof(Resources.OldFolderName_Description), typeof(Resources))]
        //[IsRequired]
        //public string FromFolderName { get; set; }

        [LocalizableScreenName(nameof(Resources.FolderToSave_Name), typeof(Resources))]
        [IsRequired]
        public string FolderToSave { get; set; }


        [LocalizableScreenName(nameof(Resources.Message_Name), typeof(Resources))]
        [IsRequired]
        public SerializableEmailMessage Message { get; set; }
        public override void Execute(int? optionID)
        {
            WorkWithExchangeService workWithService = new WorkWithExchangeService(Login, Password, Domain, URL);
            string pathToSave = FolderToSave[FolderToSave.Length - 1] == '\\' ? FolderToSave : FolderToSave + @"\";
            if (!Directory.Exists(pathToSave))
            {
                try
                {
                    Directory.CreateDirectory(pathToSave);
                }
                catch
                {
                    throw new Exception("Не удалось создать папку, проверьте указанный путь"); // разные языки
                }
            }

            //var folderOld = facad.SearchFolder(FromFolderName);

            //ItemView view = new ItemView(1, 0)
            //{
            //    PropertySet = new PropertySet(ItemSchema.Id)
            //};

            //FindItemsResults<Item> results;

            //SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

            //results = facad.service.FindItems(folderOld.Id, searchFilter, view);

            EmailMessage current = EmailMessage.Bind(workWithService._service, Message.Id); // тут проверку на null
            //var result = results.Items[0];

            Save(current, pathToSave);
        }

        private void Save(Item item, string pathToSave)
        {
            foreach (Attachment attachment in item.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;

                    fileAttachment.Load();
                    fileAttachment.Load(pathToSave + fileAttachment.Name);

                }
                else
                {
                    ItemAttachment itemAttachment = attachment as ItemAttachment;

                    itemAttachment.Load();
                    string fileName = pathToSave + attachment.Name;

                    File.WriteAllBytes(fileName, itemAttachment.Item.MimeContent.Content);
                }
            }
        }

    }
}