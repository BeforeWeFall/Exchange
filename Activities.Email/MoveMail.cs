using BR.Core;
using BR.Core.Attributes;
using Microsoft.Exchange.WebServices.Data;
using System.ComponentModel.DataAnnotations;

namespace Activities.Email.Exchange
{
    [LocalizableScreenName(nameof(Resources.MoveMail_Name), typeof(Resources))]
    [LocalizableDescription(nameof(Resources.MoveMail_Description), typeof(Resources))]
    [Path("Email.Exchange")]
    public class MoveMail : Activity
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

        [LocalizableScreenName(nameof(Resources.NewFolderName_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.NewFolderName_Description), typeof(Resources))]
        [IsRequired]
        public string NewFolderName { get; set; }

        [LocalizableScreenName(nameof(Resources.Message_Name), typeof(Resources))]
        [IsRequired]
        public SerializableEmailMessage Message { get; set; }
        public override void Execute(int? optionID)
        {
            WorkWithExchangeService workWithService = new WorkWithExchangeService(Login, Password, Domain, URL);

            var folderNew = workWithService.SearchOneFolder(NewFolderName);

            EmailMessage current = EmailMessage.Bind(workWithService._service, Message.Id);
            current.Move(folderNew.Id);
        }
    }
}