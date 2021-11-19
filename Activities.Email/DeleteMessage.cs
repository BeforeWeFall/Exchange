using BR.Core;
using BR.Core.Attributes;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace Activities.Email.Exchange
{
    [Path("Email.Exchange")]
    public class DeleteMessage : Activity
    {
        [LocalizableScreenName(nameof(Resources.Login_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Login_Description), typeof(Resources))]
        [IsRequired]
        public string Login { get; set; }

        [LocalizableDescription(nameof(Resources.Password_Description), typeof(Resources))]
        [LocalizableScreenName(nameof(Resources.Password_Name), typeof(Resources))]
        [IsRequired]
        public string Password { get; set; }

        [LocalizableDescription(nameof(Resources.URL_Description), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string URL { get; set; }

        [LocalizableDescription(nameof(Resources.MoveToDeleteItems_Name), typeof(Resources))]
        [IsCheckBox]
        public bool MoveToDeleteItems { get; set; }

        [LocalizableScreenName(nameof(Resources.Message_Name), typeof(Resources))]
        [IsRequired]
        public SerializableEmailMessage Message { get; set; }

        [LocalizableScreenName(nameof(Resources.Domain_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Domain_Description), typeof(Resources))]
        public string Domain { get; set; }

        public override void Execute(int? optionID)
        {
            WorkWithExchangeService workWithService = new WorkWithExchangeService(Login, Password, Domain, URL);

            EmailMessage current = EmailMessage.Bind(workWithService._service, Message.Id);

            current.Delete(MoveToDeleteItems ? DeleteMode.SoftDelete : DeleteMode.HardDelete);
        }
    }
}