using BR.Core;
using BR.Core.Attributes;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.ComponentModel.DataAnnotations;


namespace Activities.Email.Exchange
{
    [LocalizableScreenName(nameof(Resources.SendMail_Name), typeof(Resources))]
    [LocalizableDescription(nameof(Resources.SendMai_Description), typeof(Resources))]
    [Path("Email.Exchange")]
    public class SendMail : Activity
    {
        [LocalizableScreenName(nameof(Resources.Login_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Login_Description), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string Login { get; set; }

        [LocalizableScreenName(nameof(Resources.Domain_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Domain_Description), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string Domain { get; set; }

        [LocalizableDescription(nameof(Resources.URL_Description), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string URL { get; set; }

        [LocalizableDescription(nameof(Resources.Password_Description), typeof(Resources))]
        [LocalizableScreenName(nameof(Resources.Password_Name), typeof(Resources))]
        [Options(0)]
        [IsRequired]
        public string Password { get; set; }

        [LocalizableScreenName(nameof(Resources.Subject_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Subject_Description), typeof(Resources))]
        public string Subject { get; set; }

        [LocalizableScreenName(nameof(Resources.Text_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Text_Description), typeof(Resources))]
        [IsRequired]
        public string Text { get; set; }

        [LocalizableScreenName(nameof(Resources.Attachments_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Attachments_Description), typeof(Resources))]
        [IsFilePathChooser]
        public string Attachments { get; set; }


        [LocalizableScreenName(nameof(Resources.Html_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Html_Description), typeof(Resources))]
        [IsCheckBox]
        public bool IsHtml { get; set; }

        [LocalizableScreenName(nameof(Resources.Recipients_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Recipients_Description), typeof(Resources))]
        [IsRequired]
        public string Recipients { get; set; }

        [LocalizableScreenName(nameof(Resources.BCC_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.BCC_Description), typeof(Resources))]
        public string BCC { get; set; }

        [LocalizableScreenName(nameof(Resources.CC_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.CC_Description), typeof(Resources))]
        public string CC { get; set; }

        [LocalizableScreenName(nameof(Resources.Categories_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Categories_Description), typeof(Resources))]
        public string Categories { get; set; }

        [LocalizableScreenName(nameof(Resources.SaveCopy_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.SaveCopy_Description), typeof(Resources))]
        [IsCheckBox]
        public bool SaveCopy { get; set; }
        public override void Execute(int? optionID)
        {
            WorkWithExchangeService workWithService = new WorkWithExchangeService(Login, Password, Domain, URL);
            #region Prepear arg
            var recipients = Recipients.Split(';');
            var bcc = string.IsNullOrEmpty(BCC) ? new string[0] : BCC.Split(';');
            var cc = string.IsNullOrEmpty(CC) ? new string[0] : CC.Split(';');
            var attachments = string.IsNullOrEmpty(Attachments) ? new string[0] : Attachments.Split(',');
            var categories = string.IsNullOrEmpty(Categories) ? new string[0] : Categories.Split(',');
            #endregion

            EmailMessage Message = new EmailMessage(workWithService._service);

            Message.Body = Text;
            Message.Body.BodyType = IsHtml ? BodyType.HTML : BodyType.Text;
            Message.Subject = Subject;
            Message.ToRecipients.AddRange(recipients);
            Message.BccRecipients.AddRange(bcc);
            Message.CcRecipients.AddRange(cc);
            foreach (var at in attachments)
                Message.Attachments.AddFileAttachment(at);
            Message.Categories.AddRange(categories);

            if (SaveCopy)
                Message.SendAndSaveCopy();
            else
                Message.Send();
        }

    }
}
