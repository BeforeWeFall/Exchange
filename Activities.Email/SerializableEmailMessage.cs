using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Data;
using ReadSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Activities.Email.Exchange
{
    [Serializable]
    public class SerializableEmailMessage
    {
        public string ReceivedRepresentingAddress { get; private set; }

        public string ReceivedByAddress { get; private set; }

        public string Sender { get; private set; }

        public string[] ReplyTo { get; private set; }

        public string References { get; private set; }

        public string InternetMessageId { get; private set; }

        public bool? IsResponseRequested { get; private set; }

        public bool IsReadReceiptRequested { get; private set; }

        public bool IsRead { get; private set; }
        public string From { get; private set; }

        public byte[] ConversationIndex { get; private set; }

        public string ConversationTopic { get; private set; }

        public string[] CcRecipients { get; private set; }

        public string[] BccRecipients { get; private set; }

        public string[] ToRecipients { get; private set; }

        public string DateTimeSent { get; private set; }

        public string BodyText { get; private set; }
        public string Subject { get; private set; }

        public ItemId Id { get; private set; }
        public SerializableEmailMessage(object obj)
        {
            if (obj is EmailMessage o)
            {
                this.ReceivedRepresentingAddress = o.ReceivedRepresenting.Address;
                this.ReceivedByAddress = o.ReceivedBy.Address;
                this.Sender = o.Sender.Address;
                this.ReplyTo = GetStringArray(o.ReplyTo);
                this.References = o.References;
                this.InternetMessageId = o.InternetMessageId;
                this.IsResponseRequested = o.IsResponseRequested;
                this.IsReadReceiptRequested = o.IsReadReceiptRequested;
                this.IsRead = o.IsRead;
                this.From = o.From.Address;
                this.ConversationIndex = o.ConversationIndex;
                this.ConversationTopic = o.ConversationTopic;
                this.CcRecipients = GetStringArray(o.CcRecipients);
                this.BccRecipients = GetStringArray(o.BccRecipients);
                this.ToRecipients = GetStringArray(o.ToRecipients);
                this.Subject = o.Subject;
            }

            if (obj is Item)
            {
                this.DateTimeSent = ((Item)obj).DateTimeSent.ToString();
                this.BodyText = HtmlUtilities.ConvertToPlainText(((Item)obj).Body.Text);
                this.Id = ((Item)obj).Id;
            }

        }
        private string[] GetStringArray(EmailAddressCollection emailAddresses)
        {
            return emailAddresses.Select(x => x.Address).ToArray();
        }   
    }
}
