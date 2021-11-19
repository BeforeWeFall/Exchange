using BR.Core;
using BR.Core.Attributes;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace Activities.Email.Exchange
{
    [LocalizableScreenName(nameof(Resources.GetMessage_Name), typeof(Resources))]
    [LocalizableDescription(nameof(Resources.GetMessage_Description), typeof(Resources))]    
    public class GetMessage : AbstractExchangeActivity
    {
        

        [LocalizableScreenName(nameof(Resources.OldFolderName_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.OldFolderName_Description), typeof(Resources))]
        [IsRequired]
        public string FolderName { get; set; } = "";

        [LocalizableScreenName(nameof(Resources.NewMail_Name), typeof(Resources))]
        [IsCheckBox]
        public bool NewMail { get; set; }

        [LocalizableScreenName(nameof(Resources.EmailMessages_Name), typeof(Resources))]
        [IsRequired]
        [IsOut]
        public List<SerializableEmailMessage> EmailMessages { get; set; }

        [LocalizableScreenName(nameof(Resources.Domain_Name), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Domain_Description), typeof(Resources))]
        public string Domain { get; set; }

        [LocalizableScreenName(nameof(Resources.MailCount_Name), typeof(Resources))]
        public int MailCount { get; set; }

        [LocalizableScreenName(nameof(Resources.UnRead_Name), typeof(Resources))]
        [IsCheckBox]
        public bool UnRead { get; set; }

        [LocalizableScreenName(nameof(Resources.MarkAsRead_Name), typeof(Resources))]
        [IsCheckBox]
        public bool MarkAsRead { get; set; }

        [LocalizableScreenName(nameof(Resources.MarkAsUnRead_Name), typeof(Resources))]
        [IsCheckBox]
        public bool MarkAsUnRead { get; set; }
        public override void Execute(int? optionID)
        {
            EmailMessages = new List<SerializableEmailMessage>();

            WorkWithExchangeService workWithService = new WorkWithExchangeService(Login, Password, Domain, URL);

            var folder = workWithService.SearchOneFolder(FolderName);

            ItemView view = new ItemView(MailCount, 0)//add offset to argument
            {
                PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.IsRead)
            };

            //if (NewMail)
            //    view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            //else
            //    view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

            DataTimeSorting(view, NewMail);

            FindItemsResults<Item> result;

            if (UnRead)
            {
                //SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                //result = workWithService.service.FindItems(folder.Id, searchFilter, view);
                result = FindItemsUnread(workWithService._service, folder.Id, view);
            }
            else
                result = FindItems(workWithService._service, folder.Id, view); /*workWithService.service.FindItems(folder.Id, view);*/

            if (result.TotalCount == 0)
                return;

            workWithService.service.LoadPropertiesForItems(result, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.IsRead, ItemSchema.Body));

            foreach (var res in result.Items)
            {
                EmailMessages.Add(new SerializableEmailMessage(res));
            }

            if (MarkAsRead)
            {
                MarkRead(result.Items);
            }
            if (MarkAsUnRead)
            {
                MarkUnRead(result.Items);
            }
        }

        private void DataTimeSorting(ItemView view, bool Descending) // create datatime serch
        {
            view.OrderBy.Add(ItemSchema.DateTimeReceived, Descending ? SortDirection.Descending : SortDirection.Ascending);
        }
        private void MarkRead(System.Collections.ObjectModel.Collection<Item> items)
        {
            foreach (var r in items)
            {
                ((EmailMessage)r).IsRead = true;
                r.Update(ConflictResolutionMode.AutoResolve); // add variable
            }
        }
        private void MarkUnRead(System.Collections.ObjectModel.Collection<Item> items)
        {
            foreach (var r in items)
            {
                ((EmailMessage)r).IsRead = false;
                r.Update(ConflictResolutionMode.AutoResolve); // add variable
            }
        }
        private FindItemsResults<Item> FindItems(ExchangeService service, FolderId id, ItemView view)
        {
            return service.FindItems(id, view);
        }
        private FindItemsResults<Item> FindItemsUnread(ExchangeService service, FolderId id, ItemView view)
        {
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            return service.FindItems(id, searchFilter, view);
        }
        private void FindItemsBetweenDates(ExchangeService service, FolderId id, ItemView view, SearchFilter searchFilter, DateTime timeStart, DateTime timeFinish)
        {
            searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter, new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeSent, timeFinish),
                new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeSent, timeStart));
            //return service.FindItems(id, searchFilter, view);
        }
        private void FindItemsGreatDate(ExchangeService service, FolderId id, ItemView view, SearchFilter searchFilter, DateTime dateTime)
        {
            searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter, new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeSent, dateTime));
            //return service.FindItems(id, searchFilter, view);
        }
        private void FindItemsLessDate(ExchangeService service, FolderId id, ItemView view, SearchFilter searchFilter, DateTime dateTime)
        {
            searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter, new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeSent, dateTime));
        }
    }
}