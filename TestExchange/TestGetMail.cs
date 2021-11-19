using BR;
using BR.Core.Base;
using BR.Logic;
using BR.Engine;
using System.Threading.Tasks;
using BR.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Activities.Email;
using Activities.Email.Exchange;
using System.Collections.Generic;

namespace TestExchange
{
    [TestClass]
    public class TestGetMail
    {
        string Login = @"ULBelykh";
        string Password = "12041995Yu";
        string Domain = "bit";
        string Url = "mail.1cbit.ru";
        string FolderName = "В";

        [TestMethod]
        public void GetRead()
        {
            List<SerializableEmailMessage> list = new List<SerializableEmailMessage>();

            GetMessage getMail = new GetMessage();
            getMail.Login = Login;
            getMail.MailCount = 1;
            getMail.Password = Password;
            getMail.Domain = Domain;
            getMail.FolderName = FolderName;
            getMail.URL = Url;
            getMail.UnRead = false;
            getMail.NewMail = true;
            getMail.EmailMessages = list;


            getMail.Execute(default);

            //MoveMail moveMail = new MoveMail();
            //moveMail.Login = Login;

            //moveMail.Password = Password;
            //moveMail.Domain = Domain;
            //moveMail.NewFolderName = "В";
            //moveMail.URL = Url;
            //moveMail.Message = getMail.EmailMessages[0];
            //moveMail.Execute(default);

            Assert.AreEqual("Произошла не обработанная ошибка", getMail.EmailMessages[0].Subject);
        }
        [TestMethod]
        public void GetUnRead()
        {
            List<SerializableEmailMessage> list = new List<SerializableEmailMessage>();

            GetMessage getMail = new GetMessage();
            getMail.Login = Login;
            getMail.MailCount = 1;
            getMail.Password = Password;
            getMail.Domain = Domain;
            getMail.FolderName = FolderName;
            getMail.URL = Url;
            getMail.UnRead = true;
            getMail.NewMail = true;
            getMail.EmailMessages = list;

            getMail.Execute(default);

            Assert.AreEqual("RE: приказ о зачислении 00000003651", getMail.EmailMessages[0].Subject);
        }
        [TestMethod]
        public void GetReadEnd()
        {
            List<SerializableEmailMessage> list = new List<SerializableEmailMessage>();

            GetMessage getMail = new GetMessage();
            getMail.Login = Login;
            getMail.MailCount = 1;
            getMail.Password = Password;
            getMail.Domain = Domain;
            getMail.FolderName = FolderName;
            getMail.URL = Url;
            getMail.UnRead = false;
            getMail.NewMail = false;
            getMail.EmailMessages = list;

            getMail.Execute(default);

            Assert.AreEqual("RE: СберКУ-1Бит(RPA)", getMail.EmailMessages[0].Subject);
        }
        [TestMethod]
        public void GetUnReadEnd()
        {
            List<SerializableEmailMessage> list = new List<SerializableEmailMessage>();

            GetMessage getMail = new GetMessage();
            getMail.Login = Login;
            getMail.MailCount = 1;
            getMail.Password = Password;
            getMail.Domain = Domain;
            getMail.FolderName = FolderName;
            getMail.URL = Url;
            getMail.UnRead = true;
            getMail.NewMail = false;
            getMail.EmailMessages = list;

            getMail.Execute(default);

            Assert.AreEqual("Варианты поля \"Статус документа\"", getMail.EmailMessages[0].Subject);
        }
    }
}
