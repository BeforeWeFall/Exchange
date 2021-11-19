using Activities.Email.Exchange;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExchange
{
    [TestClass]
    public class TestGetAttachments
    {

        string Login = @"ULBelykh";
        string Password = "12041995Yu";
        string Domain = "bit";
        string Url = "mail.1cbit.ru";
        string FolderName = "В";
        string PathToFolder = @"D:\Тест";

        [TestMethod]
        public void GetFiles()
        {
            List<SerializableEmailMessage> list = new List<SerializableEmailMessage>();

            var files =Directory.GetFiles(PathToFolder);
            foreach (var f in files)
                File.Delete(f);

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

            GetAttachments getAttachments = new GetAttachments();
            getAttachments.Message = getMail.EmailMessages[0];
            getAttachments.FolderToSave = PathToFolder;
            getAttachments.Login = Login;
            getAttachments.Password = Password;
            getAttachments.Domain = Domain;
            getAttachments.URL = Url;

            getAttachments.Execute(default);

            Assert.AreEqual(1, Directory.GetFiles(PathToFolder).Count());
        }
    }
}
