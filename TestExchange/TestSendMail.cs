using Microsoft.VisualStudio.TestTools.UnitTesting;
using Activities.Email;
using BR;
using BR.Core.Base;
using BR.Logic;
using BR.Engine;
using System.Threading.Tasks;
using BR.Core;
using Activities.Email.Exchange;

namespace TestExchange
{
    [TestClass]
    public class TestSendMail
    {
        string Login = @"ULBelykh";
        string Password = "12041995Yu";
        string Subject = "Test";
        string Text = "Testtesttest";
        string Recipients = "doginthespace@gmail.com";
        string Domain = "bit";

        static TestSendMail()
        {
            //if (ActivityManager.List.Count is 0)
            //{
            //    ActivityManager.Reload();
            //}
        }

        [TestMethod]
        public void TestMethod1Async()
        {
            SendMail sendMail = new SendMail();
            sendMail.Login = Login;
            sendMail.Password = Password;
            sendMail.Text = Text;
            sendMail.Domain = Domain;
            sendMail.URL = "mail.1cbit.ru";
            sendMail.Recipients = Recipients;


            // sendMail.Execute(default);

            var scriptBuilder = new ScriptBuilder();

            //scriptBuilder
            //    .WithStep<SendMail>().WithPropertyValue("Login", Login)
            //    .WithPropertyValue("Password", Password)
            //    .WithPropertyValue("Subject", Subject)
            //    .WithPropertyValue("Text", Text)
            //    .WithPropertyValue("Recipients", Recipients)
            //    .WithPropertyValue("Domain", Domain);
            //Script script = scriptBuilder.Build();

            //ExecutionManager executionManager = new ExecutionManager();

            //await executionManager.StartAsync(script);

        }
    }
}
