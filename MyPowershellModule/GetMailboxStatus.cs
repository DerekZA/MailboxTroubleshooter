using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Host;
using System.Collections.ObjectModel;

namespace MyPowershellModule
{
    [Cmdlet(VerbsCommon.Get, "MailboxStatus")]
    public class GetMailboxStatus : Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, HelpMessage = "Provide a users MailboxGuid value")]
        public string MailboxGuid;

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Provide a database")]
        public string Database;

        protected override void BeginProcessing() //Used for initializing resources
        {
            try
            {

                ValidateParameters(); //We want to validate our parameters are set

                PowerShell importScripts = PowerShell.Create(RunspaceMode.CurrentRunspace);

                //Import the the PS1 allowing us to expose the Get-StoreQuery cmdlet
                string script = @".\ManagedStoreDiagnosticFunctions.ps1";
                importScripts.AddScript(script);
                var results = importScripts.Invoke();
                WriteObject(results);
            }
            catch (Exception)
            {
                throw;
            }
        }

        protected override void ProcessRecord()
        {
            PowerShell myfirstPowerShell = PowerShell.Create(RunspaceMode.CurrentRunspace);
            string query = String.Format("select * from [{0}].Mailbox where MailboxGuid = '{1}'", this.Database, this.MailboxGuid);

            // We add the command we want to run.
            // More information on the PowerShell.AddCommand method can be found here: https://msdn.microsoft.com/en-us/library/dd182430(v=vs.85).aspx
            myfirstPowerShell.AddCommand("Get-StoreQuery");

            // We add the parameter that we want to specify.
            // More information on the PowerShell.AddParameter method can be found here: https://msdn.microsoft.com/en-us/library/dd182434(v=vs.85).aspx

            myfirstPowerShell.AddParameter("Database", this.Database).AddArgument(query);

            // We add the command we want to run.
            // More information on the PowerShell.AddCommand method can be found here: https://msdn.microsoft.com/en-us/library/dd182430(v=vs.85).aspx
            myfirstPowerShell.AddCommand("Select-Object");

            // We select the specific property we want to return.
            // More information on the PowerShell.AddParameter method can be found here: https://msdn.microsoft.com/en-us/library/dd182434(v=vs.85).aspx
            myfirstPowerShell.AddParameter("Property", new string[] { "MailboxGuid","MailboxNumber","DisplayName","DeletedOn","Status" });

            Collection<PSObject> resultsCollection = myfirstPowerShell.Invoke();
            PSObject myFirstPsObject = resultsCollection.FirstOrDefault();

            if (myFirstPsObject != null)
            {
                //Write 
                WriteObject(myFirstPsObject);
            }
        }

        protected override void EndProcessing() //Executes after Begin Processing
        {

        }

        protected override void StopProcessing() //Called when your cmdlet execution is interrupted. Use to clean up resources
        {
        }

        /// <summary>
        /// Validates Cmdlet parameters prior to processing the record
        /// </summary>
        private void ValidateParameters()
        {
            if (String.IsNullOrEmpty(MailboxGuid))
                ThrowParameterError("MailboxGuid");

            if (String.IsNullOrEmpty(Database))
                ThrowParameterError("Database");
        }
        private void ThrowParameterError(string parameterName)
        {
            ThrowTerminatingError(
                new ErrorRecord(
                    new ArgumentException(String.Format(
                        "Must pecifify '{0}'", parameterName)),
                    Guid.NewGuid().ToString(),
                    ErrorCategory.InvalidArgument,
                    null));
        }
    }
}