using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace MyPowershellModule
{
    [Cmdlet(VerbsCommon.Get, "MyMailbox")]
    class GetMyMailbox : Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, HelpMessage = "Provide mailbox identity")]
        public string Identity;

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Provide organization")]
        public string Organization;

        protected override void BeginProcessing() //Used for initializing resources
        {
            ValidateParameters(); //We want to validate our parameters are set
        }

        protected override void ProcessRecord()
        {
            string firstGuid = string.Empty;

            // We use the current runspace to cast our calls into.
            // This grants us access to the current runspace, instead of instantiating a new one.
            // More information on the PowerShell.Create method can be found here: https://msdn.microsoft.com/en-us/library/system.management.automation.powershell.create(v=vs.85).aspx
            // More information on the RunspaceMode.CurrentRunspace enumeration value can be found here: https://msdn.microsoft.com/en-us/library/system.management.automation.runspacemode(v=vs.85).aspx
            PowerShell myinitialPowerShell = PowerShell.Create(RunspaceMode.CurrentRunspace);
            myinitialPowerShell.AddCommand("Get-Mailbox");
            myinitialPowerShell.AddParameter("Identity", this.Identity);
            myinitialPowerShell.AddParameter("Organization", this.Organization);

            Collection<PSObject> resultsCollection = myinitialPowerShell.Invoke();
            this.WriteObject(resultsCollection);
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
            if (String.IsNullOrEmpty(Identity))
                ThrowParameterError("Identity");

            if (String.IsNullOrEmpty(Organization))
                ThrowParameterError("Organization");
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
