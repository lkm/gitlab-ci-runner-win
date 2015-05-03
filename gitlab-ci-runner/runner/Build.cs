using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using gitlab_ci_runner.api;
using Microsoft.Experimental.IO;
using System.Management.Automation;
using System.Threading;
using System.Collections.ObjectModel;

namespace gitlab_ci_runner.runner
{
    class Build
    {
        /// <summary>
        /// Build completed?
        /// </summary>
        public bool completed { get; private set; }

        /// <summary>
        /// Command output
        /// Build internal!
        /// </summary>
        private ConcurrentQueue<string> outputList;

        /// <summary>
        /// Command output
        /// </summary>
        public string output
        {
            get
            {
                string t;
                while (outputList.TryPeek(out t) && string.IsNullOrEmpty(t))
                {
                    outputList.TryDequeue(out t);
                }
                return String.Join("\n", outputList.ToArray()) + "\n";
            }
        }

        /// <summary>
        /// Projects Directory
        /// </summary>
        private string sProjectsDir = Program.HomePath + @"\projects";

        /// <summary>
        /// Project Directory
        /// </summary>
        private string sProjectDir;

        /// <summary>
        /// Build Infos
        /// </summary>
        public BuildInfo buildInfo;

        /// <summary>
        /// Command list
        /// </summary>
        private LinkedList<string> commands;

        /// <summary>
        /// Execution State
        /// </summary>
        public State state = State.WAITING;

        /// <summary>
        /// Command Timeout
        /// </summary>
        public int iTimeout
        {
            get
            {
                return this.buildInfo.timeout;
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="buildInfo">Build Info</param>
        public Build(BuildInfo buildInfo)
        {
            this.buildInfo = buildInfo;
            sProjectDir = sProjectsDir + @"\project-" + buildInfo.project_id;
            commands = new LinkedList<string>();
            outputList = new ConcurrentQueue<string>();
            completed = false;
        }

        /// <summary>
        /// Run the Build Job
        /// </summary>
        public void run()
        {
            state = State.RUNNING;

            try
            {

                // Initialize project dir
                initProjectDir();

                // Add build commands

                if (buildInfo.runAsPowershell())
                {
                    var powerScript = string.Join(System.Environment.NewLine, commands.ToArray());
                    powerScript = powerScript.Replace("&&", System.Environment.NewLine);
                    powerScript += System.Environment.NewLine + System.Environment.NewLine + buildInfo.commands;

                    var script = preparePowerShellScript(powerScript);
                    if (!execPowerShell(script))
                    {
                        state = State.FAILED;
                    }
                }
                else
                {
                    foreach (string sCommand in buildInfo.GetCommands())
                    {
                        commands.AddLast(sCommand);
                    }

                    // Execute
                    foreach (string sCommand in commands)
                    {
                        if (!exec(sCommand))
                        {
                            state = State.FAILED;
                            break;
                        }
                    }
                }

                if (state == State.RUNNING)
                {
                    state = State.SUCCESS;
                }

            }
            catch (Exception rex)
            {
                outputList.Enqueue("");
                outputList.Enqueue("A runner exception occoured: " + rex.Message);
                outputList.Enqueue("");
                state = State.FAILED;
            }


            completed = true;
        }

        /// <summary>
        /// Initialize project dir and checkout repo
        /// </summary>
        private void initProjectDir()
        {
            // Check if projects directory exists
            if (!Directory.Exists(sProjectsDir))
            {
                // Create projects directory
                Directory.CreateDirectory(sProjectsDir);
            }

            // Check if already a git repo
            if (Directory.Exists(sProjectDir + @"\.git") && buildInfo.allow_git_fetch)
            {
                // Already a git repo, pull changes
                commands.AddLast(fetchCmd());
                commands.AddLast(checkoutCmd());
            }
            else
            {
                // No git repo, checkout
                if (Directory.Exists(sProjectDir))
                    DeleteDirectory(sProjectDir);

                commands.AddLast(cloneCmd());
                commands.AddLast(checkoutCmd());
            }
        }

        /// <summary>
        /// Execute a single command
        /// </summary>
        /// <param name="sCommand">Command to execute</param>
        private bool exec(string sCommand)
        {
            try
            {
                // Remove Whitespaces
                sCommand = sCommand.Trim();

                // Output command
                outputList.Enqueue("");
                outputList.Enqueue(sCommand);
                outputList.Enqueue("");

                // Build process
                Process p = new Process();
                p.StartInfo.UseShellExecute = false;
                if (Directory.Exists(sProjectDir))
                {
                    p.StartInfo.WorkingDirectory = sProjectDir; // Set Current Working Directory to project directory
                }
                p.StartInfo.FileName = "cmd.exe"; // use cmd.exe so we dont have to split our command in file name and arguments
                p.StartInfo.Arguments = "/C \"" + sCommand + "\""; // pass full command as arguments

                // Environment variables
                p.StartInfo.EnvironmentVariables["HOME"] = Program.HomePath; // Fix for missing SSH Key

                p.StartInfo.EnvironmentVariables["BUNDLE_GEMFILE"] = sProjectDir + @"\Gemfile";
                p.StartInfo.EnvironmentVariables["BUNDLE_BIN_PATH"] = "";
                p.StartInfo.EnvironmentVariables["RUBYOPT"] = "";

                p.StartInfo.EnvironmentVariables["CI_SERVER"] = "yes";
                p.StartInfo.EnvironmentVariables["CI_SERVER_NAME"] = "GitLab CI";
                p.StartInfo.EnvironmentVariables["CI_SERVER_VERSION"] = null; // GitlabCI Version
                p.StartInfo.EnvironmentVariables["CI_SERVER_REVISION"] = null; // GitlabCI Revision

                p.StartInfo.EnvironmentVariables["CI_BUILD_REF"] = buildInfo.sha;
                p.StartInfo.EnvironmentVariables["CI_BUILD_REF_NAME"] = buildInfo.ref_name;
                p.StartInfo.EnvironmentVariables["CI_BUILD_ID"] = buildInfo.id.ToString();

                // Redirect Standard Output and Standard Error
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.OutputDataReceived += new DataReceivedEventHandler(outputHandler);
                p.ErrorDataReceived += new DataReceivedEventHandler(outputHandler);

                try
                {
                    // Run the command
                    p.Start();
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();

                    if (!p.WaitForExit(iTimeout * 1000))
                    {
                        p.Kill();
                    }
                    return p.ExitCode == 0;
                }
                finally
                {
                    p.OutputDataReceived -= new DataReceivedEventHandler(outputHandler);
                    p.ErrorDataReceived -= new DataReceivedEventHandler(outputHandler);
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Execute a script
        /// </summary>
        /// <param name="script">Script to execute</param>
        private bool execPowerShell(string script)
        {
            try
            {
                using (PowerShell p = PowerShell.Create())
                {
                    p.AddScript(script);

                    PSDataCollection<PSObject> outputCollection = new PSDataCollection<PSObject>();
                    outputCollection.DataAdded += outputCollection_DataAdded;

                    p.Streams.Error.DataAdded += outputCollection_DataAdded;
                    p.Streams.Warning.DataAdded += outputCollection_DataAdded;
                    
                    PSInvocationSettings settings = new PSInvocationSettings();
                    settings.ErrorActionPreference = ActionPreference.Stop;
                    settings.ExposeFlowControlExceptions = true;
                    IAsyncResult result = p.BeginInvoke<PSObject, PSObject>(null, outputCollection, settings, null, null);


                    while (result.IsCompleted == false)
                    {
                        Thread.Sleep(1000);
                    }
                    var errors = p.Streams.Error.ReadAll();
                    return p.HadErrors == false;
                }
            }
            catch
            {
                return false;
            }
        }

        private string preparePowerShellScript(string script)
        {
            string file = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".ps1";

            using (StreamWriter fileWriter = new StreamWriter(file))
            {
                string output;

                // Halt on all errors
                output = "$ErrorActionPreference = \"Stop\";";

                // Environment variables
                output += "$env:HOME=\"" + Program.HomePath + "\";"; // Fix for missing SSH Key

                output += "$env:BUNDLE_GEMFILE=\"" + sProjectDir + "\\Gemfile\";";
                output += "$env:BUNDLE_BIN_PATH=\"\\\";";
                output += "$env:RUBYOPT=\"\\\";";

                output += "$env:CI_SERVER=\"yes\";";
                output += "$env:CI_SERVER_NAME=\"GitLab CI\";";
                output += "$env:CI_SERVER_VERSION=\"null\";"; // GitlabCI Version
                output += "$env:CI_SERVER_REVISION=\"null\";"; // GitlabCI Revision

                output += "$env:CI_BUILD_REF=\"" + buildInfo.sha + "\";";
                output += "$env:CI_BUILD_REF_NAME=\"" + buildInfo.ref_name + "\";";
                output += "$env:CI_BUILD_ID=\"" + buildInfo.id.ToString() + "\";";

                output = output.Replace(";", ";" + System.Environment.NewLine);

                output += System.Environment.NewLine + System.Environment.NewLine + script;

                fileWriter.Write(output);
            }

            return file;
        }

        void outputCollection_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<PSObject> myp = (PSDataCollection<PSObject>)sender;

            Collection<PSObject> results = myp.ReadAll();
            foreach (PSObject result in results)
            {
                outputList.Enqueue(result.ToString());
            }

        }
        /// <summary>
        /// STDOUT/STDERR Handler
        /// </summary>
        /// <param name="sendingProcess">Source process</param>
        /// <param name="outLine">Output Line</param>
        private void outputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if (!String.IsNullOrEmpty(outLine.Data))
            {
                outputList.Enqueue(outLine.Data);
            }
        }

        /// <summary>
        /// Get the Checkout CMD
        /// </summary>
        /// <returns>Checkout CMD</returns>
        private string checkoutCmd()
        {
            String sCmd = "";

            // SSH Key Path Fix

            // Change to drive
            sCmd = sProjectDir.Substring(0, 1) + ":";
            // Change to directory
            sCmd += " && cd " + sProjectDir;
            // Git Reset
            sCmd += " && git reset --hard";
            // Git Checkout
            sCmd += " && git checkout " + buildInfo.sha;

            return sCmd;
        }

        /// <summary>
        /// Get the Clone CMD
        /// </summary>
        /// <returns>Clone CMD</returns>
        private string cloneCmd()
        {
            String sCmd = "";

            // Change to drive
            sCmd = sProjectDir.Substring(0, 1) + ":";
            // Change to directory
            sCmd += " && cd " + sProjectsDir;
            // Git Clone
            sCmd += " && git clone " + buildInfo.repo_url + " project-" + buildInfo.project_id;
            // Change to directory
            sCmd += " && cd " + sProjectDir;
            // Git Checkout
            sCmd += " && git checkout " + buildInfo.sha;

            return sCmd;
        }

        /// <summary>
        /// Get the Fetch CMD
        /// </summary>
        /// <returns>Fetch CMD</returns>
        private string fetchCmd()
        {
            String sCmd = "";

            // Change to drive
            sCmd = sProjectDir.Substring(0, 1) + ":";
            // Change to directory
            sCmd += " && cd " + sProjectDir;
            // Git Reset
            sCmd += " && git reset --hard";
            // Git Clean
            sCmd += " && git clean -f";
            // Git fetch
            sCmd += " && git fetch";

            return sCmd;
        }

        /// <summary>
        /// Delete non empty directory tree
        /// </summary>
        private void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                try
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Delete(file);
                }
                catch (PathTooLongException)
                {
                    LongPathFile.Delete(file);
                }
            }

            foreach (string dir in dirs)
            {
                // Only recurse into "normal" directories
                if ((File.GetAttributes(dir) & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint)
                    try
                    {
                        Directory.Delete(dir, false);
                    }
                    catch (PathTooLongException)
                    {
                        LongPathDirectory.Delete(dir);
                    }
                else
                    DeleteDirectory(dir);
            }

            try
            {
                Directory.Delete(target_dir, false);
            }
            catch (PathTooLongException)
            {
                LongPathDirectory.Delete(target_dir);
            }
        }
    }
}
