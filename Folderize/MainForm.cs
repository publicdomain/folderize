// <copyright file="MainForm.cs" company="PublicDomain.is">
//     CC0 1.0 Universal (CC0 1.0) - Public Domain Dedication
//     https://creativecommons.org/publicdomain/zero/1.0/legalcode
// </copyright>
namespace Folderize
{
    // Directives
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Windows.Forms;
    using Microsoft.Win32;

    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form
    {
        /// <summary>
        /// Gets or sets the associated icon.
        /// </summary>
        /// <value>The associated icon.</value>
        private Icon associatedIcon = null;

        /// <summary>
        /// The folderize key list.
        /// </summary>
        private List<string> folderizeKeyList = new List<string> { @"Software\Classes\*\shell\Folderize", @"Software\Classes\directory\shell\Folderize" };

        /// <summary>
        /// Initializes a new instance of the <see cref="T:Folderize.MainForm"/> class.
        /// </summary>
        public MainForm()
        {
            // The InitializeComponent() call is required for Windows Forms designer support.
            this.InitializeComponent();

            /* Set icons */

            // Set associated icon from exe file
            this.associatedIcon = Icon.ExtractAssociatedIcon(typeof(MainForm).GetTypeInfo().Assembly.Location);

            // Set public domain weekly tool strip menu item image
            this.freeReleasesPublicDomainisToolStripMenuItem.Image = this.associatedIcon.ToBitmap();

            // Update GUI
            this.UpdateByRegistryKey();
        }

        /// <summary>
        /// Ons the add button click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnAddButtonClick(object sender, EventArgs e)
        {
            try
            {
                // Iterate folderize registry keys 
                foreach (var folderizeKey in this.folderizeKeyList)
                {
                    // Add folderize command to registry
                    RegistryKey registryKey;
                    registryKey = Registry.CurrentUser.CreateSubKey(folderizeKey);
                    registryKey.SetValue("icon", Application.ExecutablePath);
                    registryKey.SetValue("position", "Top");
                    registryKey = Registry.CurrentUser.CreateSubKey($"{folderizeKey}\\command");
                    registryKey.SetValue(string.Empty, $"{Path.Combine(Application.StartupPath, Application.ExecutablePath)} \"%1\"");
                    registryKey.Close();
                }

                // Update the program by registry key
                this.UpdateByRegistryKey();

                // Notify user
                MessageBox.Show($"Folderize context menu added!{Environment.NewLine}{Environment.NewLine}Right-click in Windows Explorer to use it.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Notify user
                MessageBox.Show($"Error when adding folderize context menu to registry.{Environment.NewLine}{Environment.NewLine}Message:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handles the remove button click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnRemoveButtonClick(object sender, EventArgs e)
        {
            try
            {
                // Iterate folderize registry keys 
                foreach (var folderizeKey in this.folderizeKeyList)
                {
                    // Remove folderize command to registry
                    Registry.CurrentUser.DeleteSubKeyTree(folderizeKey);
                }

                // Update the program by registry key
                this.UpdateByRegistryKey();

                // Notify user
                MessageBox.Show("Folderize context menu removed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Notify user
                MessageBox.Show($"Error when removing folderize command from registry.{Environment.NewLine}{Environment.NewLine}Message:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Ons the free releases public domainis tool strip menu item click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnFreeReleasesPublicDomainisToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Open our website
            Process.Start("https://publicdomain.is");
        }

        /// <summary>
        /// Ons the original thread donation codercom tool strip menu item click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnOriginalThreadDonationCodercomToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Open original thread
            Process.Start("https://www.donationcoder.com/forum/index.php?topic=19770.0");
        }

        /// <summary>
        /// Ons the source code githubcom tool strip menu item click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnSourceCodeGithubcomToolStripMenuItemClick(object sender, EventArgs e)
        {
            // TODO Add code
        }

        /// <summary>
        /// Ons the about tool strip menu item click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnAboutToolStripMenuItemClick(object sender, EventArgs e)
        {
            // TODO Add code
        }

        /// <summary>
        /// Updates the program by registry key.
        /// </summary>
        private void UpdateByRegistryKey()
        {
            // Try to set folderize key
            using (var folderizeKey = Registry.CurrentUser.OpenSubKey(this.folderizeKeyList[1]))
            {
                // Check for no returned registry key
                if (folderizeKey == null)
                {
                    // Disable remove button
                    this.removeButton.Enabled = false;

                    // Enable add button
                    this.addButton.Enabled = true;

                    // Update status text
                    this.activityToolStripStatusLabel.Text = "Inactive";
                }
                else
                {
                    // Disable add button
                    this.addButton.Enabled = false;

                    // Enable remove button
                    this.removeButton.Enabled = true;

                    // Update status text
                    this.activityToolStripStatusLabel.Text = "Active";
                }
            }
        }

        /// <summary>
        /// Ons the exit tool strip menu item click.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">E.</param>
        private void OnExitToolStripMenuItemClick(object sender, EventArgs e)
        {
            // TODO Add code
        }
    }
}
