using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIns.Forms;

namespace OutlookAddIns
{
    public partial class MainRibbon
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSetPath_Click(object sender, RibbonControlEventArgs e)
        {
            frmSetPath form = new frmSetPath();
            form.Show();
        }

        private void btnRegister_Click(object sender, RibbonControlEventArgs e)
        {
            //grab emails first
            Outlook.Selection selection = Globals.ThisAddIn.app.ActiveExplorer().Selection;
            Outlook.MailItem content;

            string regFields = Properties.Settings.Default.RegistrationFields;
            string[] fields = regFields.Split(';');
            string value = "Not Found";
            Dictionary<string, object> contactDetails = new Dictionary<string, object>();
            bool foundEmail = false;
            int position;
            string[] bodyLines;

            //we iterate through the selection, though this should be just one at a time
            foreach (object item in selection)
            {
                if (item is Outlook.MailItem)
                {
                    content = (Outlook.MailItem)item;
                    bodyLines = content.Body.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    //we iterate through the body, split in lines
                    foreach (string line in bodyLines)
                    {
                        for(int i = 0; i < fields.Length; i++)
                        {
                            string field = fields[i];
                            position = line.IndexOf(field);

                            if(position != -1)
                            {
                                if(field == "Email" && !foundEmail)
                                {
                                    int dumbChar = line.IndexOf("<");
                                    if(dumbChar != -1)
                                    {
                                        value = line.Substring(position + field.Length, dumbChar - (position + field.Length) - 1);
                                        contactDetails.Add(field, value);
                                        contactDetails[field] = value.Trim();
                                    } else {
                                        value = line.Substring(position + field.Length);
                                        contactDetails.Add(field, value);
                                        contactDetails[field] = value.Trim();
                                    }
                                    foundEmail = true;
                                }
                                else if(field != "Email")
                                {
                                    value = line.Substring(position + field.Length);
                                    value = value.Trim();
                                    contactDetails.Add(field, value);
                                }
                                
                            }

                        }
                        if (contactDetails.Count == fields.Length) break;

                    }

                    
                }
            }

            //open form with email found
            frmRegisterRequest form = new frmRegisterRequest(contactDetails);
            form.Show();
        }

    }
}
