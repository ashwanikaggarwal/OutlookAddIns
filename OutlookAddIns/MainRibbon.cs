using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIns.Forms;
using OutlookAddIns.Classes;
using System.Windows.Forms;

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
            Dictionary<string, string> contactDetails = new Dictionary<string, string>();
            bool foundEmail = false;
            int position;
            string[] bodyLines;

            //we iterate through the selection, though this should be just one at a time
            foreach (object item in selection)
            {
                if (item is Outlook.MailItem)
                {
                    //reset variables
                    contactDetails.Clear();
                    foundEmail = false;

                    content = (Outlook.MailItem)item;
                    bodyLines = content.Body.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    //we iterate through the body, split in lines
                    foreach (string line in bodyLines)
                    {
                        //then we iterate through the fields we want, and extract the information into our Dictionary contactDetails
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

                    //put the details into a class
                    ContactDetails allDetails = new ContactDetails();
                    allDetails.SetFromDictionary(contactDetails);

                    //once we found all the details we start the tests
                    KeyDetails details = new KeyDetails();
                    details.Company = contactDetails["Company on Badge"];
                    details.FirstName = contactDetails["First Name"];
                    details.LastName = contactDetails["Last Name"];
                    details.Email = contactDetails["Email"];

                    DatabaseService db = new DatabaseService();
                    int contactID;
                    //contactID = db.FindPerfectRegistration(details);
                    contactID = db.FindContactReference(details);

                    switch (contactID)
                    {
                        case 0:
                            content.Categories = content.Categories + ";Yellow Category; Found Something";
                            break;

                        case 201:
                            if (db.CreateBrandNew(allDetails))
                            {
                                content.Categories = content.Categories + ";Green Category; Created New!";
                            }
                            else
                            {
                                content.Categories = content.Categories + ";Red Category; Not Created";
                            }

                            break;

                        default:
                            if (db.AddRegistration(contactID) && db.UpdateContactInfo(contactID, contactDetails))
                            {
                                content.Categories = content.Categories + "; Registered";
                            }
                            else
                            {
                                content.Categories = content.Categories + ";Red Category; Something Went Wrong";
                            }
                            break;
                    }
                    content.MarkAsTask(Outlook.OlMarkInterval.olMarkTomorrow);
                    content.Save();



                }
            }


        }

    }
}
