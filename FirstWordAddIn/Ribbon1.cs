using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;


namespace FirstWordAddIn
{
    public partial class Ribbon1
    {
        private bool bKeepPrevText = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = "QuickExport.pdf";

            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                Path.Combine(desktopFolder, fileName),
                word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: true);
        }

        private void buttonXps_Click(object sender, RibbonControlEventArgs e)
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = "QuickExport.xps";

            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                Path.Combine(desktopFolder, fileName),
                word.WdExportFormat.wdExportFormatXPS,
                OpenAfterExport: true);
        }

        //******************************    xml BUTTONS   ************************************************//

        /****************************************************************************************
        *        AUTO ASSIGN  
        *****************************************************************************************/
        private void toggleButtonXMLDesc_Click(object sender, RibbonControlEventArgs e)
        {

            // if active -> embeed the xml name automatically
            Document myDoc = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);



            //if (this.toggleButtonXMLDesc.Checked == true )
            //Globals.ThisAddIn.Application.XMLSelectionChange += Application_XMLSelectionChange;
            //else Globals.ThisAddIn.Application.XMLSelectionChange -= Application_XMLSelectionChange;

            //Globals.ThisAddIn.Application.
            // myDoc.ContentControlAfterAdd += MyDoc_ContentControlAfterAdd;
            // += new contentcontrolafteradd
            //myDoc.ContentControlAfterAdd += new ContentControlAddedEventHandler(InsertXMLProperties);

        }


        private void MyDoc_ContentControlAfterAdd(word.ContentControl NewContentControl, bool InUndoRedo)
        {
            //xPath is empty !!!!!!!!
            String xPath = this.getXMLDesc(NewContentControl.XMLMapping.PrefixMappings, NewContentControl.XMLMapping.XPath);
            NewContentControl.Tag = xPath;
            NewContentControl.Title = xPath;
            //throw new NotImplementedException();
        }


        /****************************************************************************************
        *        ASSIGN ALL  
        *****************************************************************************************/
        private void button1_XMLDesc_All(object sender, RibbonControlEventArgs e)
        {
            Document myDoc = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            //Microsoft.Office.Interop.Word.ContentControl cc = this.Application.ActiveDocument.ContentControls[1];
            /*Microsoft.Office.Interop.Word.ContentControl cc = myDoc.ContentControls[1];
                String xPath = cc.XMLMapping.XPath;
                cc.Tag = xPath.Substring(27);
                cc.Title = xPath.Substring(27);
                cc.Range.Text = "I can run"; */

            // loop over all the xml parts and assign the name. 
            //ContentControls ContentControls { get; }
            //Microsoft.Office.Interop.Word.ContentControls ContentControls { get; }

            foreach (word.Section wordSection in myDoc.Sections) 
            {
                var footer = wordSection.Footers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var header = wordSection.Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                this.loopContentControls(footer.Range.ContentControls, this.bKeepPrevText);
                this.loopContentControls(header.Range.ContentControls, this.bKeepPrevText);

                /*word.Range footerRange = footer.Range;
                footerRange.Font.ColorIndex = word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 20;
                footerRange.Text = "Confidential";*/

            }

            foreach (Microsoft.Office.Interop.Word.ContentControl cc in myDoc.ContentControls)
            {
                //String xPath = cc.XMLMapping.XPath;
                /*try
                {
                    cc.Tag = xPath.Substring(27);
                    cc.Title = xPath.Substring(27);
                }
                catch (System.IndexOutOfRangeException ex)
                {
                    //System.Console.WriteLine(e.Message);
                    // Set IndexOutOfRangeException to the new exception's InnerException.
                    //throw new System.ArgumentOutOfRangeException("index parameter is out of range.", e);
                }
                catch (System.ArgumentOutOfRangeException ex) {
                    cc.Tag = xPath;
                    cc.Title = xPath;
                }*/

                String xPath = this.getXMLDesc(cc.XMLMapping.PrefixMappings, cc.XMLMapping.XPath);

                if (this.bKeepPrevText == true) {
                    if (cc.Tag == "" || cc.Tag == null )
                        cc.Tag = xPath;
            
                    if ( cc.Title == "" ||  cc.Title == null)
                        cc.Title = xPath;
                    
                } else   {
                    cc.Tag = xPath;
                    cc.Title = xPath;
                    //cc.Range.Text = "I can run";
                }

            }
        }

        /****************************************************************************************
        *        SELECT TO INSERT 
        *****************************************************************************************/
        private void button_InsertXMLD_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Selection wordSelection = Globals.ThisAddIn.Application.Selection;
            
            foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordSelection.ContentControls)
            {
                String xPath = this.getXMLDesc(cc.XMLMapping.PrefixMappings, cc.XMLMapping.XPath);
                cc.Tag = xPath;
                cc.Title = xPath;
                //cc.Range.Text = "I can run";
            }
        }

        /****************************************************************************************
        *        TOGGLE KEEP TEXT 
        *****************************************************************************************/
        private void toggleButton_keepText_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton target = sender as RibbonToggleButton;
            this.bKeepPrevText = target.Checked;
     
        }

        /****************************************************************************************
        *        CLEAN UP ALL  
        *****************************************************************************************/
        private void button_cleanup_xmlD_Click(object sender, RibbonControlEventArgs e)
        {
            Document myDoc = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            foreach (Microsoft.Office.Interop.Word.ContentControl cc in myDoc.ContentControls)
            {
                cc.Tag = "";
                cc.Title = "";
            }
            // Document myHeader = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application); 
        }

        //******************************    FUNCTIONS   ************************************************// 

        /****************************************************************************************
        *        GET DESCRIPTION 
        *****************************************************************************************/
        private string getXMLDesc(string sNamespace, string sPath)
        {
            string xmlDesc, slNamespace, slPath;
            slNamespace = sNamespace;
            slPath = sPath;

            if (slNamespace != "" && slPath != null)
            {
                slNamespace = slNamespace.Replace("/ns0:", " ");
                slNamespace = slNamespace.Replace("xmlns:ns0=", " ");
                slNamespace = slNamespace.Replace(":root[1]", " ");
            }
            else return xmlDesc = "";

            if (slPath != "" && slPath != null)
            {
                slPath = slPath.Replace("/ns0:", "");
                slPath = slPath.Replace("/ns1:", "");
                slPath = slPath.Replace("root[1]", "");
            }
            else return xmlDesc = "";

            // if (slNamespace.Length > 20)           xmlDesc = slPath;
            // else  xmlDesc =  slPath + slNamespace; 

            xmlDesc = slPath + slNamespace;
            return xmlDesc;

        }

        /****************************************************************************************
        *        LOOP ALL CONTENT CONTROLS   
        *****************************************************************************************/
        private void loopContentControls(word.ContentControls aContentControls, bool sKeepPrevText)
        {

            foreach (Microsoft.Office.Interop.Word.ContentControl cc in aContentControls)
            {
                String xPath = this.getXMLDesc(cc.XMLMapping.PrefixMappings, cc.XMLMapping.XPath);

                if (sKeepPrevText == true)
                {
                    if (cc.Tag == "" || cc.Tag == null) cc.Tag = xPath;
                    if (cc.Title == "" || cc.Title == null) cc.Title = xPath;
                }
                else
                {
                    cc.Tag = xPath;
                    cc.Title = xPath;
                }

            }
        }


    }
}
           