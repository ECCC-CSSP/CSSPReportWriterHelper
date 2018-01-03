using CSSPModelsDLL.Models;
using CSSPReportWriterHelperDLL.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CSSPEnumsDLL.Enums;
using System.Globalization;
using System.IO;
using System.Reflection;
using static System.Windows.Forms.ListBox;
using System.Security.Principal;

namespace CSSPReportWriterHelper
{
    public partial class CSSPReportWriterHelper : Form
    {
        #region Variables
        public string StartWebAddressCSSP = "http://wmon01dtchlebl2/csspwebtools/";
        //public string StartWebAddressCSSP = "http://localhost:11562/";
        public bool WebIsVisible = false;
        public string CSSPReportTemplatesPath = "";
        string ChangeText = "Changed\n\n";
        #endregion Variables

        #region Properties
        public ReportBaseService reportBaseService { get; set; }
        IPrincipal user { get; set; }
        #endregion Properties

        #region Constructors
        public CSSPReportWriterHelper()
        {
            InitializeComponent();

            //user = new GenericPrincipal(new GenericIdentity("charles.leblanc2@canada.ca", "Forms"), null);
            //reportBaseService = new ReportBaseService(LanguageEnum.en, treeViewCSSP, user);
            reportBaseService = new ReportBaseService(LanguageEnum.en, treeViewCSSP);
            reportBaseService.ReportFileType = ReportFileTypeEnum.CSV;
            treeViewCSSP.ExpandAll();
            Setup();
        }
        #endregion Constructors

        #region Events
        #region Events Buttons
        private void butClearAllChecks_Click(object sender, EventArgs e)
        {
            ClearAllChecks((ReportTreeNode)treeViewCSSP.Nodes[0]);
        }
        private void butClearAllExcelProcess_Click(object sender, EventArgs e)
        {
            ClearAllExcelProcess();
        }
        private void butClearAllWordProcess_Click(object sender, EventArgs e)
        {
            ClearAllWordProcess();
        }
        private void butCreateNew_Click(object sender, EventArgs e)
        {
            panelCreateNewFile.Visible = true;
            textBoxNewFileName.Focus();
        }
        private void butCreateNewFile_Click(object sender, EventArgs e)
        {
            panelCreateNewFile.Visible = false;
            CreateNewFile();
        }
        private void butCreateNewFileCancel_Click(object sender, EventArgs e)
        {
            panelCreateNewFile.Visible = false;
        }
        private void butGenerateDBCode_Click(object sender, EventArgs e)
        {
            string retStr = GenerateDBCode();
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                richTextBoxResults.AppendText("Error while generating DB Code:" + retStr + "\r\n");
            }
            else
            {
                richTextBoxResults.AppendText("DB code files generated OK.\r\n");
            }
        }
        private void butGenerateModels_Click(object sender, EventArgs e)
        {
            GenerateModels();
        }
        //private void butGenerateDBGetAndReplace_Click(object sender, EventArgs e)
        //{
        //    GeneratedReportBaseService();
        //}
        private void butGetID_Click(object sender, EventArgs e)
        {
            string url = ((HtmlDocument)webBrowserCSSPWebTools.Document).Window.Url.ToString();
            GetID_TVText(url);
        }
        private void butOpenDoc_Click(object sender, EventArgs e)
        {
            OpenFile();
        }
        private void butRefresh_Click(object sender, EventArgs e)
        {
            RefreshTemplateDocuments();
        }
        private void butRemoveFile_Click(object sender, EventArgs e)
        {
            RemoveFile();
        }
        private void butShowCode_Click(object sender, EventArgs e)
        {
            treeViewCSSPAfterSelect();
        }
        private void butShowResults_Click(object sender, EventArgs e)
        {
            //baseService.ShowResults();
        }
        private void butSortingDown_Click(object sender, EventArgs e)
        {
            SortingDown();
        }
        private void butSortingUp_Click(object sender, EventArgs e)
        {
            SortingUp();
        }
        private void butProduceTestDocument_Click(object sender, EventArgs e)
        {
            string butText = butProduceTestDocument.Text;
            richTextBoxResults.Text = "";
            butProduceTestDocument.Enabled = false;
            butProduceTestDocument.Text = "Working...";
            butProduceTestDocument.Refresh();
            Application.DoEvents();
            panelFileTools.Visible = false;
            TestProduceDocument();
            panelFileTools.Visible = true;
            butProduceTestDocument.Text = butText;
            butProduceTestDocument.Enabled = true;
            butProduceTestDocument.Refresh();
            Application.DoEvents();

        }
        private void butShowExpectedResult_Click(object sender, EventArgs e)
        {
            string butText = butShowExpectedResult.Text;
            richTextBoxResults.Text = "";
            butShowExpectedResult.Enabled = false;
            butShowExpectedResult.Text = "Working...";
            butShowExpectedResult.Refresh();
            Application.DoEvents();
            panelFileTools.Visible = false;
            TestBottomRightText();
            panelFileTools.Visible = true;
            butShowExpectedResult.Text = butText;
            butShowExpectedResult.Enabled = true;
            butShowExpectedResult.Refresh();
            Application.DoEvents();

        }
        private void butTestSelectedTemplate_Click(object sender, EventArgs e)
        {
            string butText = butTestSelectedTemplate.Text;
            richTextBoxResults.Text = "";
            butTestSelectedTemplate.Enabled = false;
            butTestSelectedTemplate.Text = "Working...";
            butTestSelectedTemplate.Refresh();
            Application.DoEvents();
            panelFileTools.Visible = false;
            TestSelectedTemplate();
            panelFileTools.Visible = true;
            butTestSelectedTemplate.Text = butText;
            butTestSelectedTemplate.Enabled = true;
            butTestSelectedTemplate.Refresh();
            Application.DoEvents();
        }
        private void butTreeViewExpandAll_Click(object sender, EventArgs e)
        {
            treeViewCSSP.ExpandAll();
        }
        private void butTreeViewCloseFields_Click(object sender, EventArgs e)
        {
            CloseAllTreeViewNodesOfTypeFields((ReportTreeNode)treeViewCSSP.Nodes[0]);
        }
        private void butWeb_Click(object sender, EventArgs e)
        {
            WebView();
        }
        #endregion Events Buttons
        #region Events CheckBoxes
        private void checkBoxFrancais_CheckedChanged(object sender, EventArgs e)
        {
            reportBaseService.LanguageRequest = (checkBoxFrancais.Checked ? LanguageEnum.fr : LanguageEnum.en);
            treeViewCSSPAfterSelect();
        }
        #endregion Events CheckBoxes
        #region Events ComboBoxes
        private void comboBoxTemplateDocuments_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxTemplateDocumentsSelectedValueChanged();
        }
        #endregion Events ComboBoxes
        #region Events Sorting
        private void comboBoxDBSorting_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBSortingSelectedValueChange();
        }
        #endregion Events Sorting
        #region Events Formating
        private void comboBoxDBFormatingNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFormatingNumberSelectedValueChange();
        }
        private void comboBoxDBFormatingDate_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFormatingDateSelectedValueChange();
        }
        private void comboBoxReportFormatingNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportFormatingNumberSelectedValueChange();
        }
        private void comboBoxReportFormatingDate_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportFormatingDateSelectedValueChange();
        }
        #endregion Events Formating
        #region Events RadioButtons
        private void radioButtonCSV_CheckedChanged(object sender, EventArgs e)
        {
            reportBaseService.ReportFileType = ReportFileTypeEnum.CSV;
            if (radioButtonCSV.Checked)
            {
                richTextBoxExample.Text = "";
                lblExt.Text = ".csv";
                lblTestTemplateStartFileName.Text = CSSPReportTemplatesPath + "Template_Test.csv";
                panelReportConditionBoolean.Visible = false;
                panelReportConditionDate.Visible = false;
                panelReportConditionEnum.Visible = false;
                panelReportConditionNumber.Visible = false;
                panelReportConditionText.Visible = false;
            }
            RefreshTemplateDocuments();
            butClearAllExcelProcess.Visible = false;
            butClearAllWordProcess.Visible = false;
            treeViewCSSPAfterSelect();
        }
        private void radioButtonExcel_CheckedChanged(object sender, EventArgs e)
        {
            reportBaseService.ReportFileType = ReportFileTypeEnum.Excel;
            if (radioButtonExcel.Checked)
            {
                richTextBoxExample.Text = "";
                lblExt.Text = ".xlsx";
                lblTestTemplateStartFileName.Text = CSSPReportTemplatesPath + "Template_Test.xlsx";
                panelReportConditionBoolean.Visible = false;
                panelReportConditionDate.Visible = false;
                panelReportConditionEnum.Visible = false;
                panelReportConditionNumber.Visible = false;
                panelReportConditionText.Visible = false;
            }
            RefreshTemplateDocuments();
            butClearAllExcelProcess.Visible = true;
            butClearAllWordProcess.Visible = false;
            treeViewCSSPAfterSelect();
        }
        private void radioButtonKML_CheckedChanged(object sender, EventArgs e)
        {
            richTextBoxExample.Text = "";
            reportBaseService.ReportFileType = ReportFileTypeEnum.KML;
            if (radioButtonKML.Checked)
            {
                lblExt.Text = ".kml";
                lblTestTemplateStartFileName.Text = CSSPReportTemplatesPath + "Template_Test.kml";
                panelReportConditionBoolean.Visible = false;
                panelReportConditionDate.Visible = false;
                panelReportConditionEnum.Visible = false;
                panelReportConditionNumber.Visible = false;
                panelReportConditionText.Visible = false;
            }
            RefreshTemplateDocuments();
            butClearAllExcelProcess.Visible = false;
            butClearAllWordProcess.Visible = false;
            treeViewCSSPAfterSelect();

        }
        private void radioButtonWord_CheckedChanged(object sender, EventArgs e)
        {
            richTextBoxExample.Text = "";
            reportBaseService.ReportFileType = ReportFileTypeEnum.Word;
            if (radioButtonWord.Checked)
            {
                lblExt.Text = ".docx";
                lblTestTemplateStartFileName.Text = CSSPReportTemplatesPath + "Template_Test.docx";
                panelReportConditionBoolean.Visible = true;
                panelReportConditionDate.Visible = true;
                panelReportConditionEnum.Visible = true;
                panelReportConditionNumber.Visible = true;
                panelReportConditionText.Visible = true;
            }
            RefreshTemplateDocuments();
            butClearAllExcelProcess.Visible = false;
            butClearAllWordProcess.Visible = true;
            treeViewCSSPAfterSelect();
        }
        private void radioButtonShowTemplateFiles_CheckedChanged(object sender, EventArgs e)
        {
            RefreshTemplateDocuments();
        }
        private void radioButtonShowResultFiles_CheckedChanged(object sender, EventArgs e)
        {
            RefreshTemplateDocuments();
        }
        #endregion Events RadioButtons
        #region Events DB Filtering Date
        private void comboBoxDBFilteringDate_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFilteringDateSelectedValueChanged();
        }

        private void comboBoxDBFilteringYear_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionYear = ((ReportItemModel)comboBoxDBFilteringYear.SelectedItem).ID;
        }
        private void comboBoxDBFilteringMonth_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMonth = ((ReportItemModel)comboBoxDBFilteringMonth.SelectedItem).ID;
        }
        private void comboBoxDBFilteringDay_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionDay = ((ReportItemModel)comboBoxDBFilteringDay.SelectedItem).ID;
        }
        private void comboBoxDBFilteringHour_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionHour = ((ReportItemModel)comboBoxDBFilteringHour.SelectedItem).ID;
        }
        private void comboBoxDBFilteringMinute_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMinute = ((ReportItemModel)comboBoxDBFilteringMinute.SelectedItem).ID;
        }

        #endregion Events DB Filtering Date
        #region "Events DB Filtering Enum"
        private void listBoxDBFilteringEnum_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxDBFilteringEnumSelectedIndexChanged();
        }
        #endregion "Events DB Filtering Enum"
        #region Events DB Filtering Number
        private void comboBoxDBFilteringNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFilteringNumberSelectedValueChanged();
        }
        private void textBoxDBFilteringNumber_Leave(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringNumberFieldList.Count == 0)
                reportTreeNode.dbFilteringNumberFieldList.Add(new ReportConditionNumberField());

            textBoxDBFilteringNumber.BackColor = Color.White;
            if (reportTreeNode.dbFilteringNumberFieldList[0] != null && reportTreeNode.dbFilteringNumberFieldList[0].ReportCondition != ReportConditionEnum.Error)
            {
                lblStatusValue.Text = "";
                if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.NumberWhole)
                {
                    int TheNumber;
                    if (int.TryParse(textBoxDBFilteringNumber.Text, out TheNumber))
                    {
                        reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition = TheNumber;
                    }
                    else
                    {
                        lblStatusValue.Text = "Please enter a valid number.";
                        textBoxDBFilteringNumber.BackColor = Color.Red;
                        reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition = null;
                    }
                }
                else
                {
                    float TheNumber;
                    if (float.TryParse(textBoxDBFilteringNumber.Text, out TheNumber))
                    {
                        reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition = TheNumber;
                    }
                    else
                    {
                        lblStatusValue.Text = "Please enter a valid number.";
                        textBoxDBFilteringNumber.BackColor = Color.Red;
                        reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition = null;
                    }
                }
            }
            else
            {
                lblStatusValue.Text = "Error: A number type should be selected in the Tree View.";
            }
        }
        #endregion Events DB Filtering Number
        #region Events DB Filtering Text
        private void comboBoxDBFilteringText_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFilteringTextSelectedValueChanged();
        }
        private void textBoxDBFilteringText_Leave(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            textBoxDBFilteringText.BackColor = Color.White;
            if (reportTreeNode.dbFilteringTextFieldList.Count > 0 && reportTreeNode.dbFilteringTextFieldList[0].ReportCondition != ReportConditionEnum.Error)
            {
                lblStatusValue.Text = "";
                if (string.IsNullOrWhiteSpace(textBoxDBFilteringText.Text.Trim()))
                {
                    lblStatusValue.Text = "Please enter a valid text.";
                    textBoxDBFilteringText.BackColor = Color.Red;
                    reportTreeNode.dbFilteringTextFieldList[0].TextCondition = null;
                }
                else
                {
                    reportTreeNode.dbFilteringTextFieldList[0].TextCondition = textBoxDBFilteringText.Text.Trim();
                }
            }
            else
            {
                lblStatusValue.Text = "Error: A text type should be selected in the Tree View.";
            }
        }
        #endregion Events DB Filtering Text
        #region Events DB Filtering Boolean
        private void comboBoxDBFilteringTrueFalse_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxDBFilteringTrueFalseSelectedValueChanged();
        }
        #endregion Events DB Filtering Boolean
        #region Events Report Condition Date
        private void comboBoxReportConditionDate_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportConditionDateSelectedValueChanged();
        }

        private void comboBoxReportConditionYear_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionYear = ((ReportItemModel)comboBoxReportConditionYear.SelectedItem).ID;
        }
        private void comboBoxReportConditionMonth_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMonth = ((ReportItemModel)comboBoxReportConditionMonth.SelectedItem).ID;
        }
        private void comboBoxReportConditionDay_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionDay = ((ReportItemModel)comboBoxReportConditionDay.SelectedItem).ID;
        }
        private void comboBoxReportConditionHour_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionHour = ((ReportItemModel)comboBoxReportConditionHour.SelectedItem).ID;
        }
        private void comboBoxReportConditionMinute_SelectedValueChanged(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMinute = ((ReportItemModel)comboBoxReportConditionMinute.SelectedItem).ID;
        }

        #endregion Report Condition Date
        #region "Events Report Condition Enum"
        private void listBoxReportConditionEnum_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxReportConditionEnumSelectedIndexChanged();
        }
        #endregion "Events Report Condition Enum"
        #region Events Report Condition Number
        private void comboBoxReportConditionNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportConditionNumberSelectedValueChanged();
        }
        private void textBoxReportConditionNumber_Leave(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionNumberFieldList.Count == 0)
                reportTreeNode.reportConditionNumberFieldList.Add(new ReportConditionNumberField());

            textBoxReportConditionNumber.BackColor = Color.White;
            if (reportTreeNode.reportConditionNumberFieldList[0] != null && reportTreeNode.reportConditionNumberFieldList[0].ReportCondition != ReportConditionEnum.Error)
            {
                lblStatusValue.Text = "";
                if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.NumberWhole)
                {
                    int TheNumber;
                    if (int.TryParse(textBoxReportConditionNumber.Text, out TheNumber))
                    {
                        reportTreeNode.reportConditionNumberFieldList[0].NumberCondition = TheNumber;
                    }
                    else
                    {
                        lblStatusValue.Text = "Please enter a valid number.";
                        textBoxReportConditionNumber.BackColor = Color.Red;
                        reportTreeNode.reportConditionNumberFieldList[0].NumberCondition = null;
                    }
                }
                else
                {
                    float TheNumber;
                    if (float.TryParse(textBoxReportConditionNumber.Text, out TheNumber))
                    {
                        reportTreeNode.reportConditionNumberFieldList[0].NumberCondition = TheNumber;
                    }
                    else
                    {
                        lblStatusValue.Text = "Please enter a valid number.";
                        textBoxReportConditionNumber.BackColor = Color.Red;
                        reportTreeNode.reportConditionNumberFieldList[0].NumberCondition = null;
                    }
                }
            }
            else
            {
                lblStatusValue.Text = "Error: A number type should be selected in the Tree View.";
            }
        }
        #endregion Events Report Condition Number
        #region Events Report Condition Text
        private void comboBoxReportConditionText_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportConditionTextSelectedValueChanged();
        }
        private void textBoxReportConditionText_Leave(object sender, EventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionTextFieldList.Count == 0)
                reportTreeNode.reportConditionTextFieldList.Add(new ReportConditionTextField());

            textBoxReportConditionText.BackColor = Color.White;
            if (reportTreeNode.reportConditionTextFieldList[0] != null && reportTreeNode.reportConditionTextFieldList[0].ReportCondition != ReportConditionEnum.Error)
            {
                lblStatusValue.Text = "";
                if (string.IsNullOrWhiteSpace(textBoxReportConditionText.Text.Trim()))
                {
                    lblStatusValue.Text = "Please enter a valid text.";
                    textBoxReportConditionText.BackColor = Color.Red;
                    reportTreeNode.reportConditionTextFieldList[0].TextCondition = null;
                }
                else
                {
                    reportTreeNode.reportConditionTextFieldList[0].TextCondition = textBoxReportConditionText.Text.Trim();
                }
            }
            else
            {
                lblStatusValue.Text = "Error: A text type should be selected in the Tree View.";
            }
        }
        #endregion Events Report Condition Text
        #region Events Report Condition Boolean
        private void comboBoxReportConditionTrueFalse_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBoxReportConditionTrueFalseSelectedValueChanged();
        }
        #endregion Events Report Condition Boolean
        #region Events treeViewCSSP
        private void treeViewCSSP_AfterSelect(object sender, TreeViewEventArgs e)
        {
            treeViewCSSPAfterSelect();
        }
        private void treeViewCSSP_AfterCheck(object sender, TreeViewEventArgs e)
        {
            treeViewCSSPAfterCheck(e);
            reportBaseService.SetParentChecked(e);
            if (!richTextBoxExample.Text.StartsWith(ChangeText))
            {
                richTextBoxExample.Text = ChangeText + richTextBoxExample.Text;
            }
        }
        #endregion Events treeViewCSSP
        #endregion Events

        #region Functions public
        private void ClearAllChecks(ReportTreeNode reportTreeNode)
        {
            reportTreeNode.Checked = false;
            foreach (ReportTreeNode RTN in reportTreeNode.Nodes)
            {
                ClearAllChecks(RTN);
            }
        }
        private void ClearAllExcelProcess()
        {
            List<Process> excelProcesses = Process.GetProcessesByName("Excel").ToList();
            foreach (Process proc in excelProcesses)
            {
                if (MessageBox.Show(proc.MainWindowTitle + "\r\n\r\nThis will close Excel without saving.", "Killing Excel Process", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        proc.Kill();
                    }
                    catch { }
                }
            }
        }
        private void ClearAllWordProcess()
        {
            List<Process> wordProcesses = Process.GetProcessesByName("WinWord").ToList();
            foreach (Process proc in wordProcesses)
            {
                if (MessageBox.Show(proc.MainWindowTitle + "\r\n\r\nThis will close Word without saving.", "Killing Word Process", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        proc.Kill();
                    }
                    catch { }
                }
            }
        }
        public void CreateCSSPReportTemplatesDirectory()
        {
            CSSPReportTemplatesPath = @"C:\CSSPReportTemplates\";

            DirectoryInfo di = new DirectoryInfo(CSSPReportTemplatesPath);
            if (!di.Exists)
            {
                try
                {
                    di.Create();
                    lblCurrentFilePath.Text = CSSPReportTemplatesPath;
                }
                catch (Exception ex)
                {
                    lblStatusValue.Text = "Could not create directory[" + CSSPReportTemplatesPath + "]" + ex.Message + "Inner: " + (ex.InnerException != null ? ex.InnerException.Message : "");
                    lblCurrentFilePath.Text = lblStatusValue.Text;
                    return;
                }
            }

        }
        public void CreateNewFile()
        {
            if (string.IsNullOrWhiteSpace(textBoxNewFileName.Text))
            {
                MessageBox.Show("Please enter a file name.");
                return;
            }
            if (textBoxNewFileName.Text.Contains("."))
            {
                MessageBox.Show("Please enter a file name without the extension. No dot.");
                return;
            }
            string extension = ".err";
            if (radioButtonCSV.Checked)
            {
                extension = ".csv";
            }
            else if (radioButtonWord.Checked)
            {
                extension = ".docx";
            }
            else if (radioButtonExcel.Checked)
            {
                extension = ".xlsx";
            }
            else if (radioButtonKML.Checked)
            {
                extension = ".kml";
            }
            FileInfo fi = new FileInfo(CSSPReportTemplatesPath + "Template_" + textBoxNewFileName.Text + extension);

            if (radioButtonCSV.Checked)
            {
                FileStream fs = fi.OpenWrite();
                fs.Close();
            }
            else if (radioButtonWord.Checked)
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                doc.SaveAs2(fi.FullName);
                wordApp.Quit();
            }
            else if (radioButtonExcel.Checked)
            {
                MessageBox.Show("Excel not implemented yet.");
                //Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                //excelApp.Visible = true;
                //Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();
                //workbook.SaveAs(fi.FullName);
                //excelApp.Quit();
            }
            else if (radioButtonKML.Checked)
            {
                MessageBox.Show("KML not implemented yet.");
            }

            RefreshTemplateDocuments();
        }
        public void CloseAllTreeViewNodesOfTypeFields(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.FieldsHolder)
            {
                reportTreeNode.Collapse();
            }
            else
            {
                foreach (ReportTreeNode RTN in reportTreeNode.Nodes)
                {
                    CloseAllTreeViewNodesOfTypeFields(RTN);
                }
            }
        }
        public void comboBoxDBFilteringDateSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFilteringDate.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
                reportTreeNode.dbFilteringDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.dbFilteringDateFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.Error:
                    {
                        comboBoxDBFilteringYear.BackColor = Color.White;
                        comboBoxDBFilteringYear.SelectedIndex = 0;
                        comboBoxDBFilteringMonth.SelectedIndex = 0;
                        comboBoxDBFilteringDay.SelectedIndex = 0;
                        comboBoxDBFilteringHour.SelectedIndex = 0;
                        comboBoxDBFilteringMinute.SelectedIndex = 0;
                    }
                    break;
                case ReportConditionEnum.ReportConditionEqual:
                case ReportConditionEnum.ReportConditionBigger:
                case ReportConditionEnum.ReportConditionSmaller:
                    {
                        if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionYear != null)
                        {
                            comboBoxDBFilteringYear.SelectedItem = reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionYear;
                        }
                        if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMonth != null)
                        {
                            comboBoxDBFilteringMonth.SelectedItem = reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMonth;
                        }
                        if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionDay != null)
                        {
                            comboBoxDBFilteringDay.SelectedItem = reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionDay;
                        }
                        if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionHour != null)
                        {
                            comboBoxDBFilteringHour.SelectedItem = reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionHour;
                        }
                        if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMinute != null)
                        {
                            comboBoxDBFilteringMinute.SelectedItem = reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMinute;
                        }
                    }
                    break;
                default:
                    {
                        textBoxDBFilteringNumber.Text = "";
                    }
                    break;
            }
        }
        public void comboBoxDBFilteringNumberSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFilteringNumber.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringNumberFieldList.Count == 0)
                reportTreeNode.dbFilteringNumberFieldList.Add(new ReportConditionNumberField());

            reportTreeNode.dbFilteringNumberFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.Error:
                    {
                        textBoxDBFilteringNumber.BackColor = Color.White;
                        textBoxDBFilteringNumber.Text = "";
                    }
                    break;
                case ReportConditionEnum.ReportConditionEqual:
                case ReportConditionEnum.ReportConditionBigger:
                case ReportConditionEnum.ReportConditionSmaller:
                    {
                        if (reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition != null)
                        {
                            textBoxDBFilteringNumber.Text = reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition.ToString();
                        }
                    }
                    break;
                default:
                    {
                        textBoxDBFilteringNumber.Text = "";
                    }
                    break;
            }
        }
        public void comboBoxDBFilteringTextSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFilteringText.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringTextFieldList.Count == 0)
                reportTreeNode.dbFilteringTextFieldList.Add(new ReportConditionTextField());

            reportTreeNode.dbFilteringTextFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.Error:
                    {
                        textBoxDBFilteringText.BackColor = Color.White;
                        textBoxDBFilteringText.Text = "";
                    }
                    break;
                case ReportConditionEnum.ReportConditionContain:
                case ReportConditionEnum.ReportConditionStart:
                case ReportConditionEnum.ReportConditionEnd:
                case ReportConditionEnum.ReportConditionEqual:
                    {
                        if (reportTreeNode.dbFilteringTextFieldList[0].TextCondition != null)
                        {
                            textBoxDBFilteringText.Text = reportTreeNode.dbFilteringTextFieldList[0].TextCondition.ToString();
                        }
                    }
                    break;
                default:
                    {
                        textBoxDBFilteringText.Text = "";
                    }
                    break;
            }
        }
        public void comboBoxDBFilteringTrueFalseSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFilteringTrueFalse.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null)
                return;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFilteringTrueFalseFieldList.Count == 0)
                reportTreeNode.dbFilteringTrueFalseFieldList.Add(new ReportConditionTrueFalseField());

            reportTreeNode.dbFilteringTrueFalseFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);
        }
        public void comboBoxDBSortingSelectedValueChange()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBSorting.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null || reportTreeNode == null)
                return;

            int NextOrdinal = GetNextSortingOdinalNumber();

            switch ((ReportSortingEnum)(reportItemModel.ID))
            {
                case ReportSortingEnum.ReportSortingAscending:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.dbSortingField.ReportSorting = ReportSortingEnum.ReportSortingAscending;
                            if (reportTreeNode.dbSortingField.Ordinal == 0)
                            {
                                reportTreeNode.dbSortingField.Ordinal = NextOrdinal;
                                butSortingDown.Enabled = false;
                            }
                        }
                    }
                    break;
                case ReportSortingEnum.ReportSortingDescending:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.dbSortingField.ReportSorting = ReportSortingEnum.ReportSortingDescending;
                            if (reportTreeNode.dbSortingField.Ordinal == 0)
                            {
                                reportTreeNode.dbSortingField.Ordinal = NextOrdinal;
                                butSortingDown.Enabled = false;
                            }
                        }
                    }
                    break;
                default:
                    {
                        reportTreeNode.dbSortingField.ReportSorting = ReportSortingEnum.Error;
                        reportTreeNode.dbSortingField.Ordinal = 0;
                    }
                    break;
            }

            if (reportTreeNode.dbSortingField.Ordinal == 0)
            {
                butSortingUp.Enabled = false;
                butSortingDown.Enabled = false;
            }
            else
            {
                butSortingUp.Enabled = true;
                butSortingDown.Enabled = true;
                if (reportTreeNode.dbSortingField.Ordinal == NextOrdinal)
                {
                    butSortingDown.Enabled = false;
                }
            }
            lblSortingOrdinal.Text = reportTreeNode.dbSortingField.Ordinal.ToString();
        }
        public void comboBoxReportConditionDateSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportConditionDate.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null)
            {
                panelReportConditionDate.Visible = false;
                return;
            }

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
                reportTreeNode.reportConditionDateFieldList.Add(new ReportConditionDateField());

            reportTreeNode.reportConditionDateFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.ReportConditionEqual:
                case ReportConditionEnum.ReportConditionBigger:
                case ReportConditionEnum.ReportConditionSmaller:
                    {
                        panelReportConditionDate.Visible = true;
                    }
                    break;
                default:
                    break;
            }
        }
        public void comboBoxReportConditionNumberSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportConditionNumber.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionNumberFieldList.Count == 0)
                reportTreeNode.reportConditionNumberFieldList.Add(new ReportConditionNumberField());

            reportTreeNode.reportConditionNumberFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.Error:
                    {
                        textBoxReportConditionNumber.BackColor = Color.White;
                        textBoxReportConditionNumber.Text = "";
                    }
                    break;
                case ReportConditionEnum.ReportConditionEqual:
                case ReportConditionEnum.ReportConditionBigger:
                case ReportConditionEnum.ReportConditionSmaller:
                    {
                        if (reportTreeNode.reportConditionNumberFieldList[0].NumberCondition != null)
                        {
                            textBoxReportConditionNumber.Text = reportTreeNode.reportConditionNumberFieldList[0].NumberCondition.ToString();
                        }
                    }
                    break;
                default:
                    {
                        textBoxReportConditionNumber.Text = "";
                    }
                    break;
            }
        }
        public void comboBoxReportConditionTextSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportConditionText.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionTextFieldList.Count == 0)
                reportTreeNode.reportConditionTextFieldList.Add(new ReportConditionTextField());

            reportTreeNode.reportConditionTextFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);

            switch ((ReportConditionEnum)(reportItemModel.ID))
            {
                case ReportConditionEnum.Error:
                    {
                        textBoxReportConditionText.BackColor = Color.White;
                        textBoxReportConditionText.Text = "";
                    }
                    break;
                case ReportConditionEnum.ReportConditionContain:
                case ReportConditionEnum.ReportConditionStart:
                case ReportConditionEnum.ReportConditionEnd:
                case ReportConditionEnum.ReportConditionEqual:
                    {
                        if (reportTreeNode.reportConditionTextFieldList[0].TextCondition != null)
                        {
                            textBoxReportConditionText.Text = reportTreeNode.reportConditionTextFieldList[0].TextCondition.ToString();
                        }
                        reportTreeNode.reportConditionTextFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);
                        reportTreeNode.reportConditionTextFieldList[0].TextCondition = textBoxReportConditionText.Text;
                    }
                    break;
                default:
                    {
                        textBoxReportConditionText.Text = "";
                    }
                    break;
            }
        }
        public void comboBoxReportConditionTrueFalseSelectedValueChanged()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportConditionTrueFalse.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null)
                return;

            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportConditionTrueFalseFieldList.Count == 0)
                reportTreeNode.reportConditionTrueFalseFieldList.Add(new ReportConditionTrueFalseField());

            reportTreeNode.reportConditionTrueFalseFieldList[0].ReportCondition = (ReportConditionEnum)(reportItemModel.ID);
        }
        public void comboBoxDBFormatingDateSelectedValueChange()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFormatingDate.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null || reportTreeNode == null)
                return;

            switch ((ReportFormatingDateEnum)(reportItemModel.ID))
            {
                case ReportFormatingDateEnum.ReportFormatingDateDayOnly:
                case ReportFormatingDateEnum.ReportFormatingDateHourOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMinuteOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearOnly:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.dbFormatingField.ReportFormatingDate = (ReportFormatingDateEnum)(reportItemModel.ID);
                        }
                    }
                    break;
                default:
                    {
                        reportTreeNode.dbFormatingField.ReportFormatingDate = ReportFormatingDateEnum.Error;
                    }
                    break;
            }
        }
        public void comboBoxDBFormatingNumberSelectedValueChange()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxDBFormatingNumber.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null || reportTreeNode == null)
                return;

            switch ((ReportFormatingNumberEnum)(reportItemModel.ID))
            {
                case ReportFormatingNumberEnum.ReportFormatingNumber0Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber1Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber2Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber3Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber4Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber5Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber6Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.dbFormatingField.ReportFormatingNumber = (ReportFormatingNumberEnum)(reportItemModel.ID);
                        }
                    }
                    break;
                default:
                    {
                        reportTreeNode.dbFormatingField.ReportFormatingNumber = ReportFormatingNumberEnum.Error;
                    }
                    break;
            }
        }
        public void comboBoxReportFormatingDateSelectedValueChange()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportFormatingDate.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null || reportTreeNode == null)
                return;

            switch ((ReportFormatingDateEnum)(reportItemModel.ID))
            {
                case ReportFormatingDateEnum.ReportFormatingDateDayOnly:
                case ReportFormatingDateEnum.ReportFormatingDateHourOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMinuteOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly:
                case ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay:
                case ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute:
                case ReportFormatingDateEnum.ReportFormatingDateYearOnly:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.reportFormatingField.ReportFormatingDate = (ReportFormatingDateEnum)(reportItemModel.ID);
                        }
                    }
                    break;
                default:
                    {
                        reportTreeNode.reportFormatingField.ReportFormatingDate = ReportFormatingDateEnum.Error;
                    }
                    break;
            }
        }
        public void comboBoxReportFormatingNumberSelectedValueChange()
        {
            ReportItemModel reportItemModel = (ReportItemModel)comboBoxReportFormatingNumber.SelectedItem;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportItemModel == null || reportTreeNode == null)
                return;

            switch ((ReportFormatingNumberEnum)(reportItemModel.ID))
            {
                case ReportFormatingNumberEnum.ReportFormatingNumber0Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber1Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber2Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber3Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber4Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber5Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumber6Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal:
                case ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal:
                    {
                        if (reportTreeNode != null)
                        {
                            reportTreeNode.reportFormatingField.ReportFormatingNumber = (ReportFormatingNumberEnum)(reportItemModel.ID);
                        }
                    }
                    break;
                default:
                    {
                        reportTreeNode.reportFormatingField.ReportFormatingNumber = ReportFormatingNumberEnum.Error;
                    }
                    break;
            }
        }
        public void comboBoxTemplateDocumentsSelectedValueChanged()
        {
            FileInfo fi = (FileInfo)comboBoxTemplateDocuments.SelectedItem;
            if (fi != null)
            {
                lblCurrentFilePath.Text = fi.FullName;
            }
        }
        public string CreateTemplateAndResultDocumentShell()
        {
            FileInfo fiBottomRight = new FileInfo(lblTestTemplateStartFileName.Text);

            if (radioButtonWord.Checked)
            {
                List<Process> wordProcesses = Process.GetProcessesByName("WinWord").ToList();
                foreach (Process proc in wordProcesses)
                {
                    if (proc.MainWindowTitle == "Template_Test.docx - Microsoft Word")
                    {
                        lblStatusValue.Text = fiBottomRight.FullName + " is currently being used. Please close it or use the Clear all Word process button";
                        richTextBoxResults.Text = lblStatusValue.Text;
                        return lblStatusValue.Text;
                    }
                }
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                Microsoft.Office.Interop.Word.Range range = doc.Range();

                doc.Range().Text = "";
                doc.Range().Select();
                doc.Range().set_Style(doc.Application.ActiveDocument.Styles["No Spacing"]);
                string retStr = InsertTemplateTextInWord(doc);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    lblStatusValue.Text = retStr;
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }

                try
                {
                    doc.SaveAs2(fiBottomRight.FullName);
                    doc.Close();
                    wordApp.Quit();
                }
                catch (Exception)
                {
                    lblStatusValue.Text = fiBottomRight.FullName + " is currently being used. Please close it or use the Clear all Word process button";
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }

                fiBottomRight = new FileInfo(fiBottomRight.FullName);
                if (!fiBottomRight.Exists)
                {
                    lblStatusValue.Text = "File [" + fiBottomRight.FullName + "] could not be created.";
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }
            }
            else if (radioButtonExcel.Checked)
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbooks workbooks = excelApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Sheets Sheets = excelApp.Sheets;
                Microsoft.Office.Interop.Excel.Worksheet workSheet = Sheets[1];

                string retStr = InsertTemplateTextInExcel(workSheet);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    lblStatusValue.Text = retStr;
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }

                excelApp.DisplayAlerts = false;

                workBook.SaveAs(fiBottomRight.FullName);
                excelApp.Quit();

                List<Process> oldExcelProcesses = Process.GetProcessesByName("Excel").ToList();
                List<Process> newExcelProcesses = new List<Process>();

                newExcelProcesses = Process.GetProcessesByName("Excel").ToList();
                foreach (Process proc in newExcelProcesses)
                {
                    if (!oldExcelProcesses.Contains(proc))
                    {
                        try
                        {
                            proc.Kill();
                        }
                        catch { }
                    }
                }

                fiBottomRight = new FileInfo(fiBottomRight.FullName);
                if (!fiBottomRight.Exists)
                {
                    lblStatusValue.Text = "File [" + fiBottomRight.FullName + "] could not be created.";
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }
            }
            else if (radioButtonKML.Checked)
            {
                StreamWriter sw = fiBottomRight.CreateText();
                sw.Write(richTextBoxExample.Text.Replace("\n", "\r\n"));
                sw.Close();
            }
            else
            {
                StreamWriter sw = fiBottomRight.CreateText();
                sw.Write(richTextBoxExample.Text.Replace("\n", "\r\n"));
                sw.Close();
            }

            FileInfo fi = new FileInfo(fiBottomRight.FullName.Replace("Template_", ""));

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    lblStatusValue.Text = ex.Message + (ex.InnerException != null ? " - Inner: " + ex.InnerException.Message : "");
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return lblStatusValue.Text;
                }
            }

            try
            {
                File.Copy(fiBottomRight.FullName, fi.FullName);
            }
            catch (Exception ex)
            {
                lblStatusValue.Text = ex.Message + (ex.InnerException != null ? " - Inner: " + ex.InnerException.Message : "");
                richTextBoxResults.Text = lblStatusValue.Text;
                return lblStatusValue.Text;
            }
            fi = new FileInfo(fi.FullName);
            if (!fi.Exists)
            {
                lblStatusValue.Text = "File [" + fi.FullName + "] could not be created.";
                richTextBoxResults.Text = lblStatusValue.Text;
                return lblStatusValue.Text;
            }

            return "";
        }
        public string GenerateDBCode()
        {

            richTextBoxResults.Text = @"Generating File [C:\CSSP latest code old\CSSPWebToolsDBDLL\CSSPWebToolsDBDLL\Services\ReportServiceGenerated_____.cs]" + "\r\n";
            richTextBoxResults.AppendText(@"Generating Files:" + "\r\n");

            List<Type> typeList = new List<Type>()
            {
                typeof(ReportArea_FileModel),
                typeof(ReportAreaModel),
                typeof(ReportBox_Model_ResultModel),
                typeof(ReportBox_ModelModel),
                typeof(ReportClimate_Site_DataModel),
                typeof(ReportClimate_SiteModel),
                typeof(ReportCountry_FileModel),
                typeof(ReportCountryModel),
                typeof(ReportHydrometric_Site_DataModel),
                typeof(ReportHydrometric_Site_Rating_Curve_ValueModel),
                typeof(ReportHydrometric_Site_Rating_CurveModel),
                typeof(ReportHydrometric_SiteModel),
                typeof(ReportInfrastructure_AddressModel),
                typeof(ReportInfrastructure_FileModel),
                typeof(ReportInfrastructureModel),
                typeof(ReportMike_Boundary_ConditionModel),
                typeof(ReportMike_Scenario_FileModel),
                typeof(ReportMike_ScenarioModel),
                typeof(ReportMike_Source_Start_EndModel),
                typeof(ReportMike_SourceModel),
                typeof(ReportMPN_LookupModel),
                typeof(ReportMunicipality_Contact_AddressModel),
                typeof(ReportMunicipality_Contact_EmailModel),
                typeof(ReportMunicipality_Contact_TelModel),
                typeof(ReportMunicipality_ContactModel),
                typeof(ReportMunicipality_FileModel),
                typeof(ReportMunicipalityModel),
                typeof(ReportSampling_Plan_Lab_Sheet_DetailModel),
                typeof(ReportSampling_Plan_Lab_Sheet_Tube_And_MPN_DetailModel),
                typeof(ReportSampling_Plan_Lab_SheetModel),
                typeof(ReportSampling_Plan_Subsector_SiteModel),
                typeof(ReportSampling_Plan_SubsectorModel),
                typeof(ReportSampling_PlanModel),
                typeof(ReportMWQM_Run_FileModel),
                typeof(ReportMWQM_Run_Lab_Sheet_DetailModel),
                typeof(ReportMWQM_Run_Lab_Sheet_Tube_And_MPN_DetailModel),
                typeof(ReportMWQM_Run_Lab_SheetModel),
                typeof(ReportMWQM_Run_SampleModel),
                typeof(ReportMWQM_RunModel),
                typeof(ReportMWQM_Site_FileModel),
                typeof(ReportMWQM_Site_SampleModel),
                typeof(ReportMWQM_Site_Start_And_End_DateModel),
                typeof(ReportMWQM_SiteModel),
                typeof(ReportPol_Source_Site_AddressModel),
                typeof(ReportPol_Source_Site_FileModel),
                typeof(ReportPol_Source_Site_Obs_IssueModel),
                typeof(ReportPol_Source_Site_ObsModel),
                typeof(ReportPol_Source_SiteModel),
                typeof(ReportProvince_FileModel),
                typeof(ReportProvinceModel),
                typeof(ReportRoot_FileModel),
                typeof(ReportRootModel),
                typeof(ReportSector_FileModel),
                typeof(ReportSectorModel),
                typeof(ReportSubsector_FileModel),
                typeof(ReportSubsector_Lab_Sheet_DetailModel),
                typeof(ReportSubsector_Lab_Sheet_Tube_And_MPN_DetailModel),
                typeof(ReportSubsector_Lab_SheetModel),
                typeof(ReportSubsectorModel),
                typeof(ReportSubsector_Special_TableModel),
                typeof(ReportSubsector_Tide_Site_DataModel),
                typeof(ReportSubsector_Tide_SiteModel),
                typeof(ReportSubsector_Climate_SiteModel),
                typeof(ReportSubsector_Climate_Site_DataModel),
                typeof(ReportSubsector_Hydrometric_SiteModel),
                typeof(ReportSubsector_Hydrometric_Site_DataModel),
                typeof(ReportSubsector_Hydrometric_Site_Rating_CurveModel),
                typeof(ReportSubsector_Hydrometric_Site_Rating_Curve_ValueModel),
                typeof(ReportVisual_Plumes_Scenario_AmbientModel),
                typeof(ReportVisual_Plumes_Scenario_ResultModel),
                typeof(ReportVisual_Plumes_ScenarioModel),
             };

            foreach (Type type in typeList)
            {
                string retStr = GenerateDBCodeOfType(type);
                if (!string.IsNullOrWhiteSpace(retStr))
                    return retStr;
            }

            richTextBoxResults.AppendText("\r\nPlease Recompile CSSPWebToolsDBDLL project\r\n");

            return "";
        }
        public string GenerateDBCodeOfType(Type type)
        {
            string PartialFileName = type.Name.Substring(6);
            PartialFileName = PartialFileName.Substring(0, PartialFileName.Length - 5);

            StringBuilder sb = new StringBuilder();
            FileInfo fi = new FileInfo(@"C:\CSSP latest code old\CSSPWebToolsDBDLL\CSSPWebToolsDBDLL\Services\ReportServiceGenerated" + PartialFileName + ".cs");

            richTextBoxResults.AppendText(@"ReportServiceGenerated" + PartialFileName + ".cs\r\n");

            sb.AppendLine(@"using System.Linq;");
            sb.AppendLine(@"using System.Security.Principal;");
            sb.AppendLine(@"using System.Text;");
            sb.AppendLine(@"using System.Threading.Tasks;");
            sb.AppendLine(@"using CSSPWebToolsDBDLL.Models;");
            sb.AppendLine(@"using System.Collections.Generic;");
            sb.AppendLine(@"using System;");
            sb.AppendLine(@"using CSSPWebToolsDBDLL.Services.Resources;");
            sb.AppendLine(@"using System.Transactions;");
            sb.AppendLine(@"using System.Web.Mvc;");
            sb.AppendLine(@"using System.Threading;");
            sb.AppendLine(@"using System.Globalization;");
            sb.AppendLine(@"using CSSPModelsDLL.Models;");
            sb.AppendLine(@"using CSSPEnumsDLL.Enums;");
            sb.AppendLine(@"using System.IO;");
            sb.AppendLine(@"using System.Reflection;");
            sb.AppendLine(@"");

            sb.AppendLine(@"namespace CSSPWebToolsDBDLL.Services");
            sb.AppendLine(@"{");
            sb.AppendLine(@"    public partial class ReportService" + PartialFileName + "");
            sb.AppendLine(@"    {");
            sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, List<ReportTreeNode> reportTreeNodeList)");
            sb.AppendLine(@"        {");
            string retStr = GenerateDBCodeOfTypeSorting(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeDate(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeText(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeNumber(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeTrueFalse(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeEnum(type, sb, PartialFileName);
            sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"");
            sb.AppendLine(@"        // Date Functions");
            retStr = GenerateDBCodeOfTypeDateFunctionYEAR(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeDateFunctionMONTH(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeDateFunctionDAY(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeDateFunctionHOUR(type, sb, PartialFileName);
            retStr = GenerateDBCodeOfTypeDateFunctionMINUTE(type, sb, PartialFileName);
            sb.AppendLine(@"");
            sb.AppendLine(@"        // Text Functions");
            retStr = GenerateDBCodeOfTypeTextFunction(type, sb, PartialFileName);
            sb.AppendLine(@"");
            sb.AppendLine(@"        // Number Functions");
            retStr = GenerateDBCodeOfTypeNumberFunction(type, sb, PartialFileName);
            sb.AppendLine(@"    }");
            sb.AppendLine(@"}");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();

            return "";
        }
        private string GenerateDBCodeOfTypeTrueFalse(Type type, StringBuilder sb, string PartialFileName)
        {
            sb.AppendLine(@"            #region Filter TrueFalse");
            sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbFilteringTrueFalseFieldList.Count > 0))");
            sb.AppendLine(@"            {");
            sb.AppendLine(@"                foreach (ReportConditionTrueFalseField reportTrueFalseField in reportTreeNode.dbFilteringTrueFalseFieldList)");
            sb.AppendLine(@"                {");
            sb.AppendLine(@"                    switch (reportTreeNode.Text)");
            sb.AppendLine(@"                    {");
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.Boolean")))
            {
                sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                sb.AppendLine(@"                            if (reportTrueFalseField.ReportCondition == ReportConditionEnum.ReportConditionTrue)");
                sb.AppendLine(@"                               report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + " == true);");
                sb.AppendLine(@"                            else");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + " == false);");
                sb.AppendLine(@"                            break;");
            }
            sb.AppendLine(@"                        default:");
            sb.AppendLine(@"                            break;");
            sb.AppendLine(@"                    }");
            sb.AppendLine(@"                }");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"            #endregion Filter TrueFalse");

            return "";
        }
        private string GenerateDBCodeOfTypeNumberFunction(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.Int32") || c.PropertyType.FullName.Contains("System.Single") || c.PropertyType.FullName.Contains("System.Double")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionNumberField dbFilteringNumberField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            switch (dbFilteringNumberField.ReportCondition)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + " > dbFilteringNumberField.NumberCondition);");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + " < dbFilteringNumberField.NumberCondition);");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + " == dbFilteringNumberField.NumberCondition);");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                default:");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");


            }

            return "";
        }
        private string GenerateDBCodeOfTypeNumber(Type type, StringBuilder sb, string PartialFileName)
        {
            sb.AppendLine(@"            #region Filter Number");
            sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbFilteringNumberFieldList.Count > 0))");
            sb.AppendLine(@"            {");
            sb.AppendLine(@"                foreach (ReportConditionNumberField dbFilteringNumberField in reportTreeNode.dbFilteringNumberFieldList)");
            sb.AppendLine(@"                {");
            sb.AppendLine(@"                    switch (reportTreeNode.Text)");
            sb.AppendLine(@"                    {");
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.Int32") || c.PropertyType.FullName.Contains("System.Single") || c.PropertyType.FullName.Contains("System.Double")))
            {
                sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                sb.AppendLine(@"                            report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "(report" + PartialFileName + "ModelQ, reportTreeNode, dbFilteringNumberField);");
                sb.AppendLine(@"                            break;");
            }
            sb.AppendLine(@"                        default:");
            sb.AppendLine(@"                            break;");
            sb.AppendLine(@"                    }");
            sb.AppendLine(@"                }");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"            #endregion Filter Number");

            return "";
        }
        private string GenerateDBCodeOfTypeTextFunction(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.String")))
            {
                if (propertyInfo.Name == "Pol_Source_Site_Last_Obs_Issue_Filtering"
                       || propertyInfo.Name == "Pol_Source_Site_Obs_Issue_Observation_Sentence"
                       || propertyInfo.Name == "Pol_Source_Site_Obs_Issue_Observation_Selection")
                {
                    continue;
                }

                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionTextField dbFilteringTextField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            switch (dbFilteringTextField.ReportCondition)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionContain:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + @".ToLower().Contains(dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" "")));");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionStart:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + @".ToLower().StartsWith(dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" "")));");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionEnd:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + @".ToLower().EndsWith(dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" "")));");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + @".ToLower() == dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" ""));");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => String.Compare(c." + propertyInfo.Name + @".ToLower(), dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" "")) > 0 );");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => String.Compare(c." + propertyInfo.Name + @".ToLower(), dbFilteringTextField.TextCondition.ToLower().Replace(""*"", "" "")) < 0 );");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"                default:");
                sb.AppendLine(@"                    break;");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeText(Type type, StringBuilder sb, string PartialFileName)
        {
            sb.AppendLine(@"            #region Filter Text");
            sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbFilteringTextFieldList.Count > 0))");
            sb.AppendLine(@"            {");
            sb.AppendLine(@"                foreach (ReportConditionTextField dbFilteringTextField in reportTreeNode.dbFilteringTextFieldList)");
            sb.AppendLine(@"                {");
            sb.AppendLine(@"                    switch (reportTreeNode.Text)");
            sb.AppendLine(@"                    {");
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.String")))
            {
                if (propertyInfo.Name == "Pol_Source_Site_Last_Obs_Issue_Filtering"
                    || propertyInfo.Name == "Pol_Source_Site_Obs_Issue_Observation_Sentence"
                    || propertyInfo.Name == "Pol_Source_Site_Obs_Issue_Observation_Selection")
                {
                    continue;
                }
                sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                sb.AppendLine(@"                            report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "(report" + PartialFileName + "ModelQ, reportTreeNode, dbFilteringTextField);");
                sb.AppendLine(@"                            break;");
            }
            sb.AppendLine(@"                        default:");
            sb.AppendLine(@"                            break;");
            sb.AppendLine(@"                    }");
            sb.AppendLine(@"                }");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"            #endregion Filter Text");

            return "";
        }
        private string GenerateDBCodeOfTypeDateFunctionYEAR(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_YEAR(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionDateField reportConditionDateField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            if (reportConditionDateField.DateTimeConditionYear != null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null && reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year > reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute > reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year < reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute < reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute == reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear != null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year > reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year < reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear != null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year > reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year < reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear != null && reportConditionDateField.DateTimeConditionMonth != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year > reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year < reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year > reportConditionDateField.DateTimeConditionYear);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year < reportConditionDateField.DateTimeConditionYear);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Year == reportConditionDateField.DateTimeConditionYear);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeDateFunctionMONTH(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_MONTH(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionDateField reportConditionDateField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null && reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute > reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute < reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute == reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth != null && reportConditionDateField.DateTimeConditionDay != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month > reportConditionDateField.DateTimeConditionMonth);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month < reportConditionDateField.DateTimeConditionMonth);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Month == reportConditionDateField.DateTimeConditionMonth);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeDateFunctionDAY(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_DAY(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionDateField reportConditionDateField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null && reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute > reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour)");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute < reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute == reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay != null && reportConditionDateField.DateTimeConditionHour != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day > reportConditionDateField.DateTimeConditionDay);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day < reportConditionDateField.DateTimeConditionDay);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Day == reportConditionDateField.DateTimeConditionDay);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeDateFunctionHOUR(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_HOUR(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionDateField reportConditionDateField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay == null && reportConditionDateField.DateTimeConditionHour != null && reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute > reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     || (c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute < reportConditionDateField.DateTimeConditionMinute));");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour");
                sb.AppendLine(@"                                                                     && c." + propertyInfo.Name + ".Value.Minute == reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            else if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay == null && reportConditionDateField.DateTimeConditionHour != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour > reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour < reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Hour == reportConditionDateField.DateTimeConditionHour);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeDateFunctionMINUTE(Type type, StringBuilder sb, string PartialFileName)
        {
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"        public IQueryable<" + type.Name + "> ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_MINUTE(IQueryable<" + type.Name + "> report" + PartialFileName + "ModelQ, ReportTreeNode reportTreeNode, ReportConditionDateField reportConditionDateField)");
                sb.AppendLine(@"        {");
                sb.AppendLine(@"            if (reportConditionDateField.DateTimeConditionYear == null && reportConditionDateField.DateTimeConditionMonth == null && reportConditionDateField.DateTimeConditionDay == null && reportConditionDateField.DateTimeConditionHour == null && reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"            {");
                sb.AppendLine(@"                switch (reportConditionDateField.ReportCondition)");
                sb.AppendLine(@"                {");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionBigger:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Minute > reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionSmaller:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Minute < reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    case ReportConditionEnum.ReportConditionEqual:");
                sb.AppendLine(@"                        report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => c." + propertyInfo.Name + ".Value.Minute == reportConditionDateField.DateTimeConditionMinute);");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                    default:");
                sb.AppendLine(@"                        break;");
                sb.AppendLine(@"                }");
                sb.AppendLine(@"            }");
                sb.AppendLine(@"            return report" + PartialFileName + "ModelQ;");
                sb.AppendLine(@"        }");
            }

            return "";
        }
        private string GenerateDBCodeOfTypeDate(Type type, StringBuilder sb, string PartialFileName)
        {
            sb.AppendLine(@"            #region Filter Date");
            sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbFilteringDateFieldList.Count > 0))");
            sb.AppendLine(@"            {");
            sb.AppendLine(@"                foreach (ReportConditionDateField reportConditionDateField in reportTreeNode.dbFilteringDateFieldList)");
            sb.AppendLine(@"                {");
            sb.AppendLine(@"                    switch (reportTreeNode.Text)");
            sb.AppendLine(@"                    {");
            foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains("System.DateTime")))
            {
                sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                sb.AppendLine(@"                            if (reportConditionDateField.DateTimeConditionYear != null)");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_YEAR(report" + PartialFileName + "ModelQ, reportTreeNode, reportConditionDateField);");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            else if (reportConditionDateField.DateTimeConditionMonth != null)");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_MONTH(report" + PartialFileName + "ModelQ, reportTreeNode, reportConditionDateField);");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            else if (reportConditionDateField.DateTimeConditionDay != null)");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_DAY(report" + PartialFileName + "ModelQ, reportTreeNode, reportConditionDateField);");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            else if (reportConditionDateField.DateTimeConditionHour != null)");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_HOUR(report" + PartialFileName + "ModelQ, reportTreeNode, reportConditionDateField);");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            else if (reportConditionDateField.DateTimeConditionMinute != null)");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = ReportServiceGenerated" + PartialFileName + "_" + propertyInfo.Name + "_MINUTE(report" + PartialFileName + "ModelQ, reportTreeNode, reportConditionDateField);");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            break;");
            }
            sb.AppendLine(@"                        default:");
            sb.AppendLine(@"                            break;");
            sb.AppendLine(@"                    }");
            sb.AppendLine(@"                }");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"            #endregion Filter Date");

            return "";
        }
        private string GenerateDBCodeOfTypeSorting(Type type, StringBuilder sb, string PartialFileName)
        {
            sb.AppendLine(@"            #region Sorting");
            sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbSortingField.ReportSorting != ReportSortingEnum.Error).OrderBy(c => c.dbSortingField.Ordinal))");
            sb.AppendLine(@"            {");
            sb.AppendLine(@"                if (reportTreeNode.dbSortingField.ReportSorting != ReportSortingEnum.Error)");
            sb.AppendLine(@"                {");
            sb.AppendLine(@"                    switch (reportTreeNode.Text)");
            sb.AppendLine(@"                    {");
            foreach (PropertyInfo propertyInfo in type.GetProperties())
            {
                sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                sb.AppendLine(@"                            {");
                sb.AppendLine(@"                                if (reportTreeNode.dbSortingField.ReportSorting == ReportSortingEnum.ReportSortingAscending)");
                sb.AppendLine(@"                                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.OrderBy(c => c." + propertyInfo.Name + ");");
                sb.AppendLine(@"                                else");
                sb.AppendLine(@"                                    report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.OrderByDescending(c => c." + propertyInfo.Name + ");");
                sb.AppendLine(@"                            }");
                sb.AppendLine(@"                            break;");
            }
            sb.AppendLine(@"                        default:");
            sb.AppendLine(@"                            break;");
            sb.AppendLine(@"                    }");
            sb.AppendLine(@"                }");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"            #endregion Sorting");

            return "";
        }
        private string GenerateDBCodeOfTypeEnum(Type type, StringBuilder sb, string PartialFileName)
        {
            List<Type> typeList = new List<Type>()
            {
                typeof(FilePurposeEnum), typeof(FileTypeEnum), typeof(TranslationStatusEnum), typeof(BoxModelResultTypeEnum), typeof(InfrastructureTypeEnum),
                typeof(FacilityTypeEnum), typeof(AerationTypeEnum), typeof(PreliminaryTreatmentTypeEnum), typeof(PrimaryTreatmentTypeEnum),
                typeof(SecondaryTreatmentTypeEnum), typeof(TertiaryTreatmentTypeEnum), typeof(TreatmentTypeEnum), typeof(DisinfectionTypeEnum),
                typeof(CollectionSystemTypeEnum), typeof(AlarmSystemTypeEnum), typeof(ScenarioStatusEnum), typeof(DailyOrHourlyDataEnum),
                typeof(StorageDataTypeEnum), typeof(LanguageEnum), typeof(SampleTypeEnum), typeof(BeaufortScaleEnum), typeof(AnalyzeMethodEnum),
                typeof(SampleMatrixEnum), typeof(LaboratoryEnum), typeof(SampleStatusEnum), typeof(SamplingPlanTypeEnum), typeof(LabSheetTypeEnum),
                typeof(LabSheetStatusEnum), typeof(PolSourceInactiveReasonEnum), typeof(PolSourceObsInfoEnum), typeof(AddressTypeEnum),
                typeof(StreetTypeEnum), typeof(ContactTitleEnum), typeof(EmailTypeEnum), typeof(TelTypeEnum), typeof(TideTextEnum),
                typeof(TideDataTypeEnum), typeof(SpecialTableTypeEnum), typeof(MWQMSiteLatestClassificationEnum), typeof(PolSourceIssueRiskEnum),
                typeof(MikeScenarioSpecialResultKMLTypeEnum)

            };
            foreach (Type typeToDo in typeList)
            {

                if (type.GetProperties().Where(c => c.PropertyType.FullName.Contains(typeToDo.FullName)).Any())
                {
                    string TypeTextShort = typeToDo.FullName.Replace("CSSPEnumsDLL.Enums.", "").Replace("Enum", "");

                    sb.AppendLine(@"            #region Filter " + TypeTextShort + "Enum");
                    sb.AppendLine(@"            foreach (ReportTreeNode reportTreeNode in reportTreeNodeList.Where(c => c.dbFilteringEnumFieldList.Count > 0 && c.ReportFieldType == ReportFieldTypeEnum." + TypeTextShort + "))");
                    sb.AppendLine(@"            {");
                    sb.AppendLine(@"                foreach (ReportConditionEnumField reportEnumField in reportTreeNode.dbFilteringEnumFieldList)");
                    sb.AppendLine(@"                {");
                    sb.AppendLine(@"                    switch (reportTreeNode.Text)");
                    sb.AppendLine(@"                    {");
                    foreach (PropertyInfo propertyInfo in type.GetProperties().Where(c => c.PropertyType.FullName.Contains(typeToDo.FullName)))
                    {
                        sb.AppendLine(@"                        case """ + propertyInfo.Name + @""":");
                        sb.AppendLine(@"                            if (reportEnumField.ReportCondition == ReportConditionEnum.ReportConditionEqual)");
                        sb.AppendLine(@"                            {");
                        sb.AppendLine(@"                                List<" + TypeTextShort + "Enum> " + TypeTextShort + "EqualList = new List<" + TypeTextShort + "Enum>();");
                        sb.AppendLine(@"                                List<string> " + TypeTextShort + @"TextList = reportEnumField.EnumConditionText.Split(""*"".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();");
                        sb.AppendLine(@"                                foreach (string s in " + TypeTextShort + "TextList)");
                        sb.AppendLine(@"                                {");
                        sb.AppendLine(@"                                    bool Found = false;");
                        sb.AppendLine(@"                                    for (int i = 1, count = Enum.GetNames(typeof(" + TypeTextShort + "Enum)).Count(); i < count; i++)");
                        sb.AppendLine(@"                                    {");
                        sb.AppendLine(@"                                        if (s == ((" + TypeTextShort + "Enum)i).ToString())");
                        sb.AppendLine(@"                                        {");
                        sb.AppendLine(@"                                            " + TypeTextShort + "EqualList.Add((" + TypeTextShort + "Enum)i);");
                        sb.AppendLine(@"                                        }");
                        sb.AppendLine(@"                                    }");
                        sb.AppendLine(@"                                    if (!Found)");
                        sb.AppendLine(@"                                        " + TypeTextShort + "EqualList.Add(" + TypeTextShort + "Enum.Error);");
                        sb.AppendLine(@"                                }");
                        sb.AppendLine(@"                                report" + PartialFileName + "ModelQ = report" + PartialFileName + "ModelQ.Where(c => " + TypeTextShort + "EqualList.Contains((" + TypeTextShort + "Enum)c." + propertyInfo.Name + "));");
                        sb.AppendLine(@"                            }");
                        sb.AppendLine(@"                            break;");
                    }
                    sb.AppendLine(@"                        default:");
                    sb.AppendLine(@"                            break;");
                    sb.AppendLine(@"                    }");
                    sb.AppendLine(@"                }");
                    sb.AppendLine(@"            }");
                    sb.AppendLine(@"            #endregion Filter " + TypeTextShort + "Enum");
                }
            }

            return "";
        }
        public string GenerateGetReportTypeChild(ReportTreeNode reportTreeNode, StringBuilder sb)
        {
            sb.AppendLine(@"                    case """ + reportTreeNode.Text + @""":");
            sb.AppendLine(@"                        return typeof(Report" + reportTreeNode.Text + "Model);");

            string retStr = "";
            foreach (ReportTreeNode RTN in reportTreeNode.Nodes)
            {
                if (RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableSelectable || RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableNotSelectable)
                {
                    retStr = GenerateGetReportTypeChild(RTN, sb);
                    if (!string.IsNullOrWhiteSpace(retStr))
                        return retStr;
                }
            }

            return "";
        }
        public string GenerateModelsChild(ReportTreeNode reportTreeNode, StringBuilder sb)
        {

            string retStr = reportBaseService.GenerateModel(reportTreeNode, sb);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                return retStr;
            }

            foreach (ReportTreeNode RTN in reportTreeNode.Nodes)
            {
                if (RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableSelectable || RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableNotSelectable)
                {
                    retStr = GenerateModelsChild(RTN, sb);
                    if (!string.IsNullOrWhiteSpace(retStr))
                        return retStr;
                }
            }

            return "";
        }
        public void GetTreeNodeTypeText(ReportTreeNode reportTreeNode, StringBuilder sb)
        {
            sb.Append(reportTreeNode.Text + ", ");

            foreach (ReportTreeNode RTN in reportTreeNode.Nodes)
            {
                if (RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableSelectable || RTN.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableNotSelectable)
                {
                    GetTreeNodeTypeText(RTN, sb);
                }
            }
        }
        public void GenerateModels()
        {
            StringBuilder sb = new StringBuilder();
            FileInfo fi = new FileInfo(@"C:\CSSP latest code old\CSSPModelsDLL\CSSPModelsDLL\Models\ReportGeneratedModel.cs");

            richTextBoxResults.Text = @"Generating File [C:\CSSP latest code old\CSSPModelsDLL\CSSPModelsDLL\Models\ReportGeneratedModel.cs]\r\n";

            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.Nodes[0];

            if (reportTreeNode == null)
            {
                richTextBoxResults.AppendText("ERROR: (ReportTreeNode)treeViewCSSP.Nodes[0] does not return the first node" + "\r\n");
                return;
            }

            sb.AppendLine(@"using CSSPEnumsDLL.Enums;");
            sb.AppendLine(@"using System;");
            sb.AppendLine(@"using System.Collections.Generic;");
            sb.AppendLine(@"using System.Linq;");
            sb.AppendLine(@"using System.Text;");
            sb.AppendLine(@"using System.Threading.Tasks;");
            sb.AppendLine(@"using System.Windows.Forms;");
            sb.AppendLine(@"");
            sb.AppendLine(@"namespace CSSPModelsDLL.Models");
            sb.AppendLine(@"{");
            sb.AppendLine(@"    public class ReportBase");
            sb.AppendLine(@"    {");
            sb.AppendLine(@"        public ReportBase()");
            sb.AppendLine(@"        {");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"");
            sb.AppendLine(@"        public string AllowableReportType()");
            sb.AppendLine(@"        {");

            StringBuilder sbART = new StringBuilder();
            GetTreeNodeTypeText(reportTreeNode, sbART);

            sb.AppendLine(@"            return """ + sbART.ToString() + @""";");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"");
            sb.AppendLine(@"        public Type GetReportType(string TypeText)");
            sb.AppendLine(@"        {");
            sb.AppendLine(@"            switch (TypeText)");
            sb.AppendLine(@"            {");

            string retStr = GenerateGetReportTypeChild(reportTreeNode, sb);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                richTextBoxResults.AppendText(retStr + "\r\n");
                return;
            }

            sb.AppendLine(@"                default:");
            sb.AppendLine(@"                    return null;");
            sb.AppendLine(@"            }");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"    }");

            reportTreeNode = (ReportTreeNode)treeViewCSSP.Nodes[0];

            retStr = GenerateModelsChild(reportTreeNode, sb);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                richTextBoxResults.AppendText(retStr + "\r\n");
                return;
            }

            richTextBoxResults.AppendText("\r\nPlease recompile CSSPModelsDLL");

            sb.AppendLine(@"}");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();
        }
        public void GetID_TVText(string url)
        {
            reportBaseService.LastHref = url;
            reportBaseService.LastCSSPTVText = "";
            if (!string.IsNullOrWhiteSpace(reportBaseService.LastHref))
            {
                int StartPos = reportBaseService.LastHref.IndexOf("|||") + 3;
                if (StartPos > 0)
                {
                    int TVTextStartPos = reportBaseService.LastHref.LastIndexOf("/", StartPos);
                    if (TVTextStartPos > 0)
                    {
                        reportBaseService.LastCSSPTVText = reportBaseService.LastHref.Substring(TVTextStartPos + 1, StartPos - TVTextStartPos - 4);
                    }
                    int EndPos = reportBaseService.LastHref.IndexOf("|||", StartPos);
                    if (EndPos > 0)
                    {
                        if (StartPos < EndPos)
                        {
                            int ID = 0;
                            if (int.TryParse(reportBaseService.LastHref.Substring(StartPos, EndPos - StartPos), out ID))
                            {
                                textBoxStartID.Text = ID.ToString();
                            }
                        }
                    }
                }
            }

            lblCSSPTVText.Text = reportBaseService.LastCSSPTVText;
        }
        public int GetNextSortingOdinalNumber()
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return 0;

            ReportTreeNode reportTreeNodeParent = (ReportTreeNode)reportTreeNode.Parent;

            if (reportTreeNodeParent == null)
                return 0;

            int Max = 0;
            foreach (ReportTreeNode RTN in reportTreeNodeParent.Nodes)
            {
                Max = Math.Max(RTN.dbSortingField.Ordinal, Max);
            }

            return Max + 1;
        }
        public string InsertTemplateTextInExcel(Microsoft.Office.Interop.Excel.Worksheet workSheet)
        {
            workSheet.Cells[1, 1] = richTextBoxExample.Text + "\r\n";

            return "";
        }
        public string InsertTemplateTextInWord(Microsoft.Office.Interop.Word.Document doc)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
            {
                lblEmptyPanelMessage.Text = "Please select an item.";
                return lblEmptyPanelMessage.Text;
            }

            reportBaseService.GetTreeViewSelectedStatusWord(reportTreeNode, doc, 0);

            return "";
        }
        private void listBoxDBFilteringEnumSelectedIndexChanged()
        {
            SelectedObjectCollection selectedObjectCollection = listBoxDBFilteringEnum.SelectedItems;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            string EnumConditionText = "";
            foreach (var item in selectedObjectCollection)
            {
                EnumConditionText += item.ToString() + "*";
            }

            if (!string.IsNullOrWhiteSpace(EnumConditionText))
            {
                EnumConditionText = EnumConditionText.Substring(0, EnumConditionText.Length - 1);
            }

            if (reportTreeNode.dbFilteringEnumFieldList.Count == 0)
                reportTreeNode.dbFilteringEnumFieldList.Add(new ReportConditionEnumField());

            reportTreeNode.dbFilteringEnumFieldList[0].ReportCondition = ReportConditionEnum.ReportConditionEqual;
            reportTreeNode.dbFilteringEnumFieldList[0].EnumConditionText = EnumConditionText;
        }
        private void listBoxReportConditionEnumSelectedIndexChanged()
        {
            SelectedObjectCollection selectedObjectCollection = listBoxReportConditionEnum.SelectedItems;
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            string EnumConditionText = "";
            foreach (var item in selectedObjectCollection)
            {
                EnumConditionText += item.ToString() + "*";
            }

            if (!string.IsNullOrWhiteSpace(EnumConditionText))
            {
                EnumConditionText = EnumConditionText.Substring(0, EnumConditionText.Length - 1);
            }

            if (reportTreeNode.reportConditionEnumFieldList.Count == 0)
                reportTreeNode.reportConditionEnumFieldList.Add(new ReportConditionEnumField());

            reportTreeNode.reportConditionEnumFieldList[0].ReportCondition = ReportConditionEnum.ReportConditionEqual;
            reportTreeNode.reportConditionEnumFieldList[0].EnumConditionText = EnumConditionText;
        }
        public void OpenFile()
        {
            FileInfo fi = (FileInfo)comboBoxTemplateDocuments.SelectedItem;
            fi = new FileInfo(fi.FullName);

            if (radioButtonCSV.Checked)
            {
                if (radioButtonShowResultFiles.Checked)
                {
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;
                    excelApp.Workbooks.Open(fi.FullName);
                }
                else
                {
                    Process.Start("notepad.exe", fi.FullName);
                }
            }
            else if (radioButtonWord.Checked)
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                wordApp.Documents.Open(fi.FullName);
            }
            else if (radioButtonExcel.Checked)
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                excelApp.Workbooks.Open(fi.FullName);
            }
            else if (radioButtonKML.Checked)
            {
                Process.Start("notepad.exe", fi.FullName);
            }

        }
        public void RefreshTemplateDocuments()
        {
            butCreateNew.Enabled = false;
            //butProduceDocument.Enabled = false;
            butTestSelectedTemplate.Enabled = false;

            if (radioButtonShowTemplateFiles.Checked)
            {
                butCreateNew.Enabled = true;
            }

            comboBoxTemplateDocuments.Items.Clear();
            comboBoxTemplateDocuments.Text = "";
            DirectoryInfo di = new DirectoryInfo(CSSPReportTemplatesPath);
            List<FileInfo> fileInfoList = di.GetFiles().ToList();
            foreach (FileInfo fi in fileInfoList)
            {
                if (radioButtonWord.Checked)
                {
                    if (fi.Extension.ToLower() == ".docx")
                    {
                        if (radioButtonShowTemplateFiles.Checked && fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                        else if (!radioButtonShowTemplateFiles.Checked && !fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                    }
                }
                else if (radioButtonExcel.Checked)
                {
                    if (fi.Extension.ToLower() == ".xlsx")
                    {
                        if (radioButtonShowTemplateFiles.Checked && fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                        else if (!radioButtonShowTemplateFiles.Checked && !fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                    }
                }
                else if (radioButtonKML.Checked)
                {
                    if (fi.Extension.ToLower() == ".kml")
                    {
                        if (radioButtonShowTemplateFiles.Checked && fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                        else if (!radioButtonShowTemplateFiles.Checked && !fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                    }
                }
                else if (radioButtonCSV.Checked)
                {
                    if (fi.Extension.ToLower() == ".csv")
                    {
                        if (radioButtonShowTemplateFiles.Checked && fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                        else if (!radioButtonShowTemplateFiles.Checked && !fi.FullName.StartsWith(CSSPReportTemplatesPath + "Template_"))
                        {
                            comboBoxTemplateDocuments.Items.Add((FileInfo)fi);
                        }
                    }
                }
            }
            if (comboBoxTemplateDocuments.Items.Count > 0)
            {
                comboBoxTemplateDocuments.SelectedIndex = 0;
                butOpenDoc.Enabled = true;
                butRemoveFile.Enabled = true;
                if (radioButtonShowTemplateFiles.Checked)
                {
                    //butProduceDocument.Enabled = true;
                    butTestSelectedTemplate.Enabled = true;
                }
            }
            else
            {
                butOpenDoc.Enabled = false;
                butRemoveFile.Enabled = false;
                //butProduceDocument.Enabled = false;
                butTestSelectedTemplate.Enabled = false;
            }
        }
        private void RemoveFile()
        {
            FileInfo fi = (FileInfo)comboBoxTemplateDocuments.SelectedItem;
            fi = new FileInfo(fi.FullName);

            DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete \r\n[" + fi.FullName + "]", "Deleting template document", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (fi.Exists)
                {
                    fi.Delete();
                }
            }
            RefreshTemplateDocuments();
        }
        public void SetItemDate(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            // DB Filtering Date
            if (reportTreeNode.dbFilteringDateFieldList.Count == 0)
            {
                comboBoxDBFilteringDate.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxDBFilteringDate.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxDBFilteringDate.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].ReportCondition)
                    {
                        comboBoxDBFilteringDate.SelectedItem = comboBoxDBFilteringDate.Items[i];
                        break;
                    }
                }

                // DB Filtering Year
                if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionYear == null)
                {
                    comboBoxDBFilteringYear.SelectedItem = comboBoxDBFilteringYear.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxDBFilteringYear.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxDBFilteringYear.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionYear)
                        {
                            comboBoxDBFilteringYear.SelectedItem = comboBoxDBFilteringYear.Items[i];
                            break;
                        }
                    }
                }

                // DB Filtering Month
                if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMonth == null)
                {
                    comboBoxDBFilteringMonth.SelectedItem = comboBoxDBFilteringMonth.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxDBFilteringMonth.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxDBFilteringMonth.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMonth)
                        {
                            comboBoxDBFilteringMonth.SelectedItem = comboBoxDBFilteringMonth.Items[i];
                            break;
                        }
                    }
                }

                // DB Filtering Day
                if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionDay == null)
                {
                    comboBoxDBFilteringDay.SelectedItem = comboBoxDBFilteringDay.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxDBFilteringDay.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxDBFilteringDay.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionDay)
                        {
                            comboBoxDBFilteringDay.SelectedItem = comboBoxDBFilteringDay.Items[i];
                            break;
                        }
                    }
                }

                // DB Filtering Hour
                if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionHour == null)
                {
                    comboBoxDBFilteringHour.SelectedItem = comboBoxDBFilteringHour.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxDBFilteringHour.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxDBFilteringHour.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionHour)
                        {
                            comboBoxDBFilteringHour.SelectedItem = comboBoxDBFilteringHour.Items[i];
                            break;
                        }
                    }
                }

                // DB Filtering Minute
                if (reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMinute == null)
                {
                    comboBoxDBFilteringMinute.SelectedItem = comboBoxDBFilteringMinute.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxDBFilteringMinute.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxDBFilteringMinute.Items[i]).ID == (int)reportTreeNode.dbFilteringDateFieldList[0].DateTimeConditionMinute)
                        {
                            comboBoxDBFilteringMinute.SelectedItem = comboBoxDBFilteringMinute.Items[i];
                            break;
                        }
                    }
                }
            }

            ////////////////////////////////////////////////////////////////
            // Report Condition Date
            if (reportTreeNode.reportConditionDateFieldList.Count == 0)
            {
                comboBoxReportConditionDate.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxReportConditionDate.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxReportConditionDate.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].ReportCondition)
                    {
                        comboBoxReportConditionDate.SelectedItem = comboBoxReportConditionDate.Items[i];
                        break;
                    }
                }

                // Report Condition Year
                if (reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionYear == null)
                {
                    comboBoxReportConditionYear.SelectedItem = comboBoxReportConditionYear.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxReportConditionYear.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxReportConditionYear.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionYear)
                        {
                            comboBoxReportConditionYear.SelectedItem = comboBoxReportConditionYear.Items[i];
                            break;
                        }
                    }
                }

                // Report Condition Month
                if (reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMonth == null)
                {
                    comboBoxReportConditionMonth.SelectedItem = comboBoxReportConditionMonth.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxReportConditionMonth.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxReportConditionMonth.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMonth)
                        {
                            comboBoxReportConditionMonth.SelectedItem = comboBoxReportConditionMonth.Items[i];
                            break;
                        }
                    }
                }

                // Report Condition Day
                if (reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionDay == null)
                {
                    comboBoxReportConditionDay.SelectedItem = comboBoxReportConditionDay.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxReportConditionDay.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxReportConditionDay.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionDay)
                        {
                            comboBoxReportConditionDay.SelectedItem = comboBoxReportConditionDay.Items[i];
                            break;
                        }
                    }
                }

                // Report Condition Hour
                if (reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionHour == null)
                {
                    comboBoxReportConditionHour.SelectedItem = comboBoxReportConditionHour.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxReportConditionHour.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxReportConditionHour.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionHour)
                        {
                            comboBoxReportConditionHour.SelectedItem = comboBoxReportConditionHour.Items[i];
                            break;
                        }
                    }
                }

                // Report Condition Minute
                if (reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMinute == null)
                {
                    comboBoxReportConditionMinute.SelectedItem = comboBoxReportConditionMinute.Items[0];
                }
                else
                {
                    for (int i = 0, count = comboBoxReportConditionMinute.Items.Count; i < count; i++)
                    {
                        if (((ReportItemModel)comboBoxReportConditionMinute.Items[i]).ID == (int)reportTreeNode.reportConditionDateFieldList[0].DateTimeConditionMinute)
                        {
                            comboBoxReportConditionMinute.SelectedItem = comboBoxReportConditionMinute.Items[i];
                            break;
                        }
                    }
                }
            }
        }
        public void SetItemEnum(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            listBoxDBFilteringEnum.BackColor = Color.White;

            if (reportTreeNode.dbFilteringEnumFieldList.Count == 0)
            {
                // nothing
            }
            else
            {
                List<string> selectionList = reportTreeNode.dbFilteringEnumFieldList[0].EnumConditionText.Split("*".ToArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (string s in selectionList)
                {
                    for (int i = 0, count = listBoxDBFilteringEnum.Items.Count; i < count; i++)
                    {
                        if (listBoxDBFilteringEnum.Items[i].ToString() == s)
                        {
                            listBoxDBFilteringEnum.SetSelected(i, true);
                        }
                    }
                }
            }

            listBoxReportConditionEnum.BackColor = Color.White;

            if (reportTreeNode.reportConditionEnumFieldList.Count == 0)
            {
                // nothing
            }
            else
            {
                List<string> selectionList2 = reportTreeNode.reportConditionEnumFieldList[0].EnumConditionText.Split("*".ToArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (string s in selectionList2)
                {
                    for (int i = 0, count = listBoxReportConditionEnum.Items.Count; i < count; i++)
                    {
                        if (listBoxReportConditionEnum.Items[i].ToString() == s)
                        {
                            listBoxReportConditionEnum.SetSelected(i, true);
                        }
                    }
                }
            }

        }
        public void SetItemNumber(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            textBoxDBFilteringNumber.BackColor = Color.White;

            // DB Filtering Number 1
            if (reportTreeNode.dbFilteringNumberFieldList.Count == 0)
            {
                comboBoxDBFilteringNumber.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxDBFilteringNumber.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxDBFilteringNumber.Items[i]).ID == (int)reportTreeNode.dbFilteringNumberFieldList[0].ReportCondition)
                    {
                        comboBoxDBFilteringNumber.SelectedItem = comboBoxDBFilteringNumber.Items[i];
                        break;
                    }
                }

                if (comboBoxDBFilteringNumber.SelectedIndex == 0)
                {
                    textBoxDBFilteringNumber.Text = reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition.ToString();
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition.ToString()))
                    {
                        textBoxDBFilteringNumber.BackColor = Color.Red;
                    }
                    textBoxDBFilteringNumber.Text = reportTreeNode.dbFilteringNumberFieldList[0].NumberCondition.ToString();
                }
            }


            //////////////////////////////////

            textBoxReportConditionNumber.BackColor = Color.White;

            if (reportTreeNode.reportConditionNumberFieldList.Count == 0)
            {
                comboBoxReportConditionNumber.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxReportConditionNumber.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxReportConditionNumber.Items[i]).ID == (int)reportTreeNode.reportConditionNumberFieldList[0].ReportCondition)
                    {
                        comboBoxReportConditionNumber.SelectedItem = comboBoxReportConditionNumber.Items[i];
                        break;
                    }
                }

                if (comboBoxReportConditionNumber.SelectedIndex == 0)
                {
                    textBoxReportConditionNumber.Text = reportTreeNode.reportConditionNumberFieldList[0].NumberCondition.ToString();
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(reportTreeNode.reportConditionNumberFieldList[0].NumberCondition.ToString()))
                    {
                        textBoxReportConditionNumber.BackColor = Color.Red;
                    }

                    textBoxReportConditionNumber.Text = reportTreeNode.reportConditionNumberFieldList[0].NumberCondition.ToString();
                }
            }
        }
        public void SetItemText(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            textBoxDBFilteringText.BackColor = Color.White;

            // DB Filtering Text
            if (reportTreeNode.dbFilteringTextFieldList.Count == 0)
            {
                comboBoxDBFilteringText.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxDBFilteringText.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxDBFilteringText.Items[i]).ID == (int)reportTreeNode.dbFilteringTextFieldList[0].ReportCondition)
                    {
                        comboBoxDBFilteringText.SelectedItem = comboBoxDBFilteringText.Items[i];
                        break;
                    }
                }

                if (comboBoxDBFilteringText.SelectedIndex == 0)
                {
                    textBoxDBFilteringText.Text = reportTreeNode.dbFilteringTextFieldList[0].TextCondition.ToString();
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(reportTreeNode.dbFilteringTextFieldList[0].TextCondition.ToString()))
                    {
                        textBoxDBFilteringText.BackColor = Color.Red;
                    }

                    textBoxDBFilteringText.Text = reportTreeNode.dbFilteringTextFieldList[0].TextCondition.ToString();
                }
            }

            /////////////////////////////////

            textBoxReportConditionText.BackColor = Color.White;
            if (reportTreeNode.reportConditionTextFieldList.Count == 0)
            {
                comboBoxReportConditionText.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxReportConditionText.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxReportConditionText.Items[i]).ID == (int)reportTreeNode.reportConditionTextFieldList[0].ReportCondition)
                    {
                        comboBoxReportConditionText.SelectedItem = comboBoxReportConditionText.Items[i];
                        break;
                    }
                }

                if (comboBoxReportConditionText.SelectedIndex == 0)
                {
                    textBoxReportConditionText.Text = reportTreeNode.reportConditionTextFieldList[0].TextCondition.ToString();
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(reportTreeNode.reportConditionTextFieldList[0].TextCondition.ToString()))
                    {
                        textBoxReportConditionText.BackColor = Color.Red;
                    }

                    textBoxReportConditionText.Text = reportTreeNode.reportConditionTextFieldList[0].TextCondition.ToString();
                }
            }
        }
        public void SetItemDBFilteringTrueFalse(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            // DB Filtering TrueFalse
            if (reportTreeNode.dbFilteringTrueFalseFieldList.Count == 0)
            {
                comboBoxDBFilteringTrueFalse.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxDBFilteringTrueFalse.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxDBFilteringTrueFalse.Items[i]).ID == (int)reportTreeNode.dbFilteringTrueFalseFieldList[0].ReportCondition)
                    {
                        comboBoxDBFilteringTrueFalse.SelectedItem = comboBoxDBFilteringTrueFalse.Items[i];
                        break;
                    }
                }
            }

            ///////////////////////
            // Report Condition TrueFalse
            if (reportTreeNode.reportConditionTrueFalseFieldList.Count == 0)
            {
                comboBoxReportConditionTrueFalse.SelectedIndex = 0;
            }
            else
            {
                for (int i = 0, count = comboBoxReportConditionTrueFalse.Items.Count; i < count; i++)
                {
                    if (((ReportItemModel)comboBoxReportConditionTrueFalse.Items[i]).ID == (int)reportTreeNode.reportConditionTrueFalseFieldList[0].ReportCondition)
                    {
                        comboBoxReportConditionTrueFalse.SelectedItem = comboBoxReportConditionTrueFalse.Items[i];
                        break;
                    }
                }
            }
        }
        public void SetItemDBFormatingDate(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFormatingField == null)
            {
                comboBoxDBFormatingDate.SelectedIndex = 0;
            }

            for (int i = 0, count = comboBoxDBFormatingDate.Items.Count; i < count; i++)
            {
                if (((ReportItemModel)comboBoxDBFormatingDate.Items[i]).ID == (int)reportTreeNode.dbFormatingField.ReportFormatingDate)
                {
                    comboBoxDBFormatingDate.SelectedItem = comboBoxDBFormatingDate.Items[i];
                    break;
                }
            }
        }
        public void SetItemReportFormatingDate(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportFormatingField == null)
            {
                comboBoxReportFormatingDate.SelectedIndex = 0;
            }

            for (int i = 0, count = comboBoxReportFormatingDate.Items.Count; i < count; i++)
            {
                if (((ReportItemModel)comboBoxReportFormatingDate.Items[i]).ID == (int)reportTreeNode.reportFormatingField.ReportFormatingDate)
                {
                    comboBoxReportFormatingDate.SelectedItem = comboBoxReportFormatingDate.Items[i];
                    break;
                }
            }
        }
        public void SetItemDBFormatingNumber(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbFormatingField == null)
            {
                comboBoxDBFormatingNumber.SelectedIndex = 0;
            }

            for (int i = 0, count = comboBoxDBFormatingNumber.Items.Count; i < count; i++)
            {
                if (((ReportItemModel)comboBoxDBFormatingNumber.Items[i]).ID == (int)reportTreeNode.dbFormatingField.ReportFormatingNumber)
                {
                    comboBoxDBFormatingNumber.SelectedItem = comboBoxDBFormatingNumber.Items[i];
                    break;
                }
            }
        }
        public void SetItemReportFormatingNumber(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            if (reportTreeNode.reportFormatingField == null)
            {
                comboBoxReportFormatingNumber.SelectedIndex = 0;
            }

            for (int i = 0, count = comboBoxReportFormatingNumber.Items.Count; i < count; i++)
            {
                if (((ReportItemModel)comboBoxReportFormatingNumber.Items[i]).ID == (int)reportTreeNode.reportFormatingField.ReportFormatingNumber)
                {
                    comboBoxReportFormatingNumber.SelectedItem = comboBoxReportFormatingNumber.Items[i];
                    break;
                }
            }
        }
        public void SetItemSorting(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            if (reportTreeNode.dbSortingField == null)
            {
                comboBoxDBSorting.SelectedIndex = 0;
            }

            for (int i = 0, count = comboBoxDBSorting.Items.Count; i < count; i++)
            {
                if (((ReportItemModel)comboBoxDBSorting.Items[i]).ID == (int)reportTreeNode.dbSortingField.ReportSorting)
                {
                    comboBoxDBSorting.SelectedItem = comboBoxDBSorting.Items[i];
                    break;
                }
            }

            if (reportTreeNode.dbSortingField.Ordinal == 0)
            {
                butSortingUp.Enabled = false;
                butSortingDown.Enabled = false;
            }
            else
            {
                butSortingUp.Enabled = true;
                butSortingDown.Enabled = true;
            }
            lblSortingOrdinal.Text = reportTreeNode.dbSortingField.Ordinal.ToString();
        }
        public void SetItemValues(ReportTreeNode reportTreeNode)
        {
            if (reportTreeNode == null)
                return;

            SetItemSorting(reportTreeNode);

            if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.DateAndTime)
            {
                SetItemReportFormatingDate(reportTreeNode);
                SetItemDBFormatingDate(reportTreeNode);
                SetItemDate(reportTreeNode);
            }
            else if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.NumberWithDecimal)
            {
                SetItemReportFormatingNumber(reportTreeNode);
                SetItemDBFormatingNumber(reportTreeNode);
                SetItemNumber(reportTreeNode);
            }
            else if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.NumberWhole)
            {
                SetItemNumber(reportTreeNode);
            }
            else if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.Text)
            {
                SetItemText(reportTreeNode);
            }
            else if (reportTreeNode.ReportFieldType == ReportFieldTypeEnum.TrueOrFalse)
            {
                SetItemDBFilteringTrueFalse(reportTreeNode);
            }
            else // this will do all Enum
            {
                SetItemEnum(reportTreeNode);
            }

        }
        public void SetPanelRightTopMiddleVisible(Panel panelRightTopMiddleToMakeVisible)
        {
            panelRightTopMiddleText.Visible = false;
            panelRightTopMiddleNumber.Visible = false;
            panelRightTopMiddleDate.Visible = false;
            panelRightTopMiddleBoolean.Visible = false;
            panelRightTopMiddleEnum.Visible = false;
            panelRightTopMiddleProperties.Visible = false;
            if (panelRightTopMiddleToMakeVisible != panelRightTopMiddleEmpty)
            {
                panelRightTopMiddleEmpty.Visible = false;
                panelRightTopMiddleProperties.Visible = true;
            }
            panelRightTopMiddleEmpty.Visible = false;
            panelRightTopMiddleToMakeVisible.Visible = true;

        }
        public void Setup()
        {
            treeViewCSSP.Dock = DockStyle.Fill;
            richTextBoxResults.Dock = DockStyle.Fill;
            webBrowserCSSPWebTools.Dock = DockStyle.Fill;
            splitContainer1.Dock = DockStyle.Fill;
            panelMiddleWebBrowser.Dock = DockStyle.Fill;
            panelMiddleTreeViewAndTest.Dock = DockStyle.Fill;
            panelMiddleWebBrowser.Visible = false;
            panelMiddleTreeViewAndTest.Visible = true;
            webBrowserCSSPWebTools.Navigate(StartWebAddressCSSP);
            lblCSSPTVText.Text = "";
            lblStatusValue.Text = "";
            panelRightTopMiddleProperties.Dock = DockStyle.Fill;
            panelRightTopMiddleEmpty.Dock = DockStyle.Fill;
            panelRightTop.Dock = DockStyle.Fill;
            panelRightTopMiddleEmpty.Visible = true;
            panelRightTopMiddleText.Visible = false;
            panelRightTopMiddleNumber.Visible = false;
            panelRightTopMiddleDate.Visible = false;
            panelRightTopMiddleBoolean.Visible = false;
            panelRightTopMiddleEnum.Visible = false;
            panelSorting.Visible = false;
            lblEmptyPanelMessage.Text = "";
            lblSortingOrdinal.Text = "";
            lblSelectedTreeViewText.Text = "";
            lblCurrentFilePath.Text = "";
            CreateCSSPReportTemplatesDirectory();
            comboBoxTemplateDocuments.ValueMember = "Name";
            comboBoxTemplateDocuments.DisplayMember = "Name";
            panelCreateNewFile.BringToFront();
            panelCreateNewFile.Visible = false;
            lblFormatReportCondition.Visible = false;
            lblFormatDatabaseFiltering.Visible = false;
            comboBoxDBFormatingNumber.Visible = false;
            comboBoxReportFormatingNumber.Visible = false;
            ToolTip toolTipGenerateModel = new ToolTip();
            toolTipGenerateModel.SetToolTip(butGenerateModels, @"C:\CSSP latest code old\CSSPModelsDLL\CSSPModelsDLL\Models\ReportGeneratedModel.cs");
            ToolTip toolTipGenerateDB = new ToolTip();
            toolTipGenerateDB.SetToolTip(butGenerateDBCode, @"C:\CSSP latest code old\CSSPWebToolsDBDLL\CSSPWebToolsDBDLL\Services\ReportServiceGenerated______.cs");
            //ToolTip toolTipGeneratReplace = new ToolTip();
            //toolTipGeneratReplace.SetToolTip(butGenerateDBGetAndReplace, @"C:\CSSP latest code old\CSSPReportWriterHelperDLL\CSSPReportWriterHelperDLL\Services\ReportGeneratedDBReplace.cs");

            RefreshTemplateDocuments();

            SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);

            comboBoxDBSorting.DisplayMember = "Text";
            comboBoxDBSorting.ValueMember = "ID";

            comboBoxDBSorting.Items.Add(new ReportItemModel() { ID = (int)ReportSortingEnum.Error, Text = "Sorting" });
            comboBoxDBSorting.Items.Add(new ReportItemModel() { ID = (int)ReportSortingEnum.ReportSortingAscending, Text = "Ascending" });
            comboBoxDBSorting.Items.Add(new ReportItemModel() { ID = (int)ReportSortingEnum.ReportSortingDescending, Text = "Descending" });
            comboBoxDBSorting.SelectedIndex = 0;

            comboBoxDBFormatingDate.DisplayMember = "Text";
            comboBoxDBFormatingDate.ValueMember = "ID";
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.Error, Text = "Formating" });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearOnly, Text = "Year only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly, Text = "Month (number) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly, Text = "Month (short) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly, Text = "Month (long) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateDayOnly, Text = "Day only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateDayOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateHourOnly, Text = "Hour only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateHourOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMinuteOnly, Text = "Minute only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMinuteOnly) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute) });
            comboBoxDBFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute) });
            comboBoxDBFormatingDate.SelectedIndex = 0;

            comboBoxReportFormatingDate.DisplayMember = "Text";
            comboBoxReportFormatingDate.ValueMember = "ID";
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.Error, Text = "Formating" });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearOnly, Text = "Year only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly, Text = "Month (number) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthDecimalOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly, Text = "Month (short) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthShortTextOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly, Text = "Month (long) only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMonthFullTextOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateDayOnly, Text = "Day only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateDayOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateHourOnly, Text = "Hour only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateHourOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateMinuteOnly, Text = "Minute only ex:" + reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateMinuteOnly) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDay) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDay) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDay) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthDecimalDayHourMinute) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthShortTextDayHourMinute) });
            comboBoxReportFormatingDate.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute, Text = reportBaseService.GetFormatDate(ReportFormatingDateEnum.ReportFormatingDateYearMonthFullTextDayHourMinute) });
            comboBoxReportFormatingDate.SelectedIndex = 0;

            comboBoxDBFormatingNumber.DisplayMember = "Text";
            comboBoxDBFormatingNumber.ValueMember = "ID";
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.Error, Text = "Formating" });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber0Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber0Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber1Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber1Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber2Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber2Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber3Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber3Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber4Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber4Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber5Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber5Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber6Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber6Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal) });
            comboBoxDBFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal) });
            comboBoxDBFormatingNumber.SelectedIndex = 0;

            comboBoxReportFormatingNumber.DisplayMember = "Text";
            comboBoxReportFormatingNumber.ValueMember = "ID";
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.Error, Text = "Formating" });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber0Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber0Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber1Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber1Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber2Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber2Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber3Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber3Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber4Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber4Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber5Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber5Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumber6Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumber6Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific0Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific1Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific2Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific3Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific4Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific5Decimal) });
            comboBoxReportFormatingNumber.Items.Add(new ReportItemModel() { ID = (int)ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal, Text = reportBaseService.GetFormatNumber(ReportFormatingNumberEnum.ReportFormatingNumberScientific6Decimal) });
            comboBoxReportFormatingNumber.SelectedIndex = 0;

            comboBoxDBFilteringTrueFalse.DisplayMember = "Text";
            comboBoxDBFilteringTrueFalse.ValueMember = "ID";
            comboBoxDBFilteringTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxDBFilteringTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionTrue, Text = "TRUE" });
            comboBoxDBFilteringTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionFalse, Text = "FALSE" });
            comboBoxDBFilteringTrueFalse.SelectedIndex = 0;

            comboBoxReportConditionTrueFalse.DisplayMember = "Text";
            comboBoxReportConditionTrueFalse.ValueMember = "ID";
            comboBoxReportConditionTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxReportConditionTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionTrue, Text = "TRUE" });
            comboBoxReportConditionTrueFalse.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionFalse, Text = "FALSE" });
            comboBoxReportConditionTrueFalse.SelectedIndex = 0;

            comboBoxDBFilteringText.DisplayMember = "Text";
            comboBoxDBFilteringText.ValueMember = "ID";
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionContain, Text = "CONTAIN" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionStart, Text = "START" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEnd, Text = "END" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxDBFilteringText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxDBFilteringText.SelectedIndex = 0;

            comboBoxReportConditionText.DisplayMember = "Text";
            comboBoxReportConditionText.ValueMember = "ID";
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionContain, Text = "CONTAIN" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionStart, Text = "START" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEnd, Text = "END" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxReportConditionText.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxReportConditionText.SelectedIndex = 0;

            comboBoxDBFilteringDate.DisplayMember = "Text";
            comboBoxDBFilteringDate.ValueMember = "ID";
            comboBoxDBFilteringDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxDBFilteringDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxDBFilteringDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxDBFilteringDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxDBFilteringDate.SelectedIndex = 0;

            comboBoxReportConditionDate.DisplayMember = "Text";
            comboBoxReportConditionDate.ValueMember = "ID";
            comboBoxReportConditionDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxReportConditionDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxReportConditionDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxReportConditionDate.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxReportConditionDate.SelectedIndex = 0;

            comboBoxDBFilteringNumber.DisplayMember = "Text";
            comboBoxDBFilteringNumber.ValueMember = "ID";
            comboBoxDBFilteringNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxDBFilteringNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxDBFilteringNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxDBFilteringNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxDBFilteringNumber.SelectedIndex = 0;

            comboBoxReportConditionNumber.DisplayMember = "Text";
            comboBoxReportConditionNumber.ValueMember = "ID";
            comboBoxReportConditionNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.Error, Text = "Filtering" });
            comboBoxReportConditionNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionEqual, Text = "EQUAL" });
            comboBoxReportConditionNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionBigger, Text = "BIGGER" });
            comboBoxReportConditionNumber.Items.Add(new ReportItemModel() { ID = (int)ReportConditionEnum.ReportConditionSmaller, Text = "SMALLER" });
            comboBoxReportConditionNumber.SelectedIndex = 0;


            // DBFiltering
            comboBoxDBFilteringYear.DisplayMember = "Text";
            comboBoxDBFilteringYear.ValueMember = "ID";
            comboBoxDBFilteringYear.Items.Add(new ReportItemModel() { ID = 0, Text = "Year" });
            for (int i = 1980; i < (DateTime.Now.Year + 2); i++)
            {
                comboBoxDBFilteringYear.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxDBFilteringYear.SelectedIndex = 0;

            comboBoxDBFilteringMonth.DisplayMember = "Text";
            comboBoxDBFilteringMonth.ValueMember = "ID";
            comboBoxDBFilteringMonth.Items.Add(new ReportItemModel() { ID = 0, Text = "Month" });
            for (int i = 1; i < 13; i++)
            {
                comboBoxDBFilteringMonth.Items.Add(new ReportItemModel() { ID = i, Text = DateTimeFormatInfo.CurrentInfo.GetMonthName(i) });
            }
            comboBoxDBFilteringMonth.SelectedIndex = 0;

            comboBoxDBFilteringDay.DisplayMember = "Text";
            comboBoxDBFilteringDay.ValueMember = "ID";
            comboBoxDBFilteringDay.Items.Add(new ReportItemModel() { ID = 0, Text = "Day" });
            for (int i = 1; i < 32; i++)
            {
                comboBoxDBFilteringDay.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxDBFilteringDay.SelectedIndex = 0;

            comboBoxDBFilteringHour.DisplayMember = "Text";
            comboBoxDBFilteringHour.ValueMember = "ID";
            comboBoxDBFilteringHour.Items.Add(new ReportItemModel() { ID = 0, Text = "Hour" });
            for (int i = 0; i < 24; i++)
            {
                comboBoxDBFilteringHour.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxDBFilteringHour.SelectedIndex = 0;

            comboBoxDBFilteringMinute.DisplayMember = "Text";
            comboBoxDBFilteringMinute.ValueMember = "ID";
            comboBoxDBFilteringMinute.Items.Add(new ReportItemModel() { ID = 0, Text = "Minute" });
            for (int i = 1; i < 60; i++)
            {
                comboBoxDBFilteringMinute.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxDBFilteringMinute.SelectedIndex = 0;


            // Report Condition
            comboBoxReportConditionYear.DisplayMember = "Text";
            comboBoxReportConditionYear.ValueMember = "ID";
            comboBoxReportConditionYear.Items.Add(new ReportItemModel() { ID = 0, Text = "Year" });
            for (int i = 1980; i < (DateTime.Now.Year + 2); i++)
            {
                comboBoxReportConditionYear.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxReportConditionYear.SelectedIndex = 0;

            comboBoxReportConditionMonth.DisplayMember = "Text";
            comboBoxReportConditionMonth.ValueMember = "ID";
            comboBoxReportConditionMonth.Items.Add(new ReportItemModel() { ID = 0, Text = "Month" });
            for (int i = 1; i < 13; i++)
            {
                comboBoxReportConditionMonth.Items.Add(new ReportItemModel() { ID = i, Text = DateTimeFormatInfo.CurrentInfo.GetMonthName(i) });
            }
            comboBoxReportConditionMonth.SelectedIndex = 0;

            comboBoxReportConditionDay.DisplayMember = "Text";
            comboBoxReportConditionDay.ValueMember = "ID";
            comboBoxReportConditionDay.Items.Add(new ReportItemModel() { ID = 0, Text = "Day" });
            for (int i = 1; i < 32; i++)
            {
                comboBoxReportConditionDay.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxReportConditionDay.SelectedIndex = 0;

            comboBoxReportConditionHour.DisplayMember = "Text";
            comboBoxReportConditionHour.ValueMember = "ID";
            comboBoxReportConditionHour.Items.Add(new ReportItemModel() { ID = 0, Text = "Hour" });
            for (int i = 0; i < 24; i++)
            {
                comboBoxReportConditionHour.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxReportConditionHour.SelectedIndex = 0;

            comboBoxReportConditionMinute.DisplayMember = "Text";
            comboBoxReportConditionMinute.ValueMember = "ID";
            comboBoxReportConditionMinute.Items.Add(new ReportItemModel() { ID = 0, Text = "Minute" });
            for (int i = 1; i < 60; i++)
            {
                comboBoxReportConditionMinute.Items.Add(new ReportItemModel() { ID = i, Text = i.ToString() });
            }
            comboBoxReportConditionMinute.SelectedIndex = 0;


            panelReportConditionBoolean.Visible = false;
            panelReportConditionDate.Visible = false;
            panelReportConditionEnum.Visible = false;
            panelReportConditionNumber.Visible = false;
            panelReportConditionText.Visible = false;

        }
        public void SortingDown()
        {
            butSortingDown.Enabled = true;

            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            ReportTreeNode reportTreeNodeParent = (ReportTreeNode)reportTreeNode.Parent;

            if (reportTreeNodeParent == null)
                return;

            int CurrentOrder = reportTreeNode.dbSortingField.Ordinal;

            int NextOrder = GetNextSortingOdinalNumber();

            if (CurrentOrder == NextOrder - 1)
            {
                butSortingDown.Enabled = false;
                return;
            }

            foreach (ReportTreeNode RTN in reportTreeNodeParent.Nodes)
            {
                if (RTN.dbSortingField.Ordinal == CurrentOrder + 1)
                {
                    RTN.dbSortingField.Ordinal = CurrentOrder;
                    reportTreeNode.dbSortingField.Ordinal = CurrentOrder + 1;
                    lblSortingOrdinal.Text = reportTreeNode.dbSortingField.Ordinal.ToString();
                    butSortingUp.Enabled = !(reportTreeNode.dbSortingField.Ordinal == (NextOrder - 1));
                    break;
                }
            }
        }
        public void SortingUp()
        {
            butSortingUp.Enabled = true;

            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
                return;

            ReportTreeNode reportTreeNodeParent = (ReportTreeNode)reportTreeNode.Parent;

            if (reportTreeNodeParent == null)
                return;

            int CurrentOrder = reportTreeNode.dbSortingField.Ordinal;

            if (CurrentOrder == 1)
            {
                butSortingUp.Enabled = false;
                return;
            }
            foreach (ReportTreeNode RTN in reportTreeNodeParent.Nodes)
            {
                if (RTN.dbSortingField.Ordinal == CurrentOrder - 1)
                {
                    RTN.dbSortingField.Ordinal = CurrentOrder;
                    reportTreeNode.dbSortingField.Ordinal = CurrentOrder - 1;
                    lblSortingOrdinal.Text = reportTreeNode.dbSortingField.Ordinal.ToString();
                    butSortingUp.Enabled = !(reportTreeNode.dbSortingField.Ordinal == 1);
                    break;
                }
            }
        }
        public void TestBottomRightText()
        {
            RadioButton radioButtonChecked = null;

            if (radioButtonWord.Checked)
            {
                radioButtonChecked = radioButtonWord;
            }
            else if (radioButtonExcel.Checked)
            {
                radioButtonChecked = radioButtonExcel;
            }
            else if (radioButtonKML.Checked)
            {
                radioButtonChecked = radioButtonKML;
            }
            else
            {
                radioButtonChecked = radioButtonCSV;
            }
            radioButtonCSV.Checked = true;

            treeViewCSSPAfterSelect();

            int Take = 0;
            int.TryParse(textBoxFirstXItems.Text, out Take);
            if (Take == 0)
            {
                lblStatusValue.Text = "Please indicate number of items";
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }

            lblStatusValue.Text = "";
            int TVItemID = 0;
            int.TryParse(textBoxStartID.Text, out TVItemID);
            if (TVItemID == 0)
            {
                richTextBoxResults.Text = "Start ID is empty. Please get start id using the Show Web and Get ID buttons.";
                return;
            }

            string retStr = "";
            FileInfo fi = new FileInfo(lblTestTemplateStartFileName.Text.Substring(0, lblTestTemplateStartFileName.Text.IndexOf(".")) + ".txt");

            StreamWriter sw = fi.CreateText();
            sw.Write(richTextBoxExample.Text.Replace("\n", "\r\n"));
            sw.Close();

            if (!fi.Exists)
            {
                lblStatusValue.Text = fi.FullName + " was not created properly";
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }

            retStr = reportBaseService.GenerateReportFromTemplateCSV(fi, TVItemID, Take, 0);
            if (!string.IsNullOrWhiteSpace(retStr))
                lblStatusValue.Text = retStr;

            radioButtonShowResultFiles.Checked = true;

            RefreshTemplateDocuments();

            for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
            {
                if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                {
                    comboBoxTemplateDocuments.SelectedIndex = i;
                }
            }
            radioButtonChecked.Checked = true;
            treeViewCSSPAfterSelect();

            richTextBoxResults.LoadFile(fi.FullName, RichTextBoxStreamType.PlainText);
        }
        public void TestProduceDocument()
        {
            int Take = 0;
            int.TryParse(textBoxFirstXItems.Text, out Take);
            if (Take == 0)
            {
                lblStatusValue.Text = "Please indicate number of items";
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }

            lblStatusValue.Text = "";
            int TVItemID = 0;
            int.TryParse(textBoxStartID.Text, out TVItemID);
            if (TVItemID == 0)
            {
                richTextBoxResults.Text = "Start ID is empty. Please get start id using the Show Web and Get ID buttons.";
                return;
            }

            string retStr = "";
            if (radioButtonCSV.Checked)
            {
                retStr = CreateTemplateAndResultDocumentShell();
                if (!string.IsNullOrWhiteSpace(retStr))
                    return;

                FileInfo fi = new FileInfo(lblTestTemplateStartFileName.Text);

                radioButtonShowTemplateFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                fi = new FileInfo(lblTestTemplateStartFileName.Text.Replace("Template_", ""));

                retStr = reportBaseService.GenerateReportFromTemplateCSV(fi, TVItemID, Take, 0);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                OpenFile();
            }
            else if (radioButtonExcel.Checked)
            {
                retStr = CreateTemplateAndResultDocumentShell();
                if (!string.IsNullOrWhiteSpace(retStr))
                    return;

                FileInfo fi = new FileInfo(lblTestTemplateStartFileName.Text);

                radioButtonShowTemplateFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                fi = new FileInfo(lblTestTemplateStartFileName.Text.Replace("Template_", ""));

                //retStr = reportBaseService.GenerateReportFromTemplateExcel(fi, TVItemID, Take);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                OpenFile();
            }
            else if (radioButtonKML.Checked)
            {
                retStr = CreateTemplateAndResultDocumentShell();
                if (!string.IsNullOrWhiteSpace(retStr))
                    return;

                FileInfo fi = new FileInfo(lblTestTemplateStartFileName.Text);

                radioButtonShowTemplateFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                fi = new FileInfo(lblTestTemplateStartFileName.Text.Replace("Template_", ""));

                retStr = reportBaseService.GenerateReportFromTemplateKML(fi, TVItemID, Take, 0);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                OpenFile();
            }
            else if (radioButtonWord.Checked)
            {
                retStr = CreateTemplateAndResultDocumentShell();
                if (!string.IsNullOrWhiteSpace(retStr))
                    return;

                FileInfo fi = new FileInfo(lblTestTemplateStartFileName.Text);

                radioButtonShowTemplateFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                fi = new FileInfo(lblTestTemplateStartFileName.Text.Replace("Template_", ""));

                retStr = reportBaseService.GenerateReportFromTemplateWord(fi, TVItemID, Take, 0);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                        break;
                    }
                }

                OpenFile();
            }
            else
            {
                retStr = "Unknown file type to create results from template";
            }

        }
        public void TestSelectedTemplate()
        {
            int Take = 0;
            int.TryParse(textBoxFirstXItems.Text, out Take);
            if (Take == 0)
            {
                lblStatusValue.Text = "Please indicate number of items";
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }

            FileInfo fiTemplate = (FileInfo)comboBoxTemplateDocuments.SelectedItem;
            FileInfo fi = new FileInfo(fiTemplate.FullName);
            fi = new FileInfo(fi.FullName.Replace("Template_", ""));

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    lblStatusValue.Text = ex.Message + (ex.InnerException != null ? " - Inner: " + ex.InnerException.Message : "");
                    richTextBoxResults.Text = lblStatusValue.Text;
                    return;
                }
            }

            try
            {
                File.Copy(((FileInfo)comboBoxTemplateDocuments.SelectedItem).FullName, fi.FullName);
            }
            catch (Exception ex)
            {
                lblStatusValue.Text = ex.Message + (ex.InnerException != null ? " - Inner: " + ex.InnerException.Message : "");
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }
            fi = new FileInfo(fi.FullName);
            if (!fi.Exists)
            {
                lblStatusValue.Text = "File [" + fi.FullName + "] could not be created.";
                richTextBoxResults.Text = lblStatusValue.Text;
                return;
            }

            lblStatusValue.Text = "";
            int TVItemID = 0;
            int.TryParse(textBoxStartID.Text, out TVItemID);
            if (TVItemID == 0)
            {
                richTextBoxResults.Text = "Start ID in empty. Please get start id using the Show Web and Get ID buttons.";
                return;
            }

            string retStr = "";
            if (radioButtonCSV.Checked)
            {
                retStr = reportBaseService.GenerateReportFromTemplateCSV(fi, TVItemID, Take, 0);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                    }
                }

                OpenFile();
            }
            else if (radioButtonWord.Checked)
            {
                string DocumentParsingError = "The document has parsing error(s). Please check below.";
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(fi.FullName);
                Microsoft.Office.Interop.Word.Range range = doc.Range();

                range.End = 0;

                wordApp.Visible = true;

                int UnderTVItemID = 0;
                int.TryParse(textBoxStartID.Text, out UnderTVItemID);
                if (UnderTVItemID == 0)
                {
                    string errStr = "Start ID in empty. Please get start id using the Show Web and Get ID buttons.";
                    range.Comments.Add(range, errStr);
                    richTextBoxResults.Text = "Start ID in empty. Please get start id using the Show Web and Get ID buttons.";
                    return;
                }

                List<ReportTag> reportTagList = new List<ReportTag>();
                ReportTag reportTagStart = new ReportTag()
                {
                    wordApp = wordApp,
                    doc = doc,
                    UnderTVItemID = UnderTVItemID,
                    OnlyImmediateChildren = true,
                    Take = Take,
                    AppTaskID = 0
                };

                reportTagStart.OnlyImmediateChildren = false;
                retStr = reportBaseService.CheckTagsAndContentOKWord(reportTagStart, reportTagList);
                if (string.IsNullOrWhiteSpace(retStr))
                {
                    retStr = reportBaseService.FillTemplateWithDBInfoWord(reportTagStart);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        range.Comments.Add(range, DocumentParsingError);
                        lblStatusValue.Text = retStr;
                        richTextBoxResults.Text = lblStatusValue.Text;
                    }
                }

                range = doc.Range();
                range.Start = 0;
                range.End = 0;
                range.Select();

                doc.SaveAs2(fi.FullName);
                //doc.Close();
                //wordApp.Quit();

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                    }
                }

                //OpenFile();

            }
            else if (radioButtonExcel.Checked)
            {
                string ExcelDocumentParsingError = "The excel document has parsing error(s). Please check below.";
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Open(fi.FullName);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = workBook.Sheets[1];

                excelApp.Visible = true;

                int UnderTVItemID = 0;
                int.TryParse(textBoxStartID.Text, out UnderTVItemID);
                if (UnderTVItemID == 0)
                {
                    string errStr = "Start ID in empty. Please get start id using the Show Web and Get ID buttons.";
                    //Microsoft.Office.Interop.Excel.Range range = workSheet.Cells[1, 1];
                    workSheet.Cells[1, 1].AddComment(errStr);
                    richTextBoxResults.Text = "Start ID in empty. Please get start id using the Show Web and Get ID buttons.";
                    return;
                }

                List<ReportTag> reportTagList = new List<ReportTag>();
                ReportTag reportTagStart = new ReportTag()
                {
                    excelApp = excelApp,
                    workbook = workBook,
                    UnderTVItemID = UnderTVItemID,
                    OnlyImmediateChildren = true,
                    Take = Take,
                };

                retStr = ""; // reportBaseService.FillTemplateWithDBInfoExcel(reportTagStart);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    workSheet.Cells[1, 1].AddComment(ExcelDocumentParsingError);
                    lblStatusValue.Text = retStr;
                    richTextBoxResults.Text = lblStatusValue.Text;
                }

                workBook.Save();

            }
            else if (radioButtonKML.Checked)
            {
                retStr = reportBaseService.GenerateReportFromTemplateKML(fi, TVItemID, Take, 0);
                if (!string.IsNullOrWhiteSpace(retStr))
                    lblStatusValue.Text = retStr;

                radioButtonShowResultFiles.Checked = true;

                RefreshTemplateDocuments();

                for (int i = 0, count = comboBoxTemplateDocuments.Items.Count; i < count; i++)
                {
                    if (((FileInfo)comboBoxTemplateDocuments.Items[i]).FullName == fi.FullName)
                    {
                        comboBoxTemplateDocuments.SelectedIndex = i;
                    }
                }

                OpenFile();
            }
        }
        public void treeViewCSSPAfterCheck(TreeViewEventArgs e)
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)e.Node;

            if (!reportTreeNode.Checked)
            {
                reportTreeNode.dbSortingField = new ReportSortingField() { Ordinal = 0, ReportSorting = ReportSortingEnum.Error };
                reportTreeNode.reportFormatingField = new ReportFormatingField() { ReportFormatingDate = ReportFormatingDateEnum.Error, ReportFormatingNumber = ReportFormatingNumberEnum.Error };
                reportTreeNode.reportConditionDateFieldList = new List<ReportConditionDateField>();
                reportTreeNode.reportConditionNumberFieldList = new List<ReportConditionNumberField>();
                reportTreeNode.reportConditionTextFieldList = new List<ReportConditionTextField>();
            }
        }
        public void treeViewCSSPAfterSelect()
        {
            ReportTreeNode reportTreeNode = (ReportTreeNode)treeViewCSSP.SelectedNode;

            if (reportTreeNode == null)
            {
                lblEmptyPanelMessage.Text = "Please select an item.";
                return;
            }

            if (reportTreeNode.ReportTreeNodeSubType == ReportTreeNodeSubTypeEnum.TableSelectable)
            {
                richTextBoxExample.Text = "";
                richTextBoxResults.Text = "";

                if (reportBaseService.ReportFileType == ReportFileTypeEnum.CSV)
                {
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sbFirstLine = new StringBuilder();
                    reportBaseService.GetTreeViewSelectedStatusCSV(reportTreeNode, sb, sbFirstLine, 0);
                    string FirstLine = sbFirstLine.ToString();
                    if (FirstLine.Length > 0)
                    {
                        FirstLine = FirstLine.Substring(0, FirstLine.Length - 1);
                    }
                    richTextBoxExample.AppendText(FirstLine + "\r\n");
                    richTextBoxExample.AppendText(sb + "\r\n");
                }
                else if (reportBaseService.ReportFileType == ReportFileTypeEnum.Word)
                {
                    // nothing
                }
                else if (reportBaseService.ReportFileType == ReportFileTypeEnum.Excel)
                {
                    // nothing
                }
                else if (reportBaseService.ReportFileType == ReportFileTypeEnum.KML)
                {
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sbTemp = new StringBuilder();

                    sbTemp = new StringBuilder("");
                    sbTemp.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
                    sbTemp.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
                    sbTemp.AppendLine(@"<Folder>");
                    string TemplateType = "Mesh";
                    sbTemp.AppendLine("\t<name>Template_" + TemplateType + "</name>");
                    sbTemp.AppendLine("\t<description>");

                    richTextBoxExample.AppendText(sbTemp.ToString());

                    reportBaseService.GetTreeViewSelectedStatusKML(reportTreeNode, sb, 0);
                    richTextBoxExample.AppendText(sb.ToString());

                    sbTemp = new StringBuilder();
                    sbTemp.AppendLine("\t</description>");
                    sbTemp.AppendLine(@"</Folder>");
                    sbTemp.AppendLine(@"</kml>");

                    richTextBoxExample.AppendText(sbTemp.ToString());

                }
                else
                {
                    richTextBoxExample.AppendText("Error: Report File Type [" + reportBaseService.ReportFileType.ToString() + "] not supported\r\n");
                    return;
                }

                butShowExpectedResult.Enabled = true;
                butProduceTestDocument.Enabled = true;
            }
            else
            {
                butShowExpectedResult.Enabled = false;
                butProduceTestDocument.Enabled = false;

                if (!richTextBoxExample.Text.StartsWith(ChangeText))
                {
                    richTextBoxExample.Text = ChangeText + richTextBoxExample.Text;
                }
                richTextBoxResults.Text = "Please select a table (green) to view information.";
                richTextBoxResults.Find("green");
                richTextBoxResults.SelectionColor = Color.Green;
            }

            lblSelectedTreeViewText.Text = reportTreeNode.Text +  (reportTreeNode.ReportFieldType != ReportFieldTypeEnum.Error ? " (" + reportTreeNode.ReportFieldType + ")" : "");
            panelSorting.Visible = false;
            lblFormatDatabaseFiltering.Visible = false;
            lblFormatReportCondition.Visible = false;
            comboBoxDBFormatingNumber.Visible = false;
            comboBoxReportFormatingNumber.Visible = false;

            switch (reportTreeNode.ReportTreeNodeSubType)
            {
                case ReportTreeNodeSubTypeEnum.Error:
                    {
                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        lblEmptyPanelMessage.Text = "An error happened in treeViewCSS_AfterSelect";
                    }
                    break;
                case ReportTreeNodeSubTypeEnum.TableSelectable:
                    {
                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        lblEmptyPanelMessage.Text = "Please select/check one of the field below to view it in the report.";
                    }
                    break;
                case ReportTreeNodeSubTypeEnum.TableNotSelectable:
                    {
                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        lblEmptyPanelMessage.Text = "Please select/check one of the field below to view it in the report.";
                    }
                    break;
                case ReportTreeNodeSubTypeEnum.Field:
                    {
                        if (!reportTreeNode.Checked)
                        {
                            SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                            lblEmptyPanelMessage.Text = "Field has to be checked in order to do filtering.";
                            return;
                        }

                        if (reportTreeNode.Text.Contains("_Counter") && reportTreeNode.ReportFieldType == ReportFieldTypeEnum.NumberWhole)
                        {
                            panelSorting.Visible = false;
                            lblEmptyPanelMessage.Text = "_Counter fields are not sortable nor filterable.";
                            SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        }
                        else
                        {

                            panelSorting.Visible = true;
                            lblFormatDatabaseFiltering.Visible = false;
                            lblFormatReportCondition.Visible = false;
                            comboBoxDBFormatingNumber.Visible = false;
                            comboBoxReportFormatingNumber.Visible = false;

                            switch (reportTreeNode.ReportFieldType)
                            {
                                case ReportFieldTypeEnum.Error:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                                    }
                                    break;
                                case ReportFieldTypeEnum.DateAndTime:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleDate);
                                    }
                                    break;
                                case ReportFieldTypeEnum.NumberWhole:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleNumber);
                                    }
                                    break;
                                case ReportFieldTypeEnum.NumberWithDecimal:
                                    {
                                        lblFormatDatabaseFiltering.Visible = true;
                                        lblFormatReportCondition.Visible = true;
                                        comboBoxDBFormatingNumber.Visible = true;
                                        comboBoxReportFormatingNumber.Visible = true;
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleNumber);
                                    }
                                    break;
                                case ReportFieldTypeEnum.Text:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleText);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TrueOrFalse:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleBoolean);
                                    }
                                    break;
                                case ReportFieldTypeEnum.FilePurpose:
                                    {
                                        FillEnumListBoxes<FilePurposeEnum>(typeof(FilePurposeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.FileType:
                                    {
                                        FillEnumListBoxes<FileTypeEnum>(typeof(FileTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TranslationStatus:
                                    {
                                        FillEnumListBoxes<TranslationStatusEnum>(typeof(TranslationStatusEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.BoxModelResultType:
                                    {
                                        FillEnumListBoxes<BoxModelResultTypeEnum>(typeof(BoxModelResultTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.InfrastructureType:
                                    {
                                        FillEnumListBoxes<InfrastructureTypeEnum>(typeof(InfrastructureTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.FacilityType:
                                    {
                                        FillEnumListBoxes<FacilityTypeEnum>(typeof(FacilityTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.AerationType:
                                    {
                                        FillEnumListBoxes<AerationTypeEnum>(typeof(AerationTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.PreliminaryTreatmentType:
                                    {
                                        FillEnumListBoxes<PreliminaryTreatmentTypeEnum>(typeof(PreliminaryTreatmentTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.PrimaryTreatmentType:
                                    {
                                        FillEnumListBoxes<PrimaryTreatmentTypeEnum>(typeof(PrimaryTreatmentTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SecondaryTreatmentType:
                                    {
                                        FillEnumListBoxes<SecondaryTreatmentTypeEnum>(typeof(SecondaryTreatmentTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TertiaryTreatmentType:
                                    {
                                        FillEnumListBoxes<TertiaryTreatmentTypeEnum>(typeof(TertiaryTreatmentTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TreatmentType:
                                    {
                                        FillEnumListBoxes<TreatmentTypeEnum>(typeof(TreatmentTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.DisinfectionType:
                                    {
                                        FillEnumListBoxes<DisinfectionTypeEnum>(typeof(DisinfectionTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.CollectionSystemType:
                                    {
                                        FillEnumListBoxes<CollectionSystemTypeEnum>(typeof(CollectionSystemTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.AlarmSystemType:
                                    {
                                        FillEnumListBoxes<AlarmSystemTypeEnum>(typeof(AlarmSystemTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.ScenarioStatus:
                                    {
                                        FillEnumListBoxes<ScenarioStatusEnum>(typeof(ScenarioStatusEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.StorageDataType:
                                    {
                                        FillEnumListBoxes<StorageDataTypeEnum>(typeof(StorageDataTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.Language:
                                    {
                                        FillEnumListBoxes<LanguageEnum>(typeof(LanguageEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SampleType:
                                    {
                                        FillEnumListBoxes<SampleTypeEnum>(typeof(SampleTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.BeaufortScale:
                                    {
                                        FillEnumListBoxes<BeaufortScaleEnum>(typeof(BeaufortScaleEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.AnalyzeMethod:
                                    {
                                        FillEnumListBoxes<AnalyzeMethodEnum>(typeof(AnalyzeMethodEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SampleMatrix:
                                    {
                                        FillEnumListBoxes<SampleMatrixEnum>(typeof(SampleMatrixEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.Laboratory:
                                    {
                                        FillEnumListBoxes<LaboratoryEnum>(typeof(LaboratoryEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SampleStatus:
                                    {
                                        FillEnumListBoxes<SampleStatusEnum>(typeof(SampleStatusEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SamplingPlanType:
                                    {
                                        FillEnumListBoxes<SamplingPlanTypeEnum>(typeof(SamplingPlanTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.LabSheetSampleType:
                                    {
                                        FillEnumListBoxes<SampleTypeEnum>(typeof(SampleTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.LabSheetType:
                                    {
                                        FillEnumListBoxes<LabSheetTypeEnum>(typeof(LabSheetTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.LabSheetStatus:
                                    {
                                        FillEnumListBoxes<LabSheetStatusEnum>(typeof(LabSheetStatusEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.PolSourceInactiveReason:
                                    {
                                        FillEnumListBoxes<PolSourceInactiveReasonEnum>(typeof(PolSourceInactiveReasonEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.PolSourceObsInfo:
                                    {
                                        FillEnumListBoxes<PolSourceObsInfoEnum>(typeof(PolSourceObsInfoEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.AddressType:
                                    {
                                        FillEnumListBoxes<AddressTypeEnum>(typeof(AddressTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.StreetType:
                                    {
                                        FillEnumListBoxes<StreetTypeEnum>(typeof(StreetTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.ContactTitle:
                                    {
                                        FillEnumListBoxes<ContactTitleEnum>(typeof(ContactTitleEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.EmailType:
                                    {
                                        FillEnumListBoxes<EmailTypeEnum>(typeof(EmailTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TelType:
                                    {
                                        FillEnumListBoxes<TelTypeEnum>(typeof(TelTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TideText:
                                    {
                                        FillEnumListBoxes<TideTextEnum>(typeof(TideTextEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.TideDataType:
                                    {
                                        FillEnumListBoxes<TideDataTypeEnum>(typeof(TideDataTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.SpecialTableType:
                                    {
                                        FillEnumListBoxes<SpecialTableTypeEnum>(typeof(SpecialTableTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.MWQMSiteLatestClassification:
                                    {
                                        FillEnumListBoxes<MWQMSiteLatestClassificationEnum>(typeof(MWQMSiteLatestClassificationEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.PolSourceIssueRisk:
                                    {
                                        FillEnumListBoxes<PolSourceIssueRiskEnum>(typeof(PolSourceIssueRiskEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                case ReportFieldTypeEnum.MikeScenarioSpecialResultKMLType:
                                    {
                                        FillEnumListBoxes<MikeScenarioSpecialResultKMLTypeEnum>(typeof(MikeScenarioSpecialResultKMLTypeEnum));
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEnum);
                                    }
                                    break;
                                default:
                                    {
                                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                                    }
                                    break;
                            }

                            SetItemValues(reportTreeNode);
                        }
                    }
                    break;
                case ReportTreeNodeSubTypeEnum.FieldsHolder:
                    {
                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        lblEmptyPanelMessage.Text = "Please select/check one of the field below to view it in the report.";
                    }
                    break;
                default:
                    {
                        SetPanelRightTopMiddleVisible(panelRightTopMiddleEmpty);
                        lblEmptyPanelMessage.Text = "Please select/check one of the field below to view it in the report.";
                    }
                    break;
            }
        }
        private void FillEnumListBoxes<T>(Type EnumType)
        {
            listBoxDBFilteringEnum.Items.Clear();
            listBoxReportConditionEnum.Items.Clear();
            if (EnumType.FullName.Contains("SampleTypeEnum"))
            {
                foreach (string name in Enum.GetNames(EnumType).Where(c => c != "Error").OrderBy(c => c))
                {
                    listBoxDBFilteringEnum.Items.Add(name);
                    listBoxReportConditionEnum.Items.Add(name);
                }
                //for (int i = 101, count = Enum.GetNames(EnumType).Count() + 100; i < count; i++)
                //{
                //    listBoxDBFilteringEnum.Items.Add(Enum.GetName(EnumType, i));
                //    listBoxReportConditionEnum.Items.Add(Enum.GetName(EnumType, i));
                //}
            }
            else if (EnumType.FullName.Contains("BeaufortScaleEnum"))
            {
                for (int i = 0, count = Enum.GetNames(EnumType).Count() - 1; i < count; i++)
                {
                    listBoxDBFilteringEnum.Items.Add(Enum.GetName(EnumType, i));
                    listBoxReportConditionEnum.Items.Add(Enum.GetName(EnumType, i));
                }
            }
            else if (EnumType.FullName.Contains("PolSourceObsInfoEnum"))
            {
                foreach (string name in Enum.GetNames(EnumType).Where(c => !c.EndsWith("Start")).OrderBy(c => c))
                {
                    listBoxDBFilteringEnum.Items.Add(name);
                    listBoxReportConditionEnum.Items.Add(name);
                }
            }
            else
            {
                foreach (string name in Enum.GetNames(EnumType).Where(c => c != "Error").OrderBy(c => c))
                {
                    listBoxDBFilteringEnum.Items.Add(name);
                    listBoxReportConditionEnum.Items.Add(name);
                }

                //for (int i = 1, count = Enum.GetNames(EnumType).Count(); i < count; i++)
                //{
                //    listBoxDBFilteringEnum.Items.Add(Enum.GetName(EnumType, i));
                //    listBoxReportConditionEnum.Items.Add(Enum.GetName(EnumType, i));
                //}
            }
        }
        public void WebView()
        {
            panelMiddleWebBrowser.Visible = !WebIsVisible;
            panelMiddleTreeViewAndTest.Visible = WebIsVisible;
            WebIsVisible = !WebIsVisible;
            butWeb.Text = (WebIsVisible ? "Hide Web" : "Show Web");
        }






        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}
